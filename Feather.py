import sys
import os
import shutil
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                             QWidget, QLineEdit, QProgressBar, QMessageBox,
                             QFileDialog, QDialog, QLabel, QMenuBar, QAction)
from PyQt5.QtCore import QSettings, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from PIL import Image
import logging
import pygetwindow as gw
import gc

# Set up logging
logging.basicConfig(
    filename='app.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s:%(message)s'
)

settings = QSettings("RonsinPhotocopy", "Feather")

# Thread-safe progress tracking
class ProgressTracker:
    def __init__(self, total_files):
        self._lock = threading.Lock()
        self._processed = 0
        self._errors = 0
        self._total = total_files
    
    def increment_processed(self):
        with self._lock:
            self._processed += 1
            return self._processed
    
    def increment_errors(self):
        with self._lock:
            self._errors += 1
            return self._errors
    
    def get_progress(self):
        with self._lock:
            return self._processed, self._errors, self._total

def process_image_batch(file_batch, target_size, progress_tracker, progress_callback):
    """Process a batch of images to optimize memory usage"""
    results = []
    
    for file_path in file_batch:
        try:
            result = process_single_image(file_path, target_size)
            results.append((file_path, result))
            
            if result.startswith('error:'):
                progress_tracker.increment_errors()
            else:
                progress_tracker.increment_processed()
                
        except Exception as e:
            error_msg = f'error: {str(e)}'
            results.append((file_path, error_msg))
            progress_tracker.increment_errors()
            logging.error(f"Batch processing error for {file_path}: {str(e)}")
        
        # Update progress after each file
        processed, errors, total = progress_tracker.get_progress()
        progress_value = int(((processed + errors) / total) * 100)
        progress_callback.emit(progress_value)
        
        # Force garbage collection after each image to free memory
        gc.collect()
    
    return results

def process_single_image(file_path, target_size):
    """Process a single image file"""
    logging.info(f"Processing {file_path}")

    try:
        if not file_path.lower().endswith(('.jpeg', '.jpg', '.png', '.tiff', '.tif')):
            return 'skipped: unsupported format'
            
        # Use context manager to ensure proper resource cleanup
        with Image.open(file_path) as img:
            # Preserve original image properties
            original_mode = img.mode
            original_format = img.format
            original_info = img.info.copy()  # Preserve metadata like DPI, compression, etc.
            
            # Get original dimensions
            original_width, original_height = img.size
            target_width, target_height = target_size

            # Determine orientation and adjust target size accordingly
            original_is_landscape = original_width > original_height
            
            # If original is landscape but target is portrait, swap target dimensions
            if original_is_landscape and target_width < target_height:
                target_width, target_height = target_height, target_width
            # If original is portrait but target is landscape, swap target dimensions  
            elif not original_is_landscape and target_width > target_height:
                target_width, target_height = target_height, target_width

            # Calculate scaling to fit within target dimensions while maintaining aspect ratio
            scale_width = target_width / original_width
            scale_height = target_height / original_height
            scale = min(scale_width, scale_height)  # Use smaller scale to ensure it fits

            # Calculate new dimensions after scaling
            new_width = int(original_width * scale)
            new_height = int(original_height * scale)

            # Scale the image while preserving mode
            scaled_img = img.resize((new_width, new_height), Image.LANCZOS)

            # Create background with target dimensions in the same mode as original
            # Handle different color modes appropriately
            if original_mode == '1':  # 1-bit (black and white)
                background_color = 1  # White in 1-bit mode
            elif original_mode == 'L':  # 8-bit grayscale
                background_color = 255  # White in grayscale
            elif original_mode == 'P':  # Palette mode
                background_color = 255  # Assume white is palette index 255, or use img.getpalette()
            elif original_mode in ['RGB', 'RGBA']:
                background_color = 'white'
            elif original_mode == 'CMYK':
                background_color = (0, 0, 0, 0)  # White in CMYK
            else:
                background_color = 'white'  # Default fallback
                
            final_img = Image.new(original_mode, (target_width, target_height), background_color)

            # Copy palette if original image had one
            if original_mode == 'P' and img.getpalette():
                final_img.putpalette(img.getpalette())

            # Calculate position to center the scaled image
            x_offset = (target_width - new_width) // 2
            y_offset = (target_height - new_height) // 2

            # Paste the scaled image onto the background
            if original_mode == 'RGBA' or 'transparency' in original_info:
                # Handle transparency properly
                final_img.paste(scaled_img, (x_offset, y_offset), scaled_img if original_mode == 'RGBA' else None)
            else:
                final_img.paste(scaled_img, (x_offset, y_offset))

            # Prepare save parameters, preserving original format and compression
            save_kwargs = {'dpi': (200, 200)}
            
            # Preserve format-specific settings
            if original_format == 'JPEG':
                save_kwargs['quality'] = original_info.get('quality', 70)
                save_kwargs['optimize'] = True
                final_img.save(file_path, 'JPEG', **save_kwargs)
            elif original_format == 'TIFF':
                # Preserve TIFF compression settings
                if 'compression' in original_info:
                    save_kwargs['compression'] = original_info['compression']
                # Preserve other TIFF-specific info
                for key in ['description', 'software', 'datetime']:
                    if key in original_info:
                        save_kwargs[key] = original_info[key]
                final_img.save(file_path, 'TIFF', **save_kwargs)
            elif original_format == 'PNG':
                # Preserve PNG settings
                if 'transparency' in original_info:
                    save_kwargs['transparency'] = original_info['transparency']
                if original_mode == 'P':
                    save_kwargs['optimize'] = True
                final_img.save(file_path, 'PNG', **save_kwargs)
            else:
                # Default save with original format
                final_img.save(file_path, original_format, **save_kwargs)
            
            # Explicitly delete image objects to free memory immediately
            del scaled_img
            del final_img

        return 'success'
        
    except Exception as e:
        logging.error(f"Error processing {file_path}: {str(e)}")
        return f'error: {str(e)}'

class ImageProcessor(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, file_paths, target_size_pixels, max_workers=None, batch_size=10):
        super().__init__()
        self.file_paths = file_paths
        self.target_size_pixels = target_size_pixels
        
        # Determine optimal number of worker threads
        # Use min of (CPU cores, 8) to avoid overwhelming the system
        import multiprocessing
        if max_workers is None:
            self.max_workers = min(multiprocessing.cpu_count(), 8)
        else:
            self.max_workers = max_workers
            
        # Batch size for memory optimization
        self.batch_size = batch_size
        
        logging.info(f"ImageProcessor initialized with {self.max_workers} workers and batch size {self.batch_size}")

    def run(self):
        total_files = len(self.file_paths)
        if total_files == 0:
            self.finished.emit()
            return

        # Initialize thread-safe progress tracker
        progress_tracker = ProgressTracker(total_files)
        
        # Split files into batches for memory optimization
        file_batches = [
            self.file_paths[i:i + self.batch_size] 
            for i in range(0, len(self.file_paths), self.batch_size)
        ]
        
        logging.info(f"Processing {total_files} files in {len(file_batches)} batches")

        try:
            # Use ThreadPoolExecutor for parallel processing
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # Submit all batch tasks
                future_to_batch = {
                    executor.submit(
                        process_image_batch, 
                        batch, 
                        self.target_size_pixels, 
                        progress_tracker,
                        self.progress
                    ): batch 
                    for batch in file_batches
                }

                # Wait for all batches to complete
                for future in as_completed(future_to_batch):
                    batch = future_to_batch[future]
                    try:
                        batch_results = future.result()
                        
                        # Log results for this batch
                        for file_path, result in batch_results:
                            logging.info(f"Processed {file_path}: {result}")
                            
                    except Exception as e:
                        logging.error(f"Batch processing exception: {str(e)}")
                        # Still increment error count for all files in failed batch
                        for _ in batch:
                            progress_tracker.increment_errors()

        except Exception as e:
            logging.error(f"ThreadPoolExecutor exception: {str(e)}")
        
        finally:
            # Ensure final progress update
            processed, errors, total = progress_tracker.get_progress()
            final_progress = int(((processed + errors) / total) * 100) if total > 0 else 100
            self.progress.emit(final_progress)
            
            # Force final garbage collection
            gc.collect()
            
            logging.info(f"Processing completed: {processed} successful, {errors} errors out of {total} total")
            self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.loadSettings()
        self.force_to_front()
    
    def force_to_front(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def initUI(self):
        self.setWindowTitle("Feather - A Lightweight Image Optimizer")
        self.setFixedSize(400, 300)
        self.setWindowIcon(QIcon('feather.ico'))

        layout = QVBoxLayout()
        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText("Enter directory path here...")
        layout.addWidget(self.input_path)

        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse)
        layout.addWidget(self.browse_button)

        self.start_button = QPushButton("Start Processing")
        self.start_button.clicked.connect(self.start_processing)
        layout.addWidget(self.start_button)

        self.progress_total = QProgressBar(self)
        self.progress_total.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress_total)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.menuBar = self.menuBar()
        extrasMenu = self.menuBar.addMenu('Extras')
        self.toggleThemeAction = QAction('Toggle Theme', self)
        self.toggleThemeAction.triggered.connect(self.toggle_theme)
        extrasMenu.addAction(self.toggleThemeAction)

        self.summonWincopyAction = QAction('Summon Wincopy After Processing Completed', self)
        self.summonWincopyAction.setCheckable(True)
        self.summonWincopyAction.triggered.connect(lambda: settings.setValue("summonWincopy", self.summonWincopyAction.isChecked()))
        extrasMenu.addAction(self.summonWincopyAction)

        self.closeFeatherAction = QAction('Close Feather After Processing', self)
        self.closeFeatherAction.setCheckable(True)
        self.closeFeatherAction.triggered.connect(self.close_feather_after_processing)
        extrasMenu.addAction(self.closeFeatherAction)

        aboutAction = QAction('About', self)
        aboutAction.triggered.connect(self.show_about_dialog)
        extrasMenu.addAction(aboutAction)

    def loadSettings(self):
        self.dark_mode = settings.value("darkMode", True, type=bool)
        self.summonWincopyAction.setChecked(settings.value("summonWincopy", False, type=bool))
        self.closeFeatherAction.setChecked(settings.value("closeAfterProcessing", False, type=bool))
        self.apply_theme()

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        settings.setValue("darkMode", self.dark_mode)
        self.apply_theme()

    def apply_theme(self):
        if self.dark_mode:
            self.setStyleSheet(""" 
                QMainWindow, QDialog {
                    background-color: #333;
                    color: white;
                }
                QLineEdit, QPushButton, QProgressBar, QMenuBar, QMenu {
                    background-color: #555;
                    color: white;
                }
                QLabel {
                    color: white;
                }
                QProgressBar {
                    border: 1px solid #666;
                    background-color: #333;
                }
                QProgressBar::chunk {
                    background-color: #06b;
                }
                QMenuBar::item:selected {
                    background-color: #06b;
                }
                QMenu::item:selected {
                    background-color: #333;
                }
            """)
        else:
            self.setStyleSheet(""" 
                QMainWindow, QDialog {
                    background-color: #eee;
                    color: black;
                }
                QLineEdit, QPushButton, QProgressBar, QMenuBar, QMenu {
                    background-color: #ccc;
                    color: black;
                }
                QLabel {
                    color: black;
                }
                QProgressBar {
                    border: 1px solid #bbb;
                    background-color: #eee;
                }
                QProgressBar::chunk {
                    background-color: #06b;
                }
                QMenuBar::item:selected {
                    background-color: #a0c4ff;
                }
                QMenu::item:selected {
                    background-color: #a0c4ff;
                }
            """)

    def browse(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Folder")
        if directory:
            self.input_path.setText(directory)

    def start_processing(self):
        directory_path = self.input_path.text()
        dpi = 200
        inches_to_pixels = lambda inches: int(inches * dpi)
        target_size_pixels = (inches_to_pixels(8.5), inches_to_pixels(11))
        
        self.progress_total.setValue(0)

        # Gather file paths for processing
        file_paths = [os.path.join(dp, f) for dp, _, filenames in os.walk(directory_path)
                      for f in filenames if f.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.tif'))]

        if not file_paths:
            QMessageBox.warning(self, "No Files Found", "No supported image files found in the selected directory.")
            return

        self.processor = ImageProcessor(file_paths, target_size_pixels)
        self.processor.progress.connect(self.progress_total.setValue)
        self.processor.finished.connect(self.processing_finished)
        self.processor.start()

    def processing_finished(self):
        QMessageBox.information(self, "Feather is Finished", "All images have been processed.", QMessageBox.Ok)
        if self.summonWincopyAction.isChecked():
            self.summon_wincopy()
        if self.closeFeatherAction.isChecked():
            self.close()

    def summon_wincopy(self):
        windows = gw.getWindowsWithTitle('Photocopy Orders: 1 - Cloud')
        if windows:
            window = windows[0]
            if window.isMinimized or not window.visible:
                window.restore()
            window.activate()

    def close_feather_after_processing(self):
        if self.closeFeatherAction.isChecked():
            settings.setValue("closeAfterProcessing", True)
        else:
            settings.setValue("closeAfterProcessing", False)

    def show_about_dialog(self):
        about_dialog = QDialog(self)
        about_dialog.setWindowTitle("About Feather - A Lightweight Image Optimizer")
        about_dialog.setFixedSize(300, 200)
        about_dialog_layout = QVBoxLayout()
        label = QLabel("Feather - A Lightweight Image Optimizer\n\nVersion 3.0 - July 2, 2025\n© Ronsin Photocopy\nAll rights reserved\n\nUse of this app is exclusive to Ronsin Photocopy\n\nThereby, unlimited copys of this software\nare granted in purpituity\n\nApplication Developer: Nathan Curtis")
        label.setAlignment(Qt.AlignCenter)
        about_dialog_layout.addWidget(label)
        about_dialog.setLayout(about_dialog_layout)
        about_dialog.setStyleSheet(self.styleSheet())
        about_dialog.exec_()

    def closeEvent(self, event):
        event.accept()

if __name__ == '__main__':
    logging.info("Starting Feather application.")
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.exit(app.exec_())