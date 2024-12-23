import sys
import os
import shutil
import tempfile
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                             QWidget, QLineEdit, QProgressBar, QMessageBox,
                             QFileDialog, QDialog, QLabel, QMenuBar, QAction)
from PyQt5.QtCore import QSettings, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from PIL import Image
import logging
import pygetwindow as gw

# Set up logging
logging.basicConfig(
    filename='app.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s:%(message)s'
)

settings = QSettings("RonsinPhotocopy", "Feather")

def process_image(data):
    file_path, target_size = data
    logging.info(f"Processing {file_path}")

    try:
        if file_path.lower().endswith(('.jpeg', '.jpg')):
            with Image.open(file_path) as img:
                img = img.convert('RGB')

                original_dpi = img.info.get('dpi', (300, 300))[0]
                original_size_inches = (img.width / original_dpi, img.height / original_dpi)
                target_size_pixels = (int(original_size_inches[0] * 300), int(original_size_inches[1] * 300))

                img = img.resize(target_size_pixels, Image.LANCZOS)
                new_img = Image.new('RGB', target_size, 'white')

                x = (target_size[0] - img.width) // 2
                y = (target_size[1] - img.height) // 2
                new_img.paste(img, (x, y))

                new_img.save(file_path, 'JPEG', quality=70, dpi=(300, 300))

        return 'success'
    except Exception as e:
        logging.error(f"Error processing {file_path}: {str(e)}")
        return f'error: {str(e)}'

class ImageProcessor(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, file_paths, target_size_pixels):
        super().__init__()
        self.file_paths = file_paths
        self.target_size_pixels = target_size_pixels

    def run(self):
        total_files = len(self.file_paths)
        progress_step = 100 / total_files if total_files > 0 else 100

        for i, file_path in enumerate(self.file_paths):
            result = process_image((file_path, self.target_size_pixels))
            logging.info(f"Processed {file_path}: {result}")
            progress_value = int((i + 1) * progress_step)
            self.progress.emit(progress_value)

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
        dpi = 300
        inches_to_pixels = lambda inches: int(inches * dpi)
        target_size_pixels = (inches_to_pixels(8.5), inches_to_pixels(11))
        
        self.progress_total.setValue(0)

        # Gather file paths for processing
        file_paths = [os.path.join(dp, f) for dp, _, filenames in os.walk(directory_path)
                      for f in filenames if f.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.tif'))]

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
        label = QLabel("Feather - A Lightweight Image Optimizer\n\nVersion 2.5 - October 26th, 2024\n© Ronsin Photocopy\nAll rights reserved\n\nUse of this app is exclusive to Ronsin Photocopy\n\nThereby, unlimited copys of this software\nare granted in purpituity\n\nApplication Developer: Nathan Curtis")
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
