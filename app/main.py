# src/main.py (modified)
import sys

from PyQt6.QtGui import QFont

from PyQt6.QtWidgets import QApplication

from controllers.app_controller import AppController

from views.windows.main_window import MainWindow
from views.windows.about_window import AboutWindow
from views.windows.process_window import ProcessWindow

from resources.QSS import QSS

__version__ = "1.0.0"
__author__ = "Piero Olivas"
__credits__ = ['Piero Olivas']
__maintainer__ = "Piero Olivas"
__email__ = "psot14022001@gmail.com"

def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)
    app.setFont(QFont("Inter", 8))  # or "Segoe UI", "Roboto"

    main_window = MainWindow()
    about_window = AboutWindow()
    process_window = ProcessWindow()

    app_controller = AppController(main_window, about_window, process_window)
 
    main_window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()