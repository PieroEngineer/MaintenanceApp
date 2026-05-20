from views.windows.main_window import MainWindow
from views.windows.about_window import AboutWindow
from views.windows.process_window import ProcessWindow

from models.path_model import PathModel

from controllers.path_controller import PathController
from controllers.process_controller import ProcessController

class AppController:
    def __init__(self, main_window: MainWindow, about_window: AboutWindow, process_window: ProcessWindow):

        # Views
        self.main_view = main_window
        self.about_view = about_window
        self.process_view = process_window

        # Models
        self.path_model = PathModel()

        # Controllers
        self.path_controller = PathController(self.path_model, self.main_view, self)
        self.process_controller = ProcessController(self.path_model, self.main_view, self.process_view, self) 

        # Navigation
        self.main_view.info_btn.clicked.connect(self.on_info_btn_clicked)
        self.about_view.accept_btn.clicked.connect(self.about_view.hide)

    def on_info_btn_clicked(self):
        self.about_view.show()

    def on_process_started(self):
        self.process_view.show()