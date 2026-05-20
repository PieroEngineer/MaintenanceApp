from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QDialog, QLabel, QPushButton, QHBoxLayout, QVBoxLayout

from views.widgets.log_block import LogBlock

class ProcessWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Registro de procesamiento")
        self.setModal(False)  # non-modal
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, False)

        self.log_top_layout = QHBoxLayout()

        self.month_lbl = QLabel()
        self.month_lbl.setText('Esperando a correr programa para saber el mes')

        self.get_log_btn = QPushButton('🗳️  Guardar registro en archivo de texto')
        self.get_log_btn.setFixedWidth(250)
        self.get_log_btn.setEnabled(False)

        self.log_top_layout.addWidget(self.month_lbl)
        self.log_top_layout.addWidget(self.get_log_btn)

        self.text_frame = LogBlock()
        
        # ---- Main Layout ----
        self.main_layout = QVBoxLayout(self)
        self.main_layout.addLayout(self.log_top_layout)
        self.main_layout.addWidget(self.text_frame)

        # self.setMinimumWidth(420)