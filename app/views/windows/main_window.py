from PyQt6.QtWidgets import QWidget, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QGroupBox, QTabWidget, QCheckBox
from PyQt6.QtCore import Qt

from views.widgets.log_block import LogBlock
from views.widgets.step_indicator import StepIndicator
from views.widgets.load_bar import LoadBar

from views.widgets.summarize_table import SummarizeTable
from views.widgets.initial_table import InitialTable

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Programa de mantenimiento")
        self.setMinimumSize(1200, 650)
        self.setWindowFlags(Qt.WindowType.Window)

        self.main_layout = QVBoxLayout()

        self.source_group = QGroupBox('Selección de rutas:')
        self.operational_group = QGroupBox('Control:')
        self.results_group = QGroupBox('Resultados:')

        self.source_layout = QHBoxLayout(self.source_group)
        self.operational_layout = QHBoxLayout(self.operational_group)
        self.results_layout = QVBoxLayout(self.results_group)

        self.info_btn = QPushButton('🤓 Acerca de...')

        self.main_layout.addWidget(self.source_group, stretch=1) 
        self.main_layout.addWidget(self.operational_group, stretch=3) 
        self.main_layout.addWidget(self.results_group, stretch=5)
        self.main_layout.addWidget(self.info_btn, stretch=1) 

        self.init_source_ui()
        self.init_loading_ui()
        self.init_tab_ui()

        central_widget = QWidget()
        central_widget.setLayout(self.main_layout)
        self.setCentralWidget(central_widget)

    def init_source_ui(self):
        self.input_source_space = QHBoxLayout()
        self.input_source_btn = QPushButton('  🗃️  Selecciona archivo  ')
        self.input_source_btn.setObjectName("accent")
        self.input_source_label = QLabel() 
        self.output_source_space = QHBoxLayout()
        self.output_source_btn = QPushButton('  📂  Selecciona dónde se guardará  ')
        self.output_source_btn.setObjectName("accent")
        self.output_source_label = QLabel()

        self.input_source_space.addWidget(self.input_source_btn)
        self.input_source_space.addWidget(self.input_source_label)
        
        self.source_layout.setSpacing(50)

        self.output_source_space.addWidget(self.output_source_btn)
        self.output_source_space.addWidget(self.output_source_label)

        self.source_layout.addLayout(self.input_source_space)
        self.source_layout.addLayout(self.output_source_space)

    def init_loading_ui(self):
        self.operation_control_layout = QVBoxLayout()
        self.operation_visual_layout = QVBoxLayout()
        
        self.op_btn = QPushButton('🛠️\n Tratamiento de datos ')
        self.op_btn.setObjectName("primary")
        
        self.voltage_correction_chbx = QCheckBox("Acceso a PME")
        self.voltage_correction_chbx.setChecked(False)

        self.driver_ads_lbl = QLabel("Recuerda instalar el\nODBC Driver 18")
        self.driver_ads_lbl.setObjectName('driver_ads')
        self.driver_ads_lbl.setVisible(False)

        self.expect_xls_chbx = QCheckBox("Aplicar XLS en resultado")
        self.expect_xls_chbx.setChecked(False)

        self.step_indicator = StepIndicator()

        self.loadbar = LoadBar()
        self.loadbar.update_tittle('Progreso de la operación actual')

        self.operation_control_layout.addWidget(self.op_btn)
        self.operation_control_layout.addWidget(self.voltage_correction_chbx)
        self.operation_control_layout.addWidget(self.driver_ads_lbl)
        self.operation_control_layout.addWidget(self.expect_xls_chbx)

        self.operation_visual_layout.addWidget(self.step_indicator, stretch = 9)
        self.operation_visual_layout.addWidget(self.loadbar, stretch = 1)

        self.operational_layout.addLayout(self.operation_control_layout, stretch = 1)
        self.operational_layout.addLayout(self.operation_visual_layout, stretch = 9)

    def init_tab_ui(self):
        self.initial_table = InitialTable()
        self.summarize_table = SummarizeTable()

        # Create the table tab
        self.table_tabs = QTabWidget()

        # Initial tab
        self.previous_table_tab = QWidget()

        self.previous_table_tab_layout = QVBoxLayout()
        self.previous_table_tab_layout.addWidget(self.initial_table)

        self.previous_table_tab.setLayout(self.previous_table_tab_layout)

        # Processing tab    (Now in Pop-up window)
        # self.processing_tab = QWidget()

        # self.processing_tab_layout = QVBoxLayout()
        # self.log_top_layout = QHBoxLayout()

        # self.month_lbl = QLabel()
        # self.month_lbl.setText('Esperando a correr programa para saber el mes')

        # self.get_log_btn = QPushButton('🗳️  Guardar registro en archivo de texto')
        # self.get_log_btn.setFixedWidth(250)
        # self.get_log_btn.setEnabled(False)

        # self.log_top_layout.addWidget(self.month_lbl)
        # self.log_top_layout.addWidget(self.get_log_btn)

        # self.text_frame = LogBlock()

        # self.processing_tab_layout.addLayout(self.log_top_layout)
        # self.processing_tab_layout.addWidget(self.text_frame)

        # self.processing_tab.setLayout(self.processing_tab_layout)

        # Summary tab
        self.summarize_table_tab = QWidget()

        self.summarize_table_tab_tab = QVBoxLayout()
        self.summarize_table_tab_tab.addWidget(self.summarize_table)

        self.summarize_table_tab.setLayout(self.summarize_table_tab_tab)

        # Adding tabs to QTabWidget
        self.table_tabs.addTab(self.previous_table_tab, "Pre-procesado")
        # self.table_tabs.addTab(self.processing_tab, "Procesado")
        self.table_tabs.addTab(self.summarize_table_tab, "Post-procesado")

        self.results_layout.addWidget(self.table_tabs)