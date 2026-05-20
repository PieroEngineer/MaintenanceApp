import textwrap

from PyQt6.QtCore import QThread

from models.process_model import ProcessModel

from views.windows.main_window import MainWindow
from views.windows.process_window import ProcessWindow

from models.path_model import PathModel

class ProcessController:
    def __init__(self, path_model: PathModel, view: MainWindow, process_view: ProcessWindow, app_controller):

        self.app_controller = app_controller
        
        self.view = view
        self.process_view = process_view
        
        self.path_model = path_model

        self.view.op_btn.clicked.connect(self.start_processing)
        self.view.voltage_correction_chbx.stateChanged.connect(self.on_drive_actived)
        self.process_view.get_log_btn.clicked.connect(self.get_log)

    def on_drive_actived(self, state):
        self.view.driver_ads_lbl.setVisible(state==2)

    def update_input_path(self, path):
        self.process_model.update_excel_path(path)

    def update_output_path(self, path):
        self.process_model.update_folder_path(path)
   
    def start_processing(self): # Thread starter
        self.app_controller.on_process_started()

        current_input_path = self.path_model.return_input_path()
        current_output_path = self.path_model.return_output_path()
        
        if current_input_path and current_output_path:

            self.view.loadbar.start_indeterminate()
            
            # Disabling UI temporally
            self.view.voltage_correction_chbx.setEnabled(False)
            self.view.expect_xls_chbx.setEnabled(False)
            self.view.op_btn.setEnabled(False)
            self.view.input_source_btn.setEnabled(False)
            self.view.output_source_btn.setEnabled(False)
            self.process_view.get_log_btn.setEnabled(False)
            
            self.view.loadbar.update_tittle('Empezando carga...')

            # Managing the process model
            self.process_model = ProcessModel()
            self.process_model.pme_access_enabled = self.view.voltage_correction_chbx.isChecked()
            self.view.step_indicator.set_steps([textwrap.fill(available_op[1], 13, break_long_words=False) for  available_op in self.process_model.get_ops()]) 

            self.update_input_path(current_input_path)
            self.update_output_path(current_output_path)

            self.process_model.use_xls_format(self.view.expect_xls_chbx.isChecked())

            # Reseting log UI
            self.process_view.text_frame.clear_log()
            self.view.step_indicator.reset_steps()

            self.process_view.text_frame.add_line('🏗️  Extrayendo datos, demorará dependiendo del tamaño...')## Change it by a loading bar time-depended

            # Configuring thread
            self.thread = QThread()
            self.process_model.moveToThread(self.thread)

            self.thread.started.connect(self.process_model.process_data)

            # Signal connections
            self.process_model.excel_to_df_completed.connect(self.register_a_log)
            self.process_model.initial_data_got.connect(self.show_initial)
            self.process_model.data_amount_established.connect(self.initialize_load_bar)    ## Maybe for those that only have one line, write them here directly
            self.process_model.month_got.connect(self.show_month_processed)
            self.process_model.operation_started.connect(self.register_a_log)
            self.process_model.data_processed.connect(self.register_a_log)
            self.process_model.one_more_proccesed.connect(self.update_operation_load_bar)
            self.process_model.error_ocurred.connect(self.register_a_log)
            self.process_model.xls_no_possible.connect(self.register_a_log)
            self.process_model.operation_done.connect(self.register_a_log)
            self.process_model.operation_done.connect(self.when_operation_done)
            self.process_model.final_data_ready.connect(self.show_summarize)
            self.process_model.final_data_ready.connect(self.change_loadbar_tittle)
            self.process_model.ready_for_report.connect(self.create_report)
            self.process_model.excel_saved.connect(self.update_operation_load_bar)
            self.process_model.df_to_excel_completed.connect(self.when_df_to_excel_completed)

            # Clean up automatically after finish
            self.process_model.df_to_excel_completed.connect(self.thread.quit)
            self.process_model.df_to_excel_completed.connect(self.process_model.deleteLater)
            self.thread.finished.connect(self.thread.deleteLater)

            self.thread.start()

#----Functions connected----
    def get_log(self):
        content = self.process_view.text_frame.get_content()
        self.process_model.convert_log_info_into_txt(content)

    def register_a_log(self, text = 'No llegó el registro'):
        self.process_view.text_frame.add_line(text)

    def show_month_processed(self, current_month):
        self.process_view.month_lbl.setText(f'Mes siendo procesado es <b>{current_month}</b>')

        #Next step
        self.view.loadbar.stop_indeterminate()

    def when_operation_done(self):
        self.view.step_indicator.one_step_finished()
        self.view.loadbar.reset_loading()

    def show_summarize(self, nms_dfs):
        self.process_view.text_frame.add_line("    🗃️  La data está lista y se está subiendo, no cerrar el programa hasta que culmine")
        self.view.summarize_table.update_table(nms_dfs)

        self.process_view.get_log_btn.setEnabled(True)

    def show_initial(self, nms_dfs):
        self.view.initial_table.update_table(nms_dfs)

    def change_loadbar_tittle(self):
        self.view.loadbar.update_tittle('Analizando posibles advertencias')
        #self.view.loadbar.update_tittle('Ahora se están guardando la data en excels; puedes revisarlos mientras carga')
    
    def initialize_load_bar(self, total):
        self.view.loadbar.set_data_amount(total)

    def update_operation_load_bar(self, current):
        self.view.loadbar.update_progress(current) ## Verify what function correspong here

    def create_report(self, errors, voltage_0_n, filled_rows, filled_hipos):
        self.register_a_log('\n\n📋  Reporte final:')
        
        self.process_view.text_frame.clear_log()
        
        if errors or voltage_0_n or filled_rows or filled_hipos:
            if errors:
                self.register_a_log('🤕  Errores encontrados:')
                for op_name, df_name in errors:
                    self.register_a_log(f'  ‼️  El documento {df_name} para la operación {op_name}')

            if voltage_0_n:
                self.register_a_log('🧐  Advertencias de columnas VII avg mean vacías detectadas\n🛠️  (Descargar el ODBC Driver 18 para servidor SQL para que la reparación automática sea positble):')
                for df_name, advertisement in voltage_0_n:
                    self.register_a_log(f'  ⚠️ Para {df_name},{advertisement}, iniciando reparación automática con PME')
            
            if filled_rows:
                self.register_a_log('🧐  Advertencias conjunto de registros corregidos:')
                for df_name, empty_groups_found in filled_rows.items():
                    self.register_a_log(f'  🔎  Elementos faltantes corregidos en {df_name}:')
                    for missing in empty_groups_found:
                        self.register_a_log(f'    ⚠️  El {missing}')

            if filled_hipos:
                self.register_a_log('🧐  Advertencias hipos corregidos:')
                for df_name, empty_groups_found in filled_hipos.items():
                    self.register_a_log(f'  🔎  Hipos corregidos en {df_name}')
                    for missing in empty_groups_found:
                        self.register_a_log(f'    ⚠️  Hipo corregidos en {missing}')
                
        else:
            self.register_a_log('  🫡  No hubo errores, inge ✅')

        # Next step
        self.view.loadbar.update_tittle('Guardando la data en excels; puedes revisarlos mientras carga')
         
    def when_df_to_excel_completed(self):
        self.process_view.text_frame.add_line('👋  Operación terminada, los datos han sido tratados y guardados en las rutas especificadas por el usuario ✅')
        self.view.loadbar.update_tittle('Todos los excels han terminado de guardarse')
        self.view.loadbar.fill_bar()
        self.thread.quit()
        self.thread.wait()
        self.view.input_source_btn.setEnabled(True)
        self.view.output_source_btn.setEnabled(True)
        self.view.op_btn.setEnabled(True)
        self.view.voltage_correction_chbx.setEnabled(True)
        self.view.expect_xls_chbx.setEnabled(True)
