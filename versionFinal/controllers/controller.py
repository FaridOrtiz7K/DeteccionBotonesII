import tkinter.messagebox as messagebox
import threading
import time
import pyautogui
import pandas as pd
from tkinter import filedialog
from datetime import datetime

class ScriptController:
    def __init__(self):
        from model import ScriptModel
        self.model = ScriptModel()
        self.view = None
        self.login_view = None
        self.pause_dialog = None
        self.script_thread = None
        
    def run(self):
        from main_view import MainView
        self.view = MainView(self)
        self.view.withdraw()  # Ocultar ventana principal hasta login
        self.show_login()
        
    def show_login(self):
        from login_view import LoginView
        self.login_view = LoginView(self.view, self)
        self.view.wait_window(self.login_view)
        
    def handle_login(self, username, password):
        if self.model.validate_credentials(username, password):
            self.login_view.destroy()
            self.view.deiconify()  # Mostrar ventana principal
            self.view.update_status("Login exitoso. Selecciona un archivo CSV para comenzar.")
            self.view.update_status_label("Listo")
        else:
            messagebox.showerror("Error", "Usuario o contraseña incorrectos")
            
    def load_csv_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            if self.model.load_csv(file_path):
                self.view.csv_path.set(file_path)
                self.view.update_status(f"Archivo CSV cargado: {file_path}")
                self.view.update_status(f"Total de líneas: {self.model.total_lines}")
                self.view.update_status_label("CSV Cargado")
                
    def start_script(self):
        if self.model.csv_data is None:
            messagebox.showerror("Error", "Por favor selecciona un archivo CSV primero")
            return
            
        start_line = self.view.start_count.get()
        loop_count = self.view.loop_count.get()
        
        if start_line < 1 or start_line > self.model.total_lines:
            messagebox.showerror("Error", f"Línea de inicio debe estar entre 1 y {self.model.total_lines}")
            return
            
        if not messagebox.askyesno("Confirmar", 
            f"¿Iniciar script desde línea {start_line} con {loop_count} lotes?"):
            return
            
        self.model.set_start_parameters(start_line, loop_count)
        self.model.is_running = True
        self.model.is_paused = False
        
        self.view.update_buttons(True, False)
        self.view.update_status("Script iniciado. Presiona ESC para pausar.")
        self.view.update_status_label("Ejecutando")
        
        # Iniciar el script en un hilo separado
        self.script_thread = threading.Thread(target=self.execute_script)
        self.script_thread.daemon = True
        self.script_thread.start()
        
    def pause_script(self):
        if self.model.is_running:
            self.model.is_paused = not self.model.is_paused
            self.view.update_buttons(True, self.model.is_paused)
            
            if self.model.is_paused:
                self.view.update_status("Script pausado")
                self.view.update_status_label("Pausado")
                self.show_pause_dialog()
            else:
                self.view.update_status("Script reanudado")
                self.view.update_status_label("Ejecutando")
                if self.pause_dialog:
                    self.pause_dialog.destroy()
                    self.pause_dialog = None
                
    def stop_script(self):
        if self.model.is_running:
            self.model.is_running = False
            self.model.is_paused = False
            self.view.update_buttons(False, False)
            self.view.update_status("Script detenido por el usuario")
            self.view.update_status_label("Detenido")
            if self.pause_dialog:
                self.pause_dialog.destroy()
                self.pause_dialog = None
            
    def toggle_pause(self):
        if self.model.is_running and not self.model.is_paused:
            self.pause_script()
            
    def show_pause_dialog(self):
        from pause_view import PauseDialog
        if self.pause_dialog is None or not self.pause_dialog.winfo_exists():
            self.pause_dialog = PauseDialog(self.view, self)
            
    def handle_app_close(self):
        if self.model.is_running:
            if messagebox.askokcancel("Salir", "El script está en ejecución. ¿Estás seguro de que quieres salir?"):
                self.stop_script()
                self.view.destroy()
        else:
            self.view.destroy()
            
    def execute_script(self):
        """Ejecuta el script principal de automatización"""
        cont = self.model.current_line
        end_line = min(cont + self.model.loop_count, self.model.total_lines)
        
        self.model.add_log(f"Iniciando procesamiento desde línea {cont + 1} hasta {end_line}")
        self.view.update_status(f"Iniciando desde línea {cont + 1} hasta {end_line}")
        
        for i in range(cont, end_line):
            if not self.model.is_running:
                break
                
            while self.model.is_paused and self.model.is_running:
                time.sleep(0.1)
                if not self.model.is_running:
                    break
                    
            if not self.model.is_running:
                break
                
            current_row = self.model.get_current_row()
            if current_row is None:
                break
                
            try:
                self.process_row(current_row, i + 1)
                self.model.current_line = i + 1
                
            except Exception as e:
                error_msg = f"Error en línea {i + 1}: {str(e)}"
                self.model.add_log(error_msg)
                self.view.update_status(error_msg)
                continue
                
        if self.model.is_running:
            completion_msg = "Script completado exitosamente"
            self.model.add_log(completion_msg)
            self.view.update_status(completion_msg)
            self.view.update_status_label("Completado")
            self.model.is_running = False
            self.view.after(0, lambda: self.view.update_buttons(False, False))
            
    def process_row(self, row, line_number):
        """Procesa una fila individual del CSV - IMPLEMENTACIÓN COMPLETA DEL SCRIPT ORIGINAL"""
        self.model.add_log(f"Procesando línea {line_number}")
        self.view.update_status(f"Procesando línea {line_number}")
        
        try:
            # 1. Left click at (89, 263) - Select in the list
            pyautogui.click(89, 263)
            time.sleep(1.5)

            # 2. Left click at (1483, 519) - Case numero
            pyautogui.click(1483, 519)
            time.sleep(1.5)

            # 3. Press delete key
            pyautogui.press('delete')
            time.sleep(1)

            # 4. Write the value from column B (index 1)
            value_b = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
            pyautogui.write(value_b)
            time.sleep(1.5)

            # 5. Check the value in column D (index 3)
            if pd.notna(row.iloc[3]) and row.iloc[3] > 0:
                # Select the point called USO at (1507, 636)
                pyautogui.click(1507, 650)
                time.sleep(2.5)

                # Send the number of clicks down based on column D
                for _ in range(int(row.iloc[3])):
                    pyautogui.press('down')
                    time.sleep(2)

            # 6. Press Actualizar at (1290, 349)
            pyautogui.click(1290, 349)
            time.sleep(1.5)

            # 7. Check the value in column E (index 4)
            if pd.notna(row.iloc[4]):
                if row.iloc[4] == "U":
                    self.process_u_logic(row)
                elif row.iloc[4] == "V":
                    self.process_v_logic(row)

            # 8. Check for services (column 16 index 15)
            if len(row) > 16 and pd.notna(row.iloc[15]) and row.iloc[15] > 0:
                self.process_services_logic(row)

            # 9. Left click again at (89, 263)
            pyautogui.click(89, 263)
            time.sleep(3)

            # 10. Press the directional key down
            pyautogui.press('down')
            time.sleep(3)

            self.model.add_log(f"Línea {line_number} procesada exitosamente")
            
        except Exception as e:
            error_msg = f"Error procesando línea {line_number}: {str(e)}"
            self.model.add_log(error_msg)
            self.view.update_status(error_msg)
            raise
    
    def process_u_logic(self, row):
        """Procesa la lógica para el caso U"""
        # Left click at (169, 189) Seleccionar en mapa
        pyautogui.click(169, 189)
        time.sleep(2)

        # Left click at (1463, 382) Asignar un solo NSE
        pyautogui.click(1463, 382)
        time.sleep(2)

        # Left click at (1266, 590) Casilla un solo NSE
        pyautogui.click(1266, 590)
        time.sleep(2)

        # Handle "U" logic for columns F to P (indices 5 to 15)
        coords_select_u = {
            5: (1268, 637), 6: (1268, 661), 7: (1268, 685), 8: (1268, 709),
            9: (1268, 733), 10: (1268, 757), 11: (1268, 781), 12: (1268, 825),
            13: (1268, 856), 14: (1268, 881), 15: (1268, 908)
        }
        
        for col_index in range(5, 16):
            if len(row) > col_index and pd.notna(row.iloc[col_index]) and row.iloc[col_index] > 0:
                if col_index in coords_select_u:
                    x, y = coords_select_u[col_index]
                    pyautogui.click(x, y)
                    time.sleep(3)

        # Left click at (1306, 639) Asignar NSE (Confirm)
        pyautogui.click(1306, 639)
        time.sleep(2)
    
    def process_v_logic(self, row):
        """Procesa la lógica para el caso V"""
        # Left click at (169, 189) Seleccionar en mapa
        pyautogui.click(169, 189)
        time.sleep(3)

        # Left click at (1491, 386) Asignar varios NSE
        pyautogui.click(1491, 386)
        time.sleep(3)

        # Handle "V" logic for columns F to P (indices 5 to 15)
        coords_select = {
            5: (1235, 563), 6: (1235, 602), 7: (1235, 630), 8: (1235, 668),
            9: (1235, 702), 10: (1600, 563), 11: (1600, 602), 12: (1600, 630),
            13: (1235, 772), 14: (1235, 804), 15: (1235, 838)
        }
        
        coords_type = {
            5: (1365, 563), 6: (1365, 602), 7: (1365, 630), 8: (1365, 668),
            9: (1365, 702), 10: (1730, 563), 11: (1730, 602), 12: (1730, 630),
            13: (1365, 772), 14: (1365, 804), 15: (1365, 838)
        }
        
        for col_index in range(5, 16):
            if len(row) > col_index and pd.notna(row.iloc[col_index]) and row.iloc[col_index] > 0:
                if col_index in coords_select and col_index in coords_type:
                    # Click en coordenada select
                    x_cs, y_cs = coords_select[col_index]
                    pyautogui.click(x_cs, y_cs)
                    time.sleep(2)
                    
                    # Click en coordenada type
                    x_ct, y_ct = coords_type[col_index]
                    pyautogui.click(x_ct, y_ct)
                    time.sleep(2)
                    
                    # Escribir valor
                    pyautogui.write(str(row.iloc[col_index]))
                    time.sleep(2)

        # Left click at (1648, 752) - Asignar NSE (Confirm)
        pyautogui.click(1648, 752)
        time.sleep(2)
        
        # Left click at (1598, 823) - Cerrar Asignar NSE 
        pyautogui.click(1598, 823)
        time.sleep(2)
    
    def process_services_logic(self, row):
        """Procesa la lógica de servicios administrativos"""
        # Left click at (1563, 385) Boton Administrar Servicios
        pyautogui.click(1563, 385)
        time.sleep(2)
        
        # Left click at (100, 114) Selecciona detalle con NSE
        pyautogui.click(100, 114)
        time.sleep(2)
        
        # Procesar diferentes tipos de servicios
        service_handlers = {
            16: self.handle_voz_cobre,    # Columna Q (index 16)
            17: self.handle_datos_sdom,   # Columna R (index 17)
            18: self.handle_datos_cobre,  # Columna S (index 18)
            19: self.handle_datos_fibra,  # Columna T (index 19)
            20: self.handle_tv_cable,     # Columna U (index 20)
            21: self.handle_dish,         # Columna V (index 21)
            22: self.handle_tvs,          # Columna W (index 22)
            23: self.handle_sky,          # Columna X (index 23)
            24: self.handle_vetv          # Columna Y (index 24)
        }
        
        for col_index, handler in service_handlers.items():
            if len(row) > col_index and pd.notna(row.iloc[col_index]) and row.iloc[col_index] > 0:
                # Siempre comenzar desde la ventana de detalle
                pyautogui.click(100, 114)
                time.sleep(2)
                handler(row, col_index)
        
        # Cerrar ventana de administrar servicios
        pyautogui.click(882, 49)
        time.sleep(5)
    
    def handle_voz_cobre(self, row, col_index):
        """Maneja VOZ COBRE TELMEX LINEAS DE COBRE"""
        # El servicio ya está seleccionado por defecto
        pyautogui.click(127, 383)  # Cantidad
        time.sleep(2)
        pyautogui.write(str(row.iloc[col_index]))  # Cantidad de servicios
        time.sleep(2)
        pyautogui.click(82, 423)  # Guardar
        time.sleep(2)
        self.handle_posible_error()
    
    def handle_datos_sdom(self, row, col_index):
        """Maneja Datos s/dom"""
        pyautogui.click(138, 269)  # Servicio
        time.sleep(2)
        # Seleccionar Datos (2 down)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(127, 383)  # Cantidad
        time.sleep(2)
        pyautogui.write(str(row.iloc[col_index]))  # Cantidad de servicios
        time.sleep(2)
        pyautogui.click(82, 423)  # Guardar
        time.sleep(2)
        self.handle_posible_error()
    
    def handle_datos_cobre(self, row, col_index):
        """Maneja Datos-cobre-telmex-inf"""
        pyautogui.click(138, 269)  # Servicio
        time.sleep(2)
        # Seleccionar Datos (2 down)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(159, 355)  # Producto
        time.sleep(2)
        pyautogui.press('down')  # Seleccionar producto
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(127, 383)  # Cantidad
        time.sleep(2)
        pyautogui.write(str(row.iloc[col_index]))  # Cantidad de servicios
        time.sleep(2)
        pyautogui.click(82, 423)  # Guardar
        time.sleep(2)
        self.handle_posible_error()
    
    def handle_datos_fibra(self, row, col_index):
        """Maneja Datos-fibra-telmex-inf"""
        pyautogui.click(138, 269)  # Servicio
        time.sleep(2)
        # Seleccionar Datos (2 down)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(152, 294)  # Tipo
        time.sleep(2)
        pyautogui.press('down')  # Seleccionar tipo
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(150, 323)  # Empresa
        time.sleep(2)
        pyautogui.press('down')  # Seleccionar empresa
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(127, 383)  # Cantidad
        time.sleep(2)
        pyautogui.write(str(row.iloc[col_index]))  # Cantidad de servicios
        time.sleep(2)
        pyautogui.click(82, 423)  # Guardar
        time.sleep(2)
        self.handle_posible_error()
    
    def handle_tv_cable(self, row, col_index):
        """Maneja TV cable otros"""
        pyautogui.click(138, 269)  # Servicio
        time.sleep(2)
        # Seleccionar TV (3 down)
        for _ in range(3):
            pyautogui.press('down')
            time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(150, 323)  # Empresa
        time.sleep(2)
        # Seleccionar otros (4 down)
        for _ in range(4):
            pyautogui.press('down')
            time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(127, 383)  # Cantidad
        time.sleep(2)
        pyautogui.write(str(row.iloc[col_index]))  # Cantidad de servicios
        time.sleep(2)
        pyautogui.click(82, 423)  # Guardar
        time.sleep(2)
        self.handle_posible_error()
    
    def handle_dish(self, row, col_index):
        """Maneja Dish"""
        pyautogui.click(138, 269)  # Servicio
        time.sleep(2)
        # Seleccionar TV (3 down)
        for _ in range(3):
            pyautogui.press('down')
            time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(152, 294)  # Tipo
        time.sleep(2)
        # Seleccionar Satelital (2 down)
        for _ in range(2):
            pyautogui.press('down')
            time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(150, 323)  # Empresa
        time.sleep(2)
        pyautogui.press('down')  # Dish
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(127, 383)  # Cantidad
        time.sleep(2)
        pyautogui.write(str(row.iloc[col_index]))  # Cantidad de servicios
        time.sleep(2)
        pyautogui.click(82, 423)  # Guardar
        time.sleep(2)
        self.handle_posible_error()
    
    def handle_tvs(self, row, col_index):
        """Maneja TVS"""
        pyautogui.click(138, 269)  # Servicio
        time.sleep(2)
        # Seleccionar TV (3 down)
        for _ in range(3):
            pyautogui.press('down')
            time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(152, 294)  # Tipo
        time.sleep(2)
        # Seleccionar Satelital (2 down)
        for _ in range(2):
            pyautogui.press('down')
            time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        
        pyautogui.click(150, 323)  # Empresa
        time.sleep(2)
        # Seleccionar otro (2 down)
        for _ in range(2):
            pyautogui.press('down')
            time.sleep(1)
        pyautogui