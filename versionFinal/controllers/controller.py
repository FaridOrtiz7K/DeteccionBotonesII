import tkinter.messagebox as messagebox
import threading
import time
import pyautogui
import pandas as pd
from tkinter import filedialog

class ScriptController:
    def __init__(self):
        from model import ScriptModel
        self.model = ScriptModel()
        self.view = None
        self.login_view = None
        self.pause_dialog = None
        self.script_thread = None
        
    def run(self):
        from view import MainView
        self.view = MainView(self)
        self.view.withdraw()
        self.show_login()
        
    def show_login(self):
        from view import LoginView
        self.login_view = LoginView(self.view, self)
        self.view.wait_window(self.login_view)
        
    def handle_login(self, username, password):
        if self.model.validate_credentials(username, password):
            self.login_view.destroy()
            self.view.deiconify()
            self.view.update_status("Login exitoso. Selecciona un archivo CSV para comenzar.")
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
                for log in self.model.get_recent_logs(5):
                    self.view.update_status(log)
                    
    def start_script(self):
        if self.model.csv_data is None:
            messagebox.showerror("Error", "Por favor selecciona un archivo CSV primero")
            return
            
        if not messagebox.askyesno("Confirmar", 
            f"¿Iniciar script desde línea {self.view.start_count.get()} con {self.view.loop_count.get()} lotes?"):
            return
            
        self.model.set_start_parameters(
            self.view.start_count.get(), 
            self.view.loop_count.get()
        )
        self.model.is_running = True
        self.model.is_paused = False
        
        self.view.update_buttons(True, False)
        self.view.update_status("Script iniciado. Presiona ESC para pausar.")
        
        self.script_thread = threading.Thread(target=self.execute_script)
        self.script_thread.daemon = True
        self.script_thread.start()
        
    def pause_script(self):
        if self.model.is_running:
            self.model.is_paused = not self.model.is_paused
            self.view.update_buttons(True, self.model.is_paused)
            
            if self.model.is_paused:
                self.view.update_status("Script pausado")
                self.show_pause_dialog()
            else:
                self.view.update_status("Script reanudado")
                if self.pause_dialog:
                    self.pause_dialog.destroy()
                    self.pause_dialog = None
                
    def stop_script(self):
        self.model.is_running = False
        self.model.is_paused = False
        self.view.update_buttons(False, False)
        self.view.update_status("Script detenido")
        if self.pause_dialog:
            self.pause_dialog.destroy()
            self.pause_dialog = None
            
    def toggle_pause(self):
        if self.model.is_running and not self.model.is_paused:
            self.pause_script()
            
    def show_pause_dialog(self):
        from view import PauseDialog
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
                
            while self.model.is_paused:
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
            self.model.add_log("Script completado exitosamente")
            self.view.update_status("Script completado")
            self.model.is_running = False
            self.view.after(0, lambda: self.view.update_buttons(False, False))
            
    def process_row(self, row, line_number):
        """Procesa una fila individual del CSV - IMPLEMENTA AQUÍ TU LÓGICA ESPECÍFICA"""
        self.model.add_log(f"Procesando línea {line_number}")
        self.view.update_status(f"Procesando línea {line_number}")
        
        # EJEMPLO DE IMPLEMENTACIÓN - COMPLETA CON TU LÓGICA ESPECÍFICA
        
        # 1. Click en lista
        pyautogui.click(self.model.coords['select_list'])
        time.sleep(1.5)
        
        # 2. Click en caso número
        pyautogui.click(self.model.coords['case_numero'])
        time.sleep(1.5)
        
        # 3. Presionar delete
        pyautogui.press('delete')
        time.sleep(1)
        
        # 4. Escribir valor de columna B (índice 1)
        value_b = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        pyautogui.write(value_b)
        time.sleep(1.5)
        
        # 5. Verificar columna D (índice 3)
        if pd.notna(row.iloc[3]) and row.iloc[3] > 0:
            pyautogui.click(1507, 650)  # Seleccionar punto USO
            time.sleep(2.5)
            
            for _ in range(int(row.iloc[3])):
                pyautogui.press('down')
                time.sleep(2)
        
        # 6. Presionar Actualizar
        pyautogui.click(self.model.coords['actualizar'])
        time.sleep(1.5)
        
        # 7. Verificar columna E (índice 4) para lógica U/V
        if pd.notna(row.iloc[4]):
            if row.iloc[4] == "U":
                self.process_u_logic(row)
            elif row.iloc[4] == "V":
                self.process_v_logic(row)
        
        # 8. Continuar con el resto de tu lógica específica...
        
        self.model.add_log(f"Línea {line_number} procesada exitosamente")
    
    def process_u_logic(self, row):
        """Procesa la lógica para el caso U"""
        pyautogui.click(self.model.coords['seleccionar_mapa'])
        time.sleep(2)
        
        pyautogui.click(self.model.coords['asignar_un_nse'])
        time.sleep(2)
        
        pyautogui.click(self.model.coords['casilla_un_nse'])
        time.sleep(2)
        
        # Procesar columnas F a P (índices 5 a 15)
        for col_index in range(5, 16):
            if pd.notna(row.iloc[col_index]) and row.iloc[col_index] > 0:
                if col_index in self.model.coords_select_u:
                    x, y = self.model.coords_select_u[col_index]
                    pyautogui.click(x, y)
                    time.sleep(3)
        
        pyautogui.click(self.model.coords['confirmar_nse_u'])
        time.sleep(2)
    
    def process_v_logic(self, row):
        """Procesa la lógica para el caso V"""
        pyautogui.click(self.model.coords['seleccionar_mapa'])
        time.sleep(3)
        
        pyautogui.click(self.model.coords['asignar_varios_nse'])
        time.sleep(3)
        
        # Procesar columnas F a P (índices 5 a 15)
        for col_index in range(5, 16):
            if pd.notna(row.iloc[col_index]) and row.iloc[col_index] > 0:
                if col_index in self.model.coords_select_v:
                    # Click en coordenada select
                    x_cs, y_cs = self.model.coords_select_v[col_index]
                    pyautogui.click(x_cs, y_cs)
                    time.sleep(2)
                    
                    # Click en coordenada type
                    x_ct, y_ct = self.model.coords_type_v[col_index]
                    pyautogui.click(x_ct, y_ct)
                    time.sleep(2)
                    
                    # Escribir valor
                    pyautogui.write(str(row.iloc[col_index]))
                    time.sleep(2)
        
        pyautogui.click(self.model.coords['confirmar_nse_v'])
        time.sleep(2)
        
        pyautogui.click(self.model.coords['cerrar_asignar_nse'])
        time.sleep(2)