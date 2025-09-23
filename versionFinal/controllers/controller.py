import tkinter.messagebox as messagebox
import threading
import time
import pyautogui
import pandas as pd
from tkinter import filedialog

class ScriptController:
    def __init__(self):
        from model import ScriptModel # pyright: ignore[reportMissingImports]
        self.model = ScriptModel()
        self.view = None
        self.login_view = None
        self.pause_dialog = None
        self.script_thread = None
        
    def run(self):
        from main_view import MainView # pyright: ignore[reportMissingImports]
        self.view = MainView(self)
        self.view.withdraw()  # Ocultar ventana principal hasta login
        self.show_login()
        
    def show_login(self):
        from login_view import LoginView # pyright: ignore[reportMissingImports]
        self.login_view = LoginView(self.view, self)
        self.view.wait_window(self.login_view)
        
    def handle_login(self, username, password):
        if self.model.validate_credentials(username, password):
            self.login_view.destroy()
            self.view.deiconify()  # Mostrar ventana principal
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
                # Mostrar logs recientes
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
        from pause_view import PauseDialog # pyright: ignore[reportMissingImports]
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
        """Procesa una fila individual del CSV"""
        self.model.add_log(f"Procesando línea {line_number}")
        self.view.update_status(f"Procesando línea {line_number}")
        
        # Implementación de la lógica específica del script original
        # Aquí debes colocar todas las acciones de tu script AutoHotkey
        
        # Ejemplo de implementación:
        pyautogui.click(self.model.coords['select_list'])
        time.sleep(1.5)
        
        pyautogui.click(self.model.coords['case_numero'])
        time.sleep(1.5)
        
        pyautogui.press('delete')
        time.sleep(1)
        
        # Escribir valor de columna B (índice 1)
        value_b = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        pyautogui.write(value_b)
        time.sleep(1.5)
        
        # Verificar columna D (índice 3)
        if pd.notna(row.iloc[3]) and row.iloc[3] > 0:
            pyautogui.click(1507, 650)  # Seleccionar punto USO
            time.sleep(2.5)
            
            for _ in range(int(row.iloc[3])):
                pyautogui.press('down')
                time.sleep(2)
        
        # Presionar Actualizar
        pyautogui.click(self.model.coords['actualizar'])
        time.sleep(1.5)
        
        # Continuar con el resto de tu lógica específica...
        # Implementa aquí todas las acciones de tu script original
        
        self.model.add_log(f"Línea {line_number} procesada exitosamente")