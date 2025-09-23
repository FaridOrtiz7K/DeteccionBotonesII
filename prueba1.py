import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pyautogui
import time
import threading
import os
from datetime import datetime

# Modelo
class ScriptModel:
    def __init__(self):
        self.csv_data = None
        self.current_line = 0
        self.total_lines = 0
        self.is_running = False
        self.is_paused = False
        self.start_count = 1
        self.loop_count = 0
        self.csv_file_path = ""
        
    def load_csv(self, file_path):
        try:
            self.csv_data = pd.read_csv(file_path)
            self.total_lines = len(self.csv_data)
            self.csv_file_path = file_path
            return True
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el CSV: {str(e)}")
            return False
    
    def reset(self):
        self.current_line = 0
        self.is_running = False
        self.is_paused = False
    
    def get_current_row(self):
        if self.csv_data is not None and self.current_line < self.total_lines:
            return self.csv_data.iloc[self.current_line]
        return None

# Vista
class LoginView(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Login")
        self.geometry("300x150")
        self.resizable(False, False)
        
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        
        self.create_widgets()
        self.center_window()
        
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Usuario
        ttk.Label(main_frame, text="Usuario:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.username).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Contraseña
        ttk.Label(main_frame, text="Contraseña:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.password, show="*").grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Login", command=self.login).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side=tk.LEFT, padx=5)
        
        # Configurar pesos de columnas
        main_frame.columnconfigure(1, weight=1)
        
    def login(self):
        if self.username.get() == "admin" and self.password.get() == "admin":
            self.destroy()
            self.parent.show_main_view()
        else:
            messagebox.showerror("Error", "Usuario o contraseña incorrectos")

class MainView(tk.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("Script Automatización SIGP")
        self.geometry("500x400")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.create_widgets()
        self.bind_events()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid weights
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Selección de archivo CSV
        ttk.Label(main_frame, text="Archivo CSV:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.csv_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.csv_path, state="readonly").grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="Examinar", command=self.controller.load_csv_file).grid(row=0, column=2, padx=5)
        
        # Configuración de inicio
        ttk.Label(main_frame, text="Iniciar en línea:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.start_count = tk.IntVar(value=1)
        ttk.Entry(main_frame, textvariable=self.start_count, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(main_frame, text="Número de lotes:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.loop_count = tk.IntVar(value=589)
        ttk.Entry(main_frame, textvariable=self.loop_count, width=10).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Botones de control
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        self.start_button = ttk.Button(button_frame, text="Iniciar", command=self.controller.start_script)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.pause_button = ttk.Button(button_frame, text="Pausar", command=self.controller.pause_script, state="disabled")
        self.pause_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Detener", command=self.controller.stop_script, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Estado del script
        self.status_text = tk.Text(main_frame, height=10, width=50, state="disabled")
        self.status_text.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Scrollbar para el texto
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=4, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Configurar peso de fila para el área de texto
        main_frame.rowconfigure(4, weight=1)
        
    def bind_events(self):
        self.bind('<Escape>', lambda e: self.controller.toggle_pause())
        
    def update_status(self, message):
        self.status_text.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.configure(state="disabled")
        
    def update_buttons(self, running, paused):
        if running:
            self.start_button.config(state="disabled")
            self.pause_button.config(state="normal")
            self.stop_button.config(state="normal")
            if paused:
                self.pause_button.config(text="Continuar")
            else:
                self.pause_button.config(text="Pausar")
        else:
            self.start_button.config(state="normal")
            self.pause_button.config(state="disabled")
            self.stop_button.config(state="disabled")
            self.pause_button.config(text="Pausar")
            
    def on_closing(self):
        if self.controller.model.is_running:
            if messagebox.askokcancel("Salir", "El script está en ejecución. ¿Estás seguro de que quieres salir?"):
                self.controller.stop_script()
                self.destroy()
        else:
            self.destroy()

class PauseDialog(tk.Toplevel):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.title("Script Pausado")
        self.geometry("300x150")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.continue_script)
        
        self.create_widgets()
        self.center_window()
        self.grab_set()  # Hace la ventana modal
        
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
        
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Script pausado", font=('Arial', 12, 'bold')).pack(pady=10)
        ttk.Label(main_frame, text=f"Línea actual: {self.controller.model.current_line}").pack(pady=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Continuar", command=self.continue_script).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Detener", command=self.stop_script).pack(side=tk.LEFT, padx=5)
        
    def continue_script(self):
        self.controller.pause_script()  # Esto quita la pausa
        self.destroy()
        
    def stop_script(self):
        self.controller.stop_script()
        self.destroy()

# Controlador
class ScriptController:
    def __init__(self):
        self.model = ScriptModel()
        self.view = None
        self.script_thread = None
        self.pause_dialog = None
        
    def run(self):
        self.view = MainView(self)
        self.show_login()
        
    def show_login(self):
        login_view = LoginView(self.view)
        self.view.wait_window(login_view)
        
    def show_main_view(self):
        self.view.deiconify()
        
    def load_csv_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            if self.model.load_csv(file_path):
                self.view.csv_path.set(file_path)
                self.view.update_status(f"CSV cargado: {file_path}")
                self.view.update_status(f"Total de líneas: {self.model.total_lines}")
                
    def start_script(self):
        if self.model.csv_data is None:
            messagebox.showerror("Error", "Por favor selecciona un archivo CSV primero")
            return
            
        if not messagebox.askyesno("Confirmar", f"¿Iniciar script desde línea {self.view.start_count.get()} con {self.view.loop_count.get()} lotes?"):
            return
            
        self.model.start_count = self.view.start_count.get()
        self.model.loop_count = self.view.loop_count.get()
        self.model.current_line = max(0, self.model.start_count - 1)
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
                
    def stop_script(self):
        self.model.is_running = False
        self.model.is_paused = False
        self.view.update_buttons(False, False)
        self.view.update_status("Script detenido")
        
    def toggle_pause(self):
        if self.model.is_running and not self.model.is_paused:
            self.pause_script()
            
    def show_pause_dialog(self):
        if self.pause_dialog is None or not self.pause_dialog.winfo_exists():
            self.pause_dialog = PauseDialog(self.view, self)
            
    def execute_script(self):
        # Coordenadas y configuraciones (debes ajustarlas según tu resolución)
        coords_select_u = {
            6: [1268, 637], 7: [1268, 661], 8: [1268, 685], 9: [1268, 709],
            10: [1268, 733], 11: [1268, 757], 12: [1268, 781], 13: [1268, 825],
            14: [1268, 856], 15: [1268, 881], 16: [1268, 908]
        }
        
        coords_select_v = {
            6: [1235, 563], 7: [1235, 602], 8: [1235, 630], 9: [1235, 668],
            10: [1235, 702], 11: [1600, 563], 12: [1600, 602], 13: [1600, 630],
            14: [1235, 772], 15: [1235, 804], 16: [1235, 838]
        }
        
        coords_type_v = {
            6: [1365, 563], 7: [1365, 602], 8: [1365, 630], 9: [1365, 668],
            10: [1365, 702], 11: [1730, 563], 12: [1730, 602], 13: [1730, 630],
            14: [1365, 772], 15: [1365, 804], 16: [1365, 838]
        }
        
        cont = self.model.current_line
        end_line = min(cont + self.model.loop_count, self.model.total_lines)
        
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
                
            self.view.update_status(f"Procesando línea {cont + 1}")
            
            try:
                # Implementación de las acciones del script original
                # (Aquí irían todas las acciones de clic y envío de teclas)
                
                # Ejemplo de las primeras acciones:
                pyautogui.click(89, 263)  # Select in the list
                time.sleep(1.5)
                
                pyautogui.click(1483, 519)  # Case numero
                time.sleep(1.5)
                
                pyautogui.press('delete')
                time.sleep(1)
                
                # Escribir valor de columna B
                value_b = str(current_row.iloc[1]) if pd.notna(current_row.iloc[1]) else ""
                pyautogui.write(value_b)
                time.sleep(1.5)
                
                # ... continuar con el resto de las acciones del script original
                
                # Actualizar el modelo
                self.model.current_line = cont
                cont += 1
                
            except Exception as e:
                self.view.update_status(f"Error en línea {cont + 1}: {str(e)}")
                continue
                
        if self.model.is_running:
            self.view.update_status("Script completado")
            self.model.is_running = False
            self.view.after(0, lambda: self.view.update_buttons(False, False))

# Ejecución principal
if __name__ == "__main__":
    # Asegurar que pyautogui falle rápido si hay error
    pyautogui.FAILSAFE = True
    
    app = ScriptController()
    app.run()
    tk.mainloop()