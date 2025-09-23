import tkinter as tk
from tkinter import ttk

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
        ttk.Label(main_frame, text=f"Línea actual: {self.controller.model.current_line + 1}").pack(pady=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Continuar", command=self.continue_script).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Detener", command=self.stop_script).pack(side=tk.LEFT, padx=5)
        
    def continue_script(self):
        self.controller.pause_script()
        self.destroy()
        
    def stop_script(self):
        self.controller.stop_script()
        self.destroy()