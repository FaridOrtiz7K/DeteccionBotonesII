import tkinter as tk
from tkinter import ttk

class LoginView(tk.Toplevel):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.title("Login - Script SIGP")
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
        self.controller.handle_login(self.username.get(), self.password.get())