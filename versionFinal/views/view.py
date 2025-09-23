import tkinter as tk
from tkinter import ttk, filedialog, messagebox

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
        
        ttk.Label(main_frame, text="Usuario:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.username).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(main_frame, text="Contraseña:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.password, show="*").grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Login", command=self.login).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side=tk.LEFT, padx=5)
        
        main_frame.columnconfigure(1, weight=1)
        
    def login(self):
        self.controller.handle_login(self.username.get(), self.password.get())

class MainView(tk.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("Script Automatización SIGP")
        self.geometry("600x500")
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
        main_frame.rowconfigure(4, weight=1)
        
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
        ttk.Label(main_frame, text="Log de ejecución:").grid(row=4, column=0, sticky=tk.W, pady=(10, 0))
        
        self.status_text = tk.Text(main_frame, height=15, width=70, state="disabled")
        self.status_text.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=5, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
    def bind_events(self):
        self.bind('<Escape>', lambda e: self.controller.toggle_pause())
        
    def update_status(self, message):
        self.status_text.configure(state="normal")
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.status_text.configure(state="disabled")
        
    def update_buttons(self, running, paused):
        if running:
            self.start_button.config(state="disabled")
            self.pause_button.config(state="normal")
            self.stop_button.config(state="normal")
            self.pause_button.config(text="Continuar" if paused else "Pausar")
        else:
            self.start_button.config(state="normal")
            self.pause_button.config(state="disabled")
            self.stop_button.config(state="disabled")
            self.pause_button.config(text="Pausar")
            
    def on_closing(self):
        self.controller.handle_app_close()

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
        self.grab_set()
        
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