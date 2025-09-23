import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

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
        main_frame.rowconfigure(5, weight=1)
        
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
        
        # Información del estado
        self.status_label = ttk.Label(main_frame, text="Estado: Listo")
        self.status_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Botones de control
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        self.start_button = ttk.Button(button_frame, text="Iniciar", command=self.controller.start_script)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.pause_button = ttk.Button(button_frame, text="Pausar", command=self.controller.pause_script, state="disabled")
        self.pause_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Detener", command=self.controller.stop_script, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Estado del script
        ttk.Label(main_frame, text="Log de ejecución:").grid(row=5, column=0, sticky=tk.W, pady=(10, 0))
        
        self.status_text = tk.Text(main_frame, height=15, width=70, state="disabled")
        self.status_text.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Scrollbar para el texto
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=6, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
    def bind_events(self):
        self.bind('<Escape>', lambda e: self.controller.toggle_pause())
        
    def update_status(self, message):
        self.status_text.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.configure(state="disabled")
        
    def update_status_label(self, message):
        self.status_label.config(text=f"Estado: {message}")
        
    def update_buttons(self, running, paused):
        if running:
            self.start_button.config(state="disabled")
            self.pause_button.config(state="normal")
            self.stop_button.config(state="normal")
            if paused:
                self.pause_button.config(text="Continuar")
                self.update_status_label("Pausado")
            else:
                self.pause_button.config(text="Pausar")
                self.update_status_label("Ejecutando")
        else:
            self.start_button.config(state="normal")
            self.pause_button.config(state="disabled")
            self.stop_button.config(state="disabled")
            self.pause_button.config(text="Pausar")
            self.update_status_label("Listo")
            
    def on_closing(self):
        self.controller.handle_app_close()