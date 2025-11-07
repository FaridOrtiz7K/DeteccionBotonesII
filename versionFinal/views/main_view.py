# view.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

class VistaAutomation:
    def __init__(self, root, controlador):
        self.root = root
        self.controlador = controlador
        self.root.title("Sistema de Automatización NSE/GE")
        self.root.geometry("800x600")
        
        # Variables
        self.csv_file = tk.StringVar()
        self.linea_maxima = tk.IntVar(value=1)
        self.estado = tk.StringVar(value="Listo")
        self.linea_actual = tk.StringVar(value="0")
        self.lineas_restantes = tk.StringVar(value="0")
        
        self.setup_ui()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Sección CSV
        csv_frame = ttk.LabelFrame(main_frame, text="Configuración CSV", padding="5")
        csv_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        csv_frame.columnconfigure(1, weight=1)
        
        ttk.Label(csv_frame, text="Archivo CSV:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Entry(csv_frame, textvariable=self.csv_file, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(csv_frame, text="Seleccionar", command=self.controlador.seleccionar_csv).grid(row=0, column=2, padx=5)
        
        # Sección ejecución
        exec_frame = ttk.LabelFrame(main_frame, text="Control de Ejecución", padding="5")
        exec_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(exec_frame, text="Línea máxima a ejecutar:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.spin_linea_maxima = ttk.Spinbox(exec_frame, from_=1, to=10000, textvariable=self.linea_maxima, width=10)
        self.spin_linea_maxima.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(exec_frame, text="Línea actual:").grid(row=1, column=0, sticky=tk.W, padx=5)
        ttk.Label(exec_frame, textvariable=self.linea_actual).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(exec_frame, text="Líneas restantes:").grid(row=2, column=0, sticky=tk.W, padx=5)
        ttk.Label(exec_frame, textvariable=self.lineas_restantes).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Botones de control
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        self.btn_iniciar = ttk.Button(control_frame, text="Iniciar Proceso", command=self.controlador.iniciar_proceso)
        self.btn_iniciar.grid(row=0, column=0, padx=5)
        
        self.btn_pausar = ttk.Button(control_frame, text="Pausar (F2)", command=self.controlador.pausar_proceso, state=tk.DISABLED)
        self.btn_pausar.grid(row=0, column=1, padx=5)
        
        self.btn_reanudar = ttk.Button(control_frame, text="Reanudar (F3)", command=self.controlador.reanudar_proceso, state=tk.DISABLED)
        self.btn_reanudar.grid(row=0, column=2, padx=5)
        
        self.btn_detener = ttk.Button(control_frame, text="Detener (F4)", command=self.controlador.detener_proceso, state=tk.DISABLED)
        self.btn_detener.grid(row=0, column=3, padx=5)
        
        # Botones adicionales
        extra_frame = ttk.Frame(main_frame)
        extra_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        ttk.Button(extra_frame, text="Escribir PRUEBA A", command=self.controlador.escribir_prueba_a).grid(row=0, column=0, padx=5)
        ttk.Button(extra_frame, text="Configurar KML", command=self.controlador.configurar_kml).grid(row=0, column=1, padx=5)
        
        # Estado
        estado_frame = ttk.LabelFrame(main_frame, text="Estado", padding="5")
        estado_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.estado_label = ttk.Label(estado_frame, textvariable=self.estado, foreground="blue")
        self.estado_label.grid(row=0, column=0, sticky=tk.W)
        
        # Log
        log_frame = ttk.LabelFrame(main_frame, text="Log de Ejecución", padding="5")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=15, width=80)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar peso para expansión
        main_frame.rowconfigure(5, weight=1)
    
    def log(self, mensaje):
        """Agregar mensaje al log"""
        self.log_text.insert(tk.END, f"{mensaje}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def clear_log(self):
        """Limpiar el log"""
        self.log_text.delete(1.0, tk.END)
    
    def actualizar_estado_botones(self, estado_programa):
        """Actualizar estado de los botones según el estado del proceso"""
        if estado_programa.ejecutando:
            self.btn_iniciar.config(state=tk.DISABLED)
            self.btn_detener.config(state=tk.NORMAL)
            
            if estado_programa.pausado:
                self.btn_pausar.config(state=tk.DISABLED)
                self.btn_reanudar.config(state=tk.NORMAL)
            else:
                self.btn_pausar.config(state=tk.NORMAL)
                self.btn_reanudar.config(state=tk.DISABLED)
        else:
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_pausar.config(state=tk.DISABLED)
            self.btn_reanudar.config(state=tk.DISABLED)
            self.btn_detener.config(state=tk.DISABLED)
    
    def actualizar_estado_lineas(self, estado_programa):
        """Actualizar display de líneas actuales y restantes"""
        self.linea_actual.set(str(estado_programa.linea_actual))
        lineas_rest = max(0, estado_programa.linea_maxima - estado_programa.linea_actual)
        self.lineas_restantes.set(str(lineas_rest))
    
    def actualizar_estado_general(self, estado_programa):
        """Actualizar estado general del programa"""
        self.estado.set(estado_programa.estado)
        
        # Cambiar color según estado
        if estado_programa.estado == "Ejecutando...":
            self.estado_label.configure(foreground="green")
        elif estado_programa.estado == "Pausado":
            self.estado_label.configure(foreground="orange")
        elif estado_programa.estado == "Detenido":
            self.estado_label.configure(foreground="red")
        elif estado_programa.estado == "Completado":
            self.estado_label.configure(foreground="green")
        else:
            self.estado_label.configure(foreground="blue")
    
    def mostrar_mensaje(self, titulo, mensaje, tipo="info"):
        """Mostrar mensaje al usuario"""
        if tipo == "info":
            messagebox.showinfo(titulo, mensaje)
        elif tipo == "error":
            messagebox.showerror(titulo, mensaje)
        elif tipo == "warning":
            messagebox.showwarning(titulo, mensaje)
    
    def pedir_configuracion_kml(self, valor_actual):
        """Pedir configuración de nombre KML"""
        return simpledialog.askstring(
            "Configurar KML", 
            "Ingrese el nuevo nombre para archivos KML:",
            initialvalue=valor_actual
        )
    
    def pedir_seleccion_csv(self):
        """Pedir selección de archivo CSV"""
        return filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )