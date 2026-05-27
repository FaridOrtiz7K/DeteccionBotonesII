import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import threading
import time
import pandas as pd
import os

class InterfazAutomation:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("Sistema de Automatización NSE/GE")
        self.root.geometry("900x700")

        # Variables de la vista (vinculadas a widgets)
        self.csv_file = tk.StringVar()
        self.linea_maxima = tk.IntVar(value=1)
        self.estado = tk.StringVar(value="Listo")
        self.linea_actual = tk.StringVar(value="0")
        self.lineas_restantes = tk.StringVar(value="0")
        self.lote_inicio = tk.IntVar(value=1)
        self.lote_fin = tk.IntVar(value=1)
        self.id_consulta = tk.StringVar()
        self.info_lote = tk.StringVar(value="Información del lote aparecerá aquí")

        # Referencia a la ventana de pausa
        self.pause_window = None

        # Configurar UI
        self.setup_ui()

    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configurar pesos de grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Sección CSV
        csv_frame = ttk.LabelFrame(main_frame, text="Configuración CSV", padding="5")
        csv_frame.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        csv_frame.columnconfigure(1, weight=1)

        ttk.Label(csv_frame, text="Archivo CSV:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Entry(csv_frame, textvariable=self.csv_file, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(csv_frame, text="Seleccionar", command=self.seleccionar_csv).grid(row=0, column=2, padx=5)

        # Sección consulta ID
        ttk.Label(csv_frame, text="Consultar ID:").grid(row=0, column=3, sticky=tk.W, padx=5)
        ttk.Entry(csv_frame, textvariable=self.id_consulta, width=15).grid(row=0, column=4, padx=5)
        ttk.Button(csv_frame, text="Buscar", command=self.consultar_id).grid(row=0, column=5, padx=5)

        # Sección control por lotes
        lote_frame = ttk.LabelFrame(main_frame, text="Control por Lotes", padding="5")
        lote_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(lote_frame, text="Lote inicio:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Spinbox(lote_frame, from_=1, to=10000, textvariable=self.lote_inicio, width=10).grid(row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(lote_frame, text="Lote fin:").grid(row=0, column=2, sticky=tk.W, padx=5)
        ttk.Spinbox(lote_frame, from_=1, to=10000, textvariable=self.lote_fin, width=10).grid(row=0, column=3, sticky=tk.W, padx=5)

        # Información del lote actual
        info_frame = ttk.LabelFrame(main_frame, text="Información del Lote Actual", padding="5")
        info_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

        self.info_label = ttk.Label(info_frame, textvariable=self.info_lote, wraplength=800)
        self.info_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)

        # Sección ejecución
        exec_frame = ttk.LabelFrame(main_frame, text="Control de Ejecución", padding="5")
        exec_frame.grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(exec_frame, text="Línea actual:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Label(exec_frame, textvariable=self.linea_actual).grid(row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(exec_frame, text="Líneas restantes:").grid(row=0, column=2, sticky=tk.W, padx=5)
        ttk.Label(exec_frame, textvariable=self.lineas_restantes).grid(row=0, column=3, sticky=tk.W, padx=5)

        # Botones de control
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=4, column=0, columnspan=4, pady=10)

        self.btn_iniciar = ttk.Button(control_frame, text="Iniciar Proceso", command=self.iniciar_proceso)
        self.btn_iniciar.grid(row=0, column=0, padx=5)

        self.btn_pausar = ttk.Button(control_frame, text="Pausar (F2/ESC)", command=self.pausar_proceso, state=tk.DISABLED)
        self.btn_pausar.grid(row=0, column=1, padx=5)

        self.btn_reanudar = ttk.Button(control_frame, text="Reanudar (F3)", command=self.reanudar_proceso, state=tk.DISABLED)
        self.btn_reanudar.grid(row=0, column=2, padx=5)

        self.btn_detener = ttk.Button(control_frame, text="Detener Inmediato (F4)", command=self.detener_proceso, state=tk.DISABLED)
        self.btn_detener.grid(row=0, column=3, padx=5)

        # Botones adicionales
        extra_frame = ttk.Frame(main_frame)
        extra_frame.grid(row=5, column=0, columnspan=4, pady=10)

        ttk.Button(extra_frame, text="Escribir PRUEBA A", command=self.escribir_prueba_a).grid(row=0, column=0, padx=5)
        ttk.Button(extra_frame, text="Configurar KML", command=self.configurar_kml).grid(row=0, column=1, padx=5)

        # Estado
        estado_frame = ttk.LabelFrame(main_frame, text="Estado", padding="5")
        estado_frame.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

        self.estado_label = ttk.Label(estado_frame, textvariable=self.estado, foreground="blue")
        self.estado_label.grid(row=0, column=0, sticky=tk.W)

        # Log
        log_frame = ttk.LabelFrame(main_frame, text="Log de Ejecución", padding="5")
        log_frame.grid(row=7, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=90)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        main_frame.rowconfigure(7, weight=1)

    # Métodos que llaman al controlador
    def seleccionar_csv(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if archivo:
            self.csv_file.set(archivo)
            self.controller.seleccionar_csv(archivo)

    def consultar_id(self):
        id_buscar = self.id_consulta.get().strip()
        self.controller.consultar_id(id_buscar)

    def iniciar_proceso(self):
        # Obtener valores actuales de los spinbox
        lote_inicio = self.lote_inicio.get()
        lote_fin = self.lote_fin.get()
        archivo = self.csv_file.get()
        self.controller.iniciar_proceso(archivo, lote_inicio, lote_fin)

    def pausar_proceso(self):
        self.controller.pausar_proceso()

    def reanudar_proceso(self):
        self.controller.reanudar_proceso()

    def detener_proceso(self):
        self.controller.detener_proceso()

    def escribir_prueba_a(self):
        self.controller.escribir_prueba_a()

    def configurar_kml(self):
        from tkinter import simpledialog
        nuevo_nombre = simpledialog.askstring(
            "Configurar KML",
            "Ingrese el nuevo nombre para archivos KML:",
            initialvalue=self.controller.kml_filename  # Asumiendo que el controlador expone kml_filename
        )
        if nuevo_nombre:
            self.controller.configurar_kml(nuevo_nombre)

    # Métodos para que el controlador actualice la UI (llamados con root.after)
    def log(self, mensaje):
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {mensaje}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def actualizar_estado_botones(self, ejecutando, pausado, en_cuenta_regresiva=False):
        if en_cuenta_regresiva:
            self.btn_iniciar.config(state=tk.DISABLED)
            self.btn_pausar.config(state=tk.DISABLED)
            self.btn_reanudar.config(state=tk.DISABLED)
            self.btn_detener.config(state=tk.NORMAL)
            return

        if ejecutando:
            self.btn_iniciar.config(state=tk.DISABLED)
            self.btn_detener.config(state=tk.NORMAL)
            if pausado:
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

    def mostrar_ventana_pausa(self, current_lote, total_lotes):
        from views.pause_window import PauseWindow
        if self.pause_window is None or not self.pause_window.winfo_exists():
            self.pause_window = PauseWindow(self.root, self.controller, current_lote, total_lotes)

    def cerrar_ventana_pausa(self):
        if self.pause_window is not None and self.pause_window.winfo_exists():
            self.pause_window.destroy()
            self.pause_window = None