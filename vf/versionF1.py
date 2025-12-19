import os
import time
import sys
import pandas as pd
import pyautogui
import cv2
import numpy as np
import logging
import threading
from PIL import ImageGrab
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, scrolledtext
import keyboard
from utils.ahk_managerCopyDelete import AHKManagerCD
from utils.ahk_writer import AHKWriter
from utils.ahk_click_down import AHKClickDown
from utils.ahk_enter import EnterAHKManager
from utils.ahk_manager import AHKManager

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('nse_automation.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Variables globales
CSV_FILE = ""
KML_FILENAME = "NN"
HILO_EJECUCION = None
LINEA_ACTUAL = 0
LINEA_MAXIMA = 0
LOTE_INICIO = 1
LOTE_FIN = 1
INFO_LOTE_ACTUAL = {}

# Clase para manejar estado de ejecución con threading.Condition
class EstadoEjecucion:
    def __init__(self):
        self.ejecutando = False
        self.pausado = False
        self.detener_inmediato = False
        self.linea_en_proceso = False
        self.en_cuenta_regresiva = False
        self.condition = threading.Condition()
    
    def set_ejecutando(self, valor):
        with self.condition:
            self.ejecutando = valor
            if not valor:
                self.pausado = False
                self.linea_en_proceso = False
                self.en_cuenta_regresiva = False
            self.condition.notify_all()
    
    def set_pausado(self, valor):
        with self.condition:
            self.pausado = valor
            self.condition.notify_all()
    
    def set_detener_inmediato(self, valor):
        with self.condition:
            self.detener_inmediato = valor
            if valor:
                self.ejecutando = False
                self.pausado = False
                self.en_cuenta_regresiva = False
            self.condition.notify_all()
    
    def set_en_cuenta_regresiva(self, valor):
        with self.condition:
            self.en_cuenta_regresiva = valor
            self.condition.notify_all()
    
    def set_linea_en_proceso(self, valor):
        with self.condition:
            self.linea_en_proceso = valor
    
    def esperar_si_pausado(self):
        with self.condition:
            while (self.pausado or self.en_cuenta_regresiva) and self.ejecutando and not self.detener_inmediato:
                self.condition.wait()
            return not self.ejecutando or self.detener_inmediato
    
    def verificar_continuar(self):
        with self.condition:
            return self.ejecutando and not self.detener_inmediato and not self.en_cuenta_regresiva

# Instancia global de estado
estado_global = EstadoEjecucion()

# Clase para la ventana emergente de pausa
class PauseWindow(tk.Toplevel):
    def __init__(self, parent, controller, current_lote, total_lotes):
        super().__init__(parent)
        self.controller = controller
        self.title("Proceso en Pausa")
        self.geometry("300x200")
        self.resizable(False, False)
        
        # Hacer la ventana modal
        self.transient(parent)
        self.grab_set()
        
        # Centrar la ventana
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")
        
        # Hacer que la ventana esté siempre encima
        self.attributes("-topmost", True)
        
        # Configurar el protocolo de cierre
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Variables
        self.countdown = 5
        self.countdown_running = False
        
        # Crear widgets
        self.create_widgets(current_lote, total_lotes)
    
    def create_widgets(self, current_lote, total_lotes):
        # Frame principal
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Información de progreso
        lotes_faltantes = total_lotes - current_lote + 1
        ttk.Label(main_frame, text=f"Lote actual: {current_lote}").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Label(main_frame, text=f"Lotes completados: {current_lote - 1}").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Label(main_frame, text=f"Lotes faltantes: {lotes_faltantes}").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Label(main_frame, text=f"Total de lotes: {total_lotes}").grid(row=3, column=0, sticky=tk.W, pady=5)
        
        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, pady=10)
        
        self.resume_button = ttk.Button(button_frame, text="Reanudar (5)", 
                                       command=self.start_countdown)
        self.resume_button.grid(row=0, column=0, padx=5)
        
        self.exit_button = ttk.Button(button_frame, text="Salir", 
                                     command=self.controller.detener_proceso)
        self.exit_button.grid(row=0, column=1, padx=5)
        
        # Configurar expansión
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
    
    def start_countdown(self):
        """Inicia la cuenta regresiva para reanudar"""
        if not self.countdown_running:
            self.countdown_running = True
            self.resume_button.config(state=tk.DISABLED)
            self.update_countdown()
    
    def update_countdown(self):
        """Actualiza la cuenta regresiva"""
        if self.countdown > 0 and self.countdown_running:
            self.resume_button.config(text=f"Reanudar ({self.countdown})")
            self.countdown -= 1
            self.after(1000, self.update_countdown)
        elif self.countdown_running:
            logger.info("Reanudando proceso desde ventana de pausa")
            self.controller.reanudar_desde_ventana()
            self.destroy()
    
    def on_close(self):
        """Maneja el cierre de la ventana"""
        self.countdown_running = False
        self.destroy()

class InterfazAutomation:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Automatización NSE/GE")
        self.root.geometry("900x700")
        
        # Variables
        self.csv_file = tk.StringVar()
        self.linea_maxima = tk.IntVar(value=1)
        self.estado = tk.StringVar(value="Listo")
        self.linea_actual = tk.StringVar(value="0")
        self.lineas_restantes = tk.StringVar(value="0")
        self.lote_inicio = tk.IntVar(value=1)
        self.lote_fin = tk.IntVar(value=1)
        self.id_consulta = tk.StringVar()
        self.info_lote = tk.StringVar(value="Información del lote aparecerá aquí")
        
        # Referencia a ventana de pausa
        self.pause_window = None
        
        self.setup_ui()
        self.setup_bindings()
        
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
        
    def setup_bindings(self):
        keyboard.add_hotkey('esc', self.pausar_proceso)
        keyboard.add_hotkey('f2', self.pausar_proceso)
        keyboard.add_hotkey('f3', self.reanudar_proceso)
        keyboard.add_hotkey('f4', self.detener_proceso)
        
    def log(self, mensaje):
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {mensaje}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def seleccionar_csv(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if archivo:
            self.csv_file.set(archivo)
            global CSV_FILE
            CSV_FILE = archivo
            self.log(f"CSV seleccionado: {archivo}")
            
            try:
                df = pd.read_csv(archivo)
                self.linea_maxima.set(len(df))
                self.lote_fin.set(len(df))
                self.actualizar_estado_lineas()
                self.log(f"CSV cargado: {len(df)} registros encontrados")
            except Exception as e:
                self.log(f"Error al leer CSV: {e}")
    
    def consultar_id(self):
        if not self.csv_file.get():
            messagebox.showerror("Error", "Primero seleccione un archivo CSV")
            return
            
        id_buscar = self.id_consulta.get()
        if not id_buscar:
            messagebox.showwarning("Advertencia", "Ingrese un ID para consultar")
            return
            
        try:
            df = pd.read_csv(self.csv_file.get())
            resultado = df[df.iloc[:, 0].astype(str) == str(id_buscar)]
            
            if len(resultado) == 0:
                messagebox.showinfo("Resultado", f"ID {id_buscar} no encontrado en el CSV")
            else:
                info = f"ID {id_buscar} encontrado:\n"
                for idx, col in enumerate(df.columns):
                    valor = resultado.iloc[0, idx]
                    if pd.isna(valor):
                        valor = ""
                    info += f"{col}: {valor}\n"
                
                if len(df.columns) > 0:
                    ultima_col = df.columns[-1]
                    valor_ultimo = resultado.iloc[0, -1]
                    if pd.isna(valor_ultimo):
                        valor_ultimo = ""
                    info += f"\n--- Resumen del Lote (última columna) ---\n"
                    info += f"{ultima_col}: {valor_ultimo}"
                
                messagebox.showinfo("Información del ID", info)
                self.log(f"ID {id_buscar} consultado: {len(resultado)} coincidencias")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al consultar ID: {e}")
    
    def actualizar_info_lote(self, linea_actual, datos):
        global INFO_LOTE_ACTUAL
        INFO_LOTE_ACTUAL = {
            'linea': linea_actual,
            'datos': datos if datos is not None else {},
            'timestamp': time.time()
        }
        
        if datos is not None and len(datos) > 0:
            id_valor = datos.iloc[0] if len(datos) > 0 else ''
            if pd.isna(id_valor):
                id_valor = ""
            
            info_text = f"Línea {linea_actual}: ID={id_valor}"
            
            if len(datos) > 1:
                col2_valor = datos.iloc[1]
                if pd.isna(col2_valor):
                    col2_valor = ""
                if col2_valor:
                    info_text += f", Col2={col2_valor}"
            
            if len(datos) > 3:
                col4_valor = datos.iloc[3]
                if pd.isna(col4_valor):
                    col4_valor = ""
                if col4_valor:
                    info_text += f", Col4={col4_valor}"
            
            if len(datos) > 0:
                ultima_col_valor = datos.iloc[-1]
                if pd.isna(ultima_col_valor):
                    ultima_col_valor = ""
                if ultima_col_valor:
                    info_text += f", Resumen={ultima_col_valor}"
            
            self.info_lote.set(info_text)
        else:
            self.info_lote.set(f"Línea {linea_actual}: Datos vacíos o no encontrados")
    
    def escribir_prueba_a(self):
        if not self.csv_file.get():
            messagebox.showerror("Error", "Primero seleccione un archivo CSV")
            return
            
        try:
            df = pd.read_csv(self.csv_file.get())
            if len(df) == 0:
                messagebox.showerror("Error", "El CSV está vacío")
                return
                
            ultima_columna = df.iloc[0, -1]
            if pd.isna(ultima_columna):
                texto_a_escribir = ""
            else:
                texto_a_escribir = str(ultima_columna)
            
            self.log(f"Escribiendo: '{texto_a_escribir}'")
            self.log("⚠️ Coloque el cursor en la posición deseada - escribiendo en 5 segundos...")
            
            for i in range(5, 0, -1):
                self.estado.set(f"Escribiendo en {i} segundos...")
                time.sleep(1)
            
            x, y = pyautogui.position()
            self.log(f"Posición del cursor: ({x}, {y})")
            
            ahk_writer = AHKWriter()
            if ahk_writer.start_ahk():
                exito = ahk_writer.ejecutar_escritura_ahk(x, y, texto_a_escribir)
                ahk_writer.stop_ahk()
                
                if exito:
                    self.log("✅ Texto escrito exitosamente")
                    self.estado.set("Listo")
                else:
                    self.log("❌ Error al escribir texto")
                    self.estado.set("Error")
            else:
                self.log("❌ No se pudo iniciar AHKWriter")
                self.estado.set("Error")
                
        except Exception as e:
            self.log(f"❌ Error al escribir PRUEBA A: {e}")
            self.estado.set("Error")
    
    def configurar_kml(self):
        global KML_FILENAME
        nuevo_nombre = simpledialog.askstring(
            "Configurar KML", 
            "Ingrese el nuevo nombre para archivos KML:",
            initialvalue=KML_FILENAME
        )
        if nuevo_nombre:
            KML_FILENAME = nuevo_nombre
            self.log(f"Nombre KML configurado a: {KML_FILENAME}")
    
    def actualizar_estado_lineas(self):
        global LINEA_ACTUAL, LINEA_MAXIMA
        self.linea_actual.set(str(LINEA_ACTUAL))
        lineas_rest = max(0, LINEA_MAXIMA - LINEA_ACTUAL)
        self.lineas_restantes.set(str(lineas_rest))
    
    def mostrar_estado_actual(self):
        if estado_global.ejecutando:
            lineas_restantes = LINEA_MAXIMA - LINEA_ACTUAL
            estado_linea = "EN PROCESO" if estado_global.linea_en_proceso else "ESPERANDO"
            mensaje = f"Línea actual: {LINEA_ACTUAL}\nLíneas restantes: {lineas_restantes}\nEstado línea: {estado_linea}"
            
            if INFO_LOTE_ACTUAL:
                mensaje += f"\n\nInfo Lote Actual:\n{self.info_lote.get()}"
            
            messagebox.showinfo("Estado Actual", mensaje)
    
    def guardar_progreso_manual(self):
        try:
            self.log("💾 Guardando progreso manualmente...")
            pyautogui.hotkey('ctrl', 's')
            time.sleep(6)
            self.log("✅ Progreso guardado exitosamente")
            return True
        except Exception as e:
            self.log(f"❌ Error al guardar progreso: {e}")
            return False
    
    def iniciar_proceso(self):
        global CSV_FILE, LOTE_INICIO, LOTE_FIN, LINEA_ACTUAL, LINEA_MAXIMA
        
        if not self.csv_file.get():
            messagebox.showerror("Error", "Seleccione un archivo CSV primero")
            return
            
        CSV_FILE = self.csv_file.get()
        LOTE_INICIO = self.lote_inicio.get()
        LOTE_FIN = self.lote_fin.get()
        
        if LOTE_INICIO > LOTE_FIN:
            messagebox.showerror("Error", "El lote inicio no puede ser mayor que el lote fin")
            return
            
        try:
            df = pd.read_csv(CSV_FILE)
            if len(df) == 0:
                self.log("⚠️ CSV está vacío. Continuando sin procesar datos...")
                LINEA_MAXIMA = 0
                LINEA_ACTUAL = 0
            else:
                LINEA_MAXIMA = min(LOTE_FIN, len(df))
                LINEA_ACTUAL = max(1, LOTE_INICIO)
                
                if LINEA_ACTUAL > LINEA_MAXIMA:
                    messagebox.showerror("Error", "El lote inicio está fuera del rango disponible")
                    return
        except Exception as e:
            self.log(f"❌ Error al leer CSV: {e}")
            return
        
        estado_global.set_ejecutando(True)
        estado_global.set_pausado(False)
        estado_global.set_detener_inmediato(False)
        estado_global.set_linea_en_proceso(False)
        
        self.actualizar_estado_botones()
        self.estado.set("Ejecutando...")
        self.estado_label.configure(foreground="green")
        
        self.hilo_ejecucion = threading.Thread(target=self.ejecutar_procesos)
        self.hilo_ejecucion.daemon = True
        self.hilo_ejecucion.start()
    
    def pausar_proceso(self):
        if estado_global.ejecutando and not estado_global.pausado:
            if estado_global.en_cuenta_regresiva:
                estado_global.set_en_cuenta_regresiva(False)
                self.log("⏸️ Cuenta regresiva cancelada - Proceso permanece pausado")
            else:
                estado_global.set_pausado(True)
                self.estado.set("Pausado")
                self.estado_label.configure(foreground="orange")
                self.log("⏸️ Proceso pausado")
                
                # Mostrar ventana emergente de pausa
                if self.pause_window is None or not self.pause_window.winfo_exists():
                    self.pause_window = PauseWindow(
                        self.root, 
                        self, 
                        LINEA_ACTUAL,
                        LINEA_MAXIMA
                    )
                
                self.actualizar_estado_botones()
    
    def reanudar_desde_ventana(self):
        """Método para reanudar desde la ventana de pausa"""
        estado_global.set_pausado(False)
        self.estado.set("Ejecutando...")
        self.estado_label.configure(foreground="green")
        self.log("▶️ Proceso reanudado desde ventana de pausa")
        self.actualizar_estado_botones()
    
    def reanudar_proceso(self):
        if estado_global.ejecutando and estado_global.pausado:
            # Si hay ventana de pausa abierta, cerrarla
            if self.pause_window is not None and self.pause_window.winfo_exists():
                self.pause_window.destroy()
                self.pause_window = None
            
            self.btn_reanudar.config(state=tk.DISABLED)
            estado_global.set_en_cuenta_regresiva(True)
            
            self.estado.set("Reanudando en 3 segundos...")
            self.estado_label.configure(foreground="blue")
            self.log("Reanudando en 3 segundos...")
            
            threading.Thread(target=self._cuenta_regresiva_reanudacion, daemon=True).start()

    def _cuenta_regresiva_reanudacion(self):
        try:
            for i in range(3, 0, -1):
                if estado_global.detener_inmediato:
                    self.root.after(0, self._cancelar_reanudacion, "Reanudación cancelada (proceso detenido)")
                    return
                    
                self.root.after(0, lambda x=i: self.estado.set(f"Reanudando en {x} segundos..."))
                self.root.after(0, self.root.update)
                time.sleep(1)
            
            if not estado_global.detener_inmediato:
                self.root.after(0, self._completar_reanudacion)
            else:
                self.root.after(0, self._cancelar_reanudacion, 
                            "Reanudación cancelada (proceso detenido durante la cuenta regresiva)")
        except Exception as e:
            self.root.after(0, self._cancelar_reanudacion, f"Error en cuenta regresiva: {e}")

    def _completar_reanudacion(self):
        estado_global.set_en_cuenta_regresiva(False)
        estado_global.set_pausado(False)
        self.estado.set("Ejecutando...")
        self.estado_label.configure(foreground="green")
        self.log("▶️ Proceso reanudado")
        self.actualizar_estado_botones()

    def _cancelar_reanudacion(self, mensaje):
        estado_global.set_en_cuenta_regresiva(False)
        self.log(mensaje)
        self.actualizar_estado_botones()
        
    def detener_proceso(self):
        # Cerrar ventana de pausa si está abierta
        if self.pause_window is not None and self.pause_window.winfo_exists():
            self.pause_window.destroy()
            self.pause_window = None
        
        estado_global.set_detener_inmediato(True)
        estado_global.set_ejecutando(False)
        
        self.estado.set("Detenido")
        self.estado_label.configure(foreground="red")
        self.log("⏹️ Proceso detenido inmediatamente")
        self.actualizar_estado_botones()
        self.actualizar_estado_lineas()
    
    def actualizar_estado_botones(self):
        estado_actual = self.estado.get()
        if "Reanudando en" in estado_actual:
            self.btn_iniciar.config(state=tk.DISABLED)
            self.btn_pausar.config(state=tk.DISABLED)
            self.btn_reanudar.config(state=tk.DISABLED)
            self.btn_detener.config(state=tk.NORMAL)
            return
        
        if estado_global.ejecutando:
            self.btn_iniciar.config(state=tk.DISABLED)
            self.btn_detener.config(state=tk.NORMAL)
            
            if estado_global.pausado:
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
    
    def ejecutar_procesos(self):
        global LINEA_ACTUAL, LINEA_MAXIMA, KML_FILENAME
        
        lotes_desde_ultimo_guardado = 0
        
        try:
            while LINEA_ACTUAL <= LINEA_MAXIMA and estado_global.verificar_continuar():
                if estado_global.esperar_si_pausado():
                    break
                
                estado_global.set_linea_en_proceso(True)
                    
                self.log(f"🔄 Procesando línea {LINEA_ACTUAL}/{LINEA_MAXIMA}")
                self.actualizar_estado_lineas()
                
                self.log("Iniciando Programa 1 - Procesador CSV")
                resultado1, linea_procesada, datos_lote = ejecutar_programa1_interfaz(LINEA_ACTUAL, self.log)
                
                if not estado_global.verificar_continuar():
                    break
                    
                self.actualizar_info_lote(LINEA_ACTUAL, datos_lote)
                
                if not resultado1 or linea_procesada is None:
                    self.log(f"⚠️ Programa 1 no procesó la línea {LINEA_ACTUAL} (ID no encontrado después de 2 intentos). Saltando al siguiente lote...")
                    estado_global.set_linea_en_proceso(False)
                    LINEA_ACTUAL += 1
                    lotes_desde_ultimo_guardado += 1
                    continue
                
                for _ in range(3):
                    if estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)
                
                if not estado_global.verificar_continuar():
                    break
                    
                self.log("Iniciando Programa 2 - Automatización NSE")
                resultado2 = ejecutar_programa2_interfaz(linea_procesada, self.log)
                
                if not resultado2:
                    self.log(f"❌ Programa 2 falló en línea {LINEA_ACTUAL}")
                    estado_global.set_linea_en_proceso(False)
                    LINEA_ACTUAL += 1
                    lotes_desde_ultimo_guardado += 1
                    continue
                
                for _ in range(3):
                    if estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)
                
                if not estado_global.verificar_continuar():
                    break
                    
                self.log("Iniciando Programa 3 - Servicios NSE")
                resultado3 = ejecutar_programa3_interfaz(linea_procesada, self.log)
                
                if not resultado3:
                    self.log(f"❌ Programa 3 falló en línea {LINEA_ACTUAL}")
                    estado_global.set_linea_en_proceso(False)
                    LINEA_ACTUAL += 1
                    lotes_desde_ultimo_guardado += 1
                    continue
                
                for _ in range(3):
                    if estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)
                
                if not estado_global.verificar_continuar():
                    break
                    
                self.log("Iniciando Programa 4 - Automatización GE")
                resultado4 = ejecutar_programa4_interfaz(linea_procesada, KML_FILENAME, self.log)
                
                if resultado4:
                    self.log(f"✅ Línea {LINEA_ACTUAL} procesada exitosamente")
                else:
                    self.log(f"⚠️ Línea {LINEA_ACTUAL} completada con advertencias")
                
                estado_global.set_linea_en_proceso(False)
                lotes_desde_ultimo_guardado += 1
                
                if lotes_desde_ultimo_guardado >= 10:
                    self.log("📁 Guardando progreso después de 10 lotes...")
                    try:
                        pyautogui.hotkey('ctrl', 's')
                        for _ in range(6):
                            if estado_global.esperar_si_pausado():
                                break
                            time.sleep(1)
                        self.log("✅ Progreso guardado exitosamente")
                        lotes_desde_ultimo_guardado = 0
                    except Exception as e:
                        self.log(f"⚠️ Error al guardar progreso: {e}")
                
                LINEA_ACTUAL += 1
                self.actualizar_estado_lineas()
                
                for _ in range(4):
                    if estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)
            
            if not estado_global.detener_inmediato and LINEA_ACTUAL > LINEA_MAXIMA:
                self.log("📁 Guardando progreso final al completar todos los lotes...")
                try:
                    pyautogui.hotkey('ctrl', 's')
                    time.sleep(6)
                    self.log("✅ Progreso final guardado exitosamente")
                except Exception as e:
                    self.log(f"⚠️ Error al guardar progreso final: {e}")
            
            if estado_global.detener_inmediato:
                self.log("🛑 Proceso detenido inmediatamente por usuario")
                self.estado.set("Detenido")
                self.estado_label.configure(foreground="red")
            elif estado_global.ejecutando and LINEA_ACTUAL > LINEA_MAXIMA:
                self.log("🎉 Proceso completado exitosamente")
                self.estado.set("Completado")
                self.estado_label.configure(foreground="green")
            elif not estado_global.ejecutando:
                self.log("Proceso detenido por el usuario")
                
        except Exception as e:
            self.log(f"❌ Error en ejecución: {e}")
            self.estado.set("Error")
            self.estado_label.configure(foreground="red")
            estado_global.set_linea_en_proceso(False)
        
        finally:
            estado_global.set_ejecutando(False)
            self.actualizar_estado_botones()
            # Cerrar ventana de pausa si está abierta
            if self.pause_window is not None and self.pause_window.winfo_exists():
                self.pause_window.destroy()
                self.pause_window = None

def ejecutar_programa1_interfaz(linea_especifica, log_func):
    try:
        log_func("=" * 60)
        log_func("INICIANDO PROGRAMA 1 - PROCESADOR CSV")
        log_func("=" * 60)
        
        archivo_csv = CSV_FILE
        procesador = ProcesadorCSV(archivo_csv)
        
        log_func("Iniciando procesamiento automático del Programa 1...")
        
        
        if estado_global.esperar_si_pausado():
            return False, linea_especifica, None
        time.sleep(0.5)
        
        if not procesador.cargar_csv():
            log_func("⚠️ CSV vacío detectado. Continuando sin procesar...")
            return True, None, None
        
        if linea_especifica > len(procesador.df):
            log_func(f"⚠️ Línea {linea_especifica} no existe en CSV. Saltando...")
            return True, None, None
        
        procesador.df = procesador.df.iloc[linea_especifica-1:linea_especifica]
        datos_lote = procesador.df.iloc[0] if len(procesador.df) > 0 else None
        
        resultado, linea_procesada = procesador.procesar_todo()
        
        if resultado and linea_procesada:
            log_func(f"✅ Programa 1 completado. Línea procesada: {linea_procesada}")
            return True, linea_procesada, datos_lote
        else:
            log_func("❌ Programa 1 falló o no encontró ID")
            return False, None, datos_lote
            
    except Exception as e:
        log_func(f"❌ Error en Programa 1: {e}")
        return False, None, None

def ejecutar_programa2_interfaz(linea_especifica, log_func):
    try:
        log_func("\n" + "=" * 60)
        log_func("INICIANDO PROGRAMA 2 - AUTOMATIZACIÓN NSE")
        log_func("=" * 60)
        
        nse = NSEAutomation(linea_especifica=linea_especifica)
        nse.is_running = True
        
        if not os.path.exists(nse.csv_file):
            log_func(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
            return True
        
        try:
            df = pd.read_csv(nse.csv_file)
            if len(df) == 0:
                log_func("⚠️ CSV vacío. Saltando Programa 2...")
                return True
        except:
            pass
        
        log_func(f"🎯 Procesando línea: {linea_especifica}")
        
        for _ in range(1):
            if estado_global.esperar_si_pausado():
                return False
            time.sleep(0.5)
        
        resultado = nse.execute_nse_script()
        
        if resultado:
            log_func("✅ Programa 2 finalizado exitosamente")
        else:
            log_func("⚠️ Programa 2 completado con advertencias")
            
        return resultado
        
    except Exception as e:
        log_func(f"❌ Error en Programa 2: {e}")
        return False

def ejecutar_programa3_interfaz(linea_especifica, log_func):
    try:
        log_func("\n" + "=" * 60)
        log_func("INICIANDO PROGRAMA 3 - SERVICIOS NSE")
        log_func("=" * 60)
        
        nse_services = NSEServicesAutomation(linea_especifica=linea_especifica)
        
        if not os.path.exists(nse_services.csv_file):
            log_func(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
            return True
        
        log_func(f"🎯 Procesando línea: {linea_especifica}")
        
        if not nse_services.iniciar_ahk():
            log_func("⚠️ No se pudieron iniciar los servicios AHK. Continuando...")
            return True
        
        nse_services.is_running = True
        resultado = nse_services.procesar_linea_especifica()
        
        if resultado:
            log_func(f"✅ Programa 3 completado exitosamente")
        else:
            log_func(f"⚠️ Programa 3 completado con advertencias")
        
        nse_services.detener_ahk()
        return resultado
        
    except Exception as e:
        log_func(f"❌ Error en Programa 3: {e}")
        return False

def ejecutar_programa4_interfaz(linea_especifica, kml_filename, log_func):
    try:
        log_func("\n" + "=" * 60)
        log_func("INICIANDO PROGRAMA 4 - AUTOMATIZACIÓN GE")
        log_func("=" * 60)
        
        ge_auto = GEAutomation(linea_especifica=linea_especifica)
        ge_auto.is_running = True
        
        ge_auto.nombre = f"{kml_filename}.kml"
        
        if not os.path.exists(ge_auto.csv_file):
            log_func(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
            return True
        
        try:
            df = pd.read_csv(ge_auto.csv_file)
            if len(df) == 0:
                log_func("⚠️ CSV vacío. Saltando Programa 4...")
                return True
        except:
            pass
        
        log_func(f"🎯 Procesando línea: {linea_especifica}")
        log_func(f"📁 Archivo KML: {ge_auto.nombre}")
        
        for _ in range(1):
            if estado_global.esperar_si_pausado():
                return False
            time.sleep(0.5)
        
        success = ge_auto.perform_actions()
        
        if success:
            log_func("✅ Programa 4 finalizado exitosamente")
        else:
            log_func("⚠️ Programa 4 completado con advertencias")
            
        return success
        
    except Exception as e:
        log_func(f"❌ Error en Programa 4: {e}")
        return False

class ProcesadorCSV:
    def __init__(self, archivo_csv):
        self.archivo_csv = archivo_csv
        self.df = None
        self.ahk_manager = AHKManagerCD()
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        
    def cargar_csv(self):
        try:
            self.df = pd.read_csv(self.archivo_csv)
            if len(self.df) == 0:
                logger.info("CSV vacío detectado")
                return True
            logger.info(f"CSV cargado correctamente: {len(self.df)} registros")
            return True
        except Exception as e:
            logger.error(f"Error cargando CSV: {e}")
            return False
    
    def iniciar_ahk(self):
        logger.info("Iniciando procesos AHK...")
        time.sleep(1.5)
        return (self.ahk_manager.start_ahk() and 
                self.ahk_writer.start_ahk() and 
                self.ahk_click_down.start_ahk())
    
    def detener_ahk(self):
        logger.info("Deteniendo procesos AHK...")
        self.ahk_manager.stop_ahk()
        self.ahk_writer.stop_ahk()
        self.ahk_click_down.stop_ahk()
        time.sleep(1.5)
    
    def buscar_por_id(self, id_buscar, max_intentos=2):
        if self.df is None or len(self.df) == 0:
            logger.warning("CSV no cargado o vacío")
            return None
            
        for intento in range(1, max_intentos + 1):
            if estado_global.esperar_si_pausado():
                return None
                
            resultado = self.df[self.df.iloc[:, 0] == id_buscar]
            
            if len(resultado) > 0:
                logger.info(f"ID {id_buscar} encontrado en intento {intento}")
                return resultado.iloc[0]
            else:
                logger.warning(f"Intento {intento}: ID {id_buscar} no encontrado en el CSV")
                if intento < max_intentos:
                    logger.info(f"Esperando 2 segundos antes de reintentar...")
                    for _ in range(2):
                        if estado_global.esperar_si_pausado():
                            return None
                        time.sleep(1)
        
        logger.error(f"ID {id_buscar} no encontrado después de {max_intentos} intentos")
        return None
    
    def procesar_registro(self):
        try:
            logger.info("Paso 2: Click en (89, 263)")
            pyautogui.click(89, 263)
            
            for _ in range(2):
                if estado_global.esperar_si_pausado():
                    return False, None
                time.sleep(1)
            
            logger.info("Paso 3: Obteniendo ID con AHKManager en (1483, 519)")
            id_obtenido = self.ahk_manager.ejecutar_acciones_ahk(1483, 519)
            
            if not id_obtenido:
                logger.error("No se pudo obtener el ID")
                return False, None
            
            id_obtenido = int(id_obtenido)
            logger.info(f"ID obtenido: {id_obtenido}")

            logger.info(f"Paso 4: Buscando ID {id_obtenido} en CSV (2 intentos máx)")
            registro = self.buscar_por_id(id_obtenido, max_intentos=2)
            
            if registro is None:
                logger.error(f"ID {id_obtenido} no encontrado en CSV después de 2 intentos. Saltando...")
                return True, None
            
            linea_procesada = None
            for idx in range(len(self.df)):
                if self.df.iloc[idx, 0] == id_obtenido:
                    linea_procesada = idx + 1
                    break
            
            if len(registro) >= 2:
                valor_columna_2 = registro.iloc[1]
                if pd.isna(valor_columna_2):
                    valor_columna_2 = ""
                else:
                    valor_columna_2 = str(valor_columna_2)
                
                logger.info(f"Paso 5: Escribiendo valor '{valor_columna_2}' en (1483, 519)")
                
                if valor_columna_2:
                    exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, valor_columna_2)
                    if not exito_escritura:
                        logger.error("Error en la escritura")
                        return False, linea_procesada
                else:
                    logger.info("Columna 2 vacía, no se escribe nada")
                
                for _ in range(2):
                    if estado_global.esperar_si_pausado():
                        return False, linea_procesada
                    time.sleep(1)
            else:
                logger.warning("No hay columna 2 en el registro")
            
            if len(registro) >= 4:
                valor_columna_4 = registro.iloc[3]
                if pd.isna(valor_columna_4):
                    logger.info("Paso 6: Valor columna 4 = ")
                else:
                    logger.info(f"Paso 6: Valor columna 4 = {valor_columna_4}")
                
                if pd.notna(valor_columna_4) and float(valor_columna_4) > 0:
                    veces_down = int(float(valor_columna_4))
                    logger.info(f"Paso 7: Ejecutando {veces_down} veces DOWN en (1507, 636)")
                    
                    exito_down = self.ahk_click_down.ejecutar_click_down(1507, 636, veces_down)
                    if not exito_down:
                        logger.error("Error en click + down")
                        return False, linea_procesada
                    
                    for _ in range(2):
                        if estado_global.esperar_si_pausado():
                            return False, linea_procesada
                        time.sleep(1)
                else:
                    logger.info("Paso 7: Saltado (columna 4 <= 0 o vacía)")
            else:
                logger.warning("No hay columna 4 en el registro")
            
            logger.info("Paso 8: Click en (1290, 349)")
            pyautogui.click(1290, 349)
            
            for _ in range(2):
                if estado_global.esperar_si_pausado():
                    return False, linea_procesada
                time.sleep(1)
            
            logger.info("Procesamiento completado exitosamente")
            return True, linea_procesada
            
        except Exception as e:
            logger.error(f"Error en procesar_registro: {e}")
            return False, None
    
    def procesar_todo(self, pausa_entre_registros=2):
        if not self.cargar_csv():
            return False, None
            
        if self.df is None or len(self.df) == 0:
            logger.info("CSV vacío. No hay registros para procesar.")
            return True, None
            
        if not self.iniciar_ahk():
            return False, None
        
        try:
            logger.info("Iniciando procesamiento de registro...")
            exito, linea_procesada = self.procesar_registro()
            
            if exito:
                logger.info(f"Procesamiento completado. Línea procesada: {linea_procesada}")
            else:
                logger.error("Procesamiento falló")
                
            return exito, linea_procesada
            
        finally:
            self.detener_ahk()

class NSEAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica
        self.csv_file = CSV_FILE
        self.reference_image = "img/VentanaAsignar.png"
        self.is_running = False
        
        self.ahk_writer = AHKWriter()
        
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.8
        
        self.coords_select = {
            7: [33, 92], 8: [33, 131], 9: [33, 159], 10: [33, 197],
            11: [33, 231], 12: [398, 92], 13: [398, 131], 14: [398, 159],
            15: [33, 301], 16: [33, 333], 17: [33, 367]
        }
        
        self.coords_type = {
            7: [163, 92], 8: [163, 131], 9: [163, 159], 10: [163, 197],
            11: [163, 231], 12: [528, 92], 13: [528, 131], 14: [528, 159],
            15: [163, 301], 16: [163, 333], 17: [163, 367]
        }
        
        self.coords_asignar = [446, 281]
        self.coords_cerrar = [396, 352]

    def click(self, x, y, duration=0.2):
        pyautogui.click(x, y, duration=duration)
        for _ in range(1):
            if estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def write_with_ahk(self, x, y, text):
        if pd.isna(text) or text is None or str(text).strip() == "":
            return True
            
        text_str = str(text).strip()
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, text_str)
        if not success:
            print(f"❌ Error al escribir con AHK en ({x}, {y}): {text_str}")
        for _ in range(1):
            if estado_global.esperar_si_pausado():
                return False
            time.sleep(1.5)
        return success

    def sleep(self, seconds):
        for _ in range(int(seconds * 1.5)):
            if estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def detect_image_with_cv2(self, image_path, confidence=0.6):
        try:
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            template = cv2.imread(image_path)
            if template is None:
                print(f"Error: No se pudo cargar la imagen {image_path}")
                return False, None
            
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            if max_val < confidence:
                print(f"Imagen no encontrada. Mejor coincidencia: {max_val:.2f}")
                return False, None
            
            print(f"Imagen encontrada con confianza: {max_val:.2f}")
            return True, max_loc
        except Exception as e:
            print(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.6):
        print(f"🔍 Buscando imagen: {image_path}")
        
        for attempt in range(1, max_attempts + 1):
            if estado_global.esperar_si_pausado():
                return False, None
                
            found, location = self.detect_image_with_cv2(image_path, confidence)
            
            if found:
                print(f"✅ Imagen detectada en el intento {attempt} en coordenadas: {location}")
                return True, location
            
            print(f"⏳ Intento {attempt}/{max_attempts} - Imagen no encontrada")
            
            if attempt < max_attempts:
                if attempt % 10 == 0:
                    print("⏰ Espera prolongada de 15 segundos...")
                    for _ in range(15):
                        if estado_global.esperar_si_pausado():
                            return False, None
                        time.sleep(1)
                else:
                    for _ in range(3):
                        if estado_global.esperar_si_pausado():
                            return False, None
                        time.sleep(1)
        
        print("❌ Imagen no encontrada después de 30 intentos. Terminando proceso.")
        return False, None

    def should_skip_process(self, row):
        if pd.notna(row.iloc[5]):
            col_value = str(row.iloc[5]).strip()
            if col_value.lower() == 'nan':
                return False
            if col_value and col_value != "" and col_value != "  ":
                return True
        return False

    def execute_nse_script(self):
        if not self.ahk_writer.start_ahk():
            print("❌ No se pudo iniciar AHKWriter")
            return False
            
        try:
            try:
                df = pd.read_csv(self.csv_file)
                if len(df) == 0:
                    print("⚠️ CSV vacío. Saltando Programa 2...")
                    return True
            except:
                pass
            
            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False
            
            row = df.iloc[self.linea_especifica - 1]
            print(f"🔄 Procesando línea {self.linea_especifica}/{total_lines}")
            
            if self.should_skip_process(row):
                print(f"⏭️  Saltando línea {self.linea_especifica} - Columna 6 tiene valor: {row.iloc[5]}")
                return True
            
            if str(row.iloc[4]).strip().upper() != "V":
                print(f"⚠️  Saltando línea {self.linea_especifica} - No es tipo V: {row.iloc[4]}")
                return True
            
            self.click(169, 189)
            self.sleep(3)
            self.click(1491, 386)
            self.sleep(3)
            
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=30)
            
            if not image_found:
                print("❌ No se puede continuar sin detectar la imagen de referencia.")
                return False
            
            print("🎯 Imagen detectada, procediendo con tipo V")
            self.handle_type_v(row, base_location)
            
            print(f"✅ Línea {self.linea_especifica} completada (hasta CERRAR)")
            return True
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            return False
        finally:
            self.ahk_writer.stop_ahk()

    def handle_type_v(self, row, base_location):
        base_x, base_y = base_location
        
        for col_index in range(7, 18):
            if estado_global.esperar_si_pausado():
                return
                
            if pd.notna(row.iloc[col_index-1]) and row.iloc[col_index-1] > 0:
                x_cs_rel, y_cs_rel = self.coords_select[col_index]
                x_ct_rel, y_ct_rel = self.coords_type[col_index]
                
                x_cs_abs = base_x + x_cs_rel
                y_cs_abs = base_y + y_cs_rel
                x_ct_abs = base_x + x_ct_rel
                y_ct_abs = base_y + y_ct_rel
                
                self.click(x_cs_abs, y_cs_abs)
                self.sleep(3)
                
                texto = str(int(row.iloc[col_index-1]))
                self.write_with_ahk(x_ct_abs, y_ct_abs, texto)
                self.sleep(3)
        
        x_asignar_rel, y_asignar_rel = self.coords_asignar
        x_asignar_abs = base_x + x_asignar_rel
        y_asignar_abs = base_y + y_asignar_rel
        self.click(x_asignar_abs, y_asignar_abs)
        self.sleep(3)
        
        x_cerrar_rel, y_cerrar_rel = self.coords_cerrar
        x_cerrar_abs = base_x + x_cerrar_rel
        y_cerrar_abs = base_y + y_cerrar_rel
        self.click(x_cerrar_abs, y_cerrar_abs)
        self.sleep(3)

class NSEServicesAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica
        self.csv_file = CSV_FILE
        self.current_line = 0
        self.is_running = False
        self.reference_point = None
        
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        self.ahk_enter = EnterAHKManager()
        
        self.coords = {
            'menu_principal': (81, 81),
            'campo_cantidad': (108, 350),
            'boton_guardar': (63, 390),
            'boton_error': (704, 384),
            'cierre': (863, 16),
            'inicio_servicios': (1563, 385),
            'casilla_servicio': (121, 236),
            'casilla_tipo': (121, 261),
            'casilla_empresa': (121, 290),
            'casilla_producto': (121, 322),
        }

    def buscar_imagen(self, imagen_path, timeout=30, confidence=0.8):
        logging.info(f"🔍 Buscando imagen: {imagen_path}")
        
        try:
            template = cv2.imread(imagen_path)
            if template is None:
                logging.error(f"❌ No se pudo cargar la imagen: {imagen_path}")
                return None
            
            template_height, template_width = template.shape[:2]
            
            for intento in range(timeout):
                if estado_global.esperar_si_pausado():
                    return None
                    
                screenshot = ImageGrab.grab()
                screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confidence:
                    x, y = max_loc
                    logging.info(f"✅ Imagen encontrada en intento {intento + 1} - Coordenadas: ({x}, {y}) - Confianza: {max_val:.2f}")
                    return (x, y)
                
                logging.info(f"⏳ Intento {intento + 1}/{timeout} - Confianza máxima: {max_val:.2f}")
                
                for _ in range(2):
                    if estado_global.esperar_si_pausado():
                        return None
                    time.sleep(1)
            
            logging.error(f"❌ No se encontró la imagen después de {timeout} intentos")
            return None
            
        except Exception as e:
            logging.error(f"❌ Error en búsqueda de imagen: {e}")
            return None

    def actualizar_coordenadas_relativas(self, referencia):
        if referencia is None:
            logging.error("❌ No se puede actualizar coordenadas: referencia es None")
            return False
        
        ref_x, ref_y = referencia
        
        self.coords_relativas = {
            'menu_principal': (ref_x + 81, ref_y + 81),
            'campo_cantidad': (ref_x + 108, ref_y + 350),
            'boton_guardar': (ref_x + 63, ref_y + 390),
            'cierre': (ref_x + 863, ref_y + 16),
            'casilla_servicio': (ref_x + 121, ref_y + 236),
            'casilla_tipo': (ref_x + 121, ref_y + 261),
            'casilla_empresa': (ref_x + 121, ref_y + 290),
            'casilla_producto': (ref_x + 121, ref_y + 322),
        }
        
        self.coords_relativas['boton_error'] = self.coords['boton_error']
        self.coords_relativas['inicio_servicios'] = self.coords['inicio_servicios']
        
        self.reference_point = referencia
        logging.info("✅ Coordenadas actualizadas a relativas")
        return True

    def iniciar_ahk(self):
        try:
            if not self.ahk_writer.start_ahk():
                logging.error("No se pudo iniciar AHK Writer")
                return False
            if not self.ahk_click_down.start_ahk():
                logging.error("No se pudo iniciar AHK Click Down")
                return False
            if not self.ahk_enter.start_ahk():
                logging.error("No se pudo iniciar AHK Enter")
                return False
            logging.info("✅ Todos los servicios AHK iniciados correctamente")
            return True
        except Exception as e:
            logging.error(f"Error iniciando servicios AHK: {e}")
            return False

    def detener_ahk(self):
        try:
            self.ahk_writer.stop_ahk()
            self.ahk_click_down.stop_ahk()
            self.ahk_enter.stop_ahk()
            logging.info("✅ Todos los servicios AHK detenidos correctamente")
        except Exception as e:
            logging.error(f"Error deteniendo servicios AHK: {e}")

    def click(self, x, y, duration=0.2):
        pyautogui.click(x, y, duration=duration)
        for _ in range(1):
            if estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def write(self, text):
        try:
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                campo_coords = self.coords_relativas['campo_cantidad']
                logger.info(f"Usando coordenadas relativas para escribir: {campo_coords}")
            else:
                campo_coords = self.coords['campo_cantidad']
                logger.info(f"Usando coordenadas absolutas para escribir: {campo_coords}")

            success = self.ahk_writer.ejecutar_escritura_ahk(
                campo_coords[0],
                campo_coords[1],
                str(text)
            )
            if success:
                logging.info(f"✅ Texto escrito exitosamente: {text}")
                for _ in range(2):
                    if estado_global.esperar_si_pausado():
                        return False
                    time.sleep(1)
            else:
                logging.error(f"❌ Error al escribir texto: {text}")
            return success
        
        except Exception as e:
            logging.error(f"Error escribiendo texto '{text}': {e}")
            return False

    def press_down(self, x, y, times=1):
        try:
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                click_coords = (x, y)
            else:
                click_coords = (x, y)
                
            return self.ahk_click_down.ejecutar_click_down(click_coords[0], click_coords[1], times)
        except Exception as e:
            logging.error(f"Error presionando DOWN {times} veces: {e}")
            return False

    def press_enter(self):
        try:                
            return self.ahk_enter.presionar_enter(1)
        except Exception as e:
            logging.error(f"Error presionando enter: {e}")
            return False

    def sleep(self, seconds):
        for _ in range(int(seconds * 1.5)):
            if estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def handle_error_click(self):
        for _ in range(5):
            if estado_global.esperar_si_pausado():
                return
                
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                self.click(*self.coords_relativas['boton_error'])
            else:
                self.click(*self.coords['boton_error'])
            self.sleep(3)

    def procesar_linea_especifica(self):
        try:
            df = pd.read_csv(self.csv_file, encoding='utf-8')
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False
            
            linea_idx = self.linea_especifica - 1
            self.current_line = self.linea_especifica
            
            print(f"🎯 PROCESANDO LÍNEA ESPECÍFICA: {self.current_line}/{total_lines}")
            logger.info(f"PROCESANDO LÍNEA ESPECÍFICA: {self.current_line}/{total_lines}")
            
            row = df.iloc[linea_idx]
            
            if pd.notna(row.iloc[17]) and row.iloc[17] > 0:
                print(f"✅ Línea {self.current_line} tiene servicios para procesar")
                
                self.click(*self.coords['inicio_servicios'])
                self.sleep(3)
                
                print("🔍 Buscando ventana de servicios...")
                referencia = self.buscar_imagen("img/ventanaAdministracion4.PNG", timeout=30)
                
                if referencia is None:
                    print("❌ ERROR: No se pudo encontrar la ventana de servicios")
                    return False
                
                if not self.actualizar_coordenadas_relativas(referencia):
                    print("❌ ERROR: No se pudieron actualizar las coordenadas relativas")
                    return False
                
                self.click(*self.coords_relativas['menu_principal'])
                self.sleep(3)
                    
                servicios_procesados = 0
                logger.info(f"Procesando servicios para línea {self.current_line}")
                
                servicios = [
                    (18, self.handle_voz_cobre, "VOZ COBRE TELMEX"),
                    (19, self.handle_datos_sdom, "DATOS S/DOM"),
                    (20, self.handle_datos_cobre_telmex, "DATOS COBRE TELMEX"),
                    (21, self.handle_datos_fibra_telmex, "DATOS FIBRA TELMEX"),
                    (22, self.handle_tv_cable_otros, "TV CABLE OTROS"),
                    (23, self.handle_dish, "DISH"),
                    (24, self.handle_tvs, "TVS"),
                    (25, self.handle_sky, "SKY"),
                    (26, self.handle_vetv, "VETV"),
                ]
                
                for col_idx, metodo, nombre in servicios:
                    if estado_global.esperar_si_pausado():
                        return False
                        
                    if pd.notna(row.iloc[col_idx]) and row.iloc[col_idx] > 0:
                        print(f"  └─ Procesando {nombre}: {row.iloc[col_idx]}")
                        logger.info(f"  └─ Procesando {nombre}: {row.iloc[col_idx]}")
                        metodo(row.iloc[col_idx])
                        servicios_procesados += 1
                
                self.click(*self.coords_relativas['cierre'])
                self.sleep(5)
                
                print(f"✅ Línea {self.current_line} completada: {servicios_procesados} servicios procesados")
                return True
            else:
                print(f"⏭️  Línea {self.current_line} no tiene servicios para procesar")
                return True
            
        except Exception as e:
            print(f"❌ Error procesando línea {self.current_line}: {e}")
            logging.error(f"Error en procesar_linea_especifica: {e}")
            return False

    def handle_voz_cobre(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.write(cantidad)
        logger.info(f"Escribio cantidad de VOZ COBRE TELMEX: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

    def handle_datos_sdom(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.sleep(3)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de DATOS S/DOM: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

    def handle_datos_cobre_telmex(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_producto'], 1)
        self.press_enter()
        self.sleep(3)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de DATOS COBRE TELMEX: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

    def handle_datos_fibra_telmex(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 1)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 1)
        self.press_enter()
        self.sleep(3)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de DATOS FIBRA TELMEX: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

    def handle_tv_cable_otros(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 4)
        self.press_enter()
        self.sleep(3)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de TV CABLE OTROS: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

    def handle_dish(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 1)
        self.press_enter()
        self.sleep(3)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de DISH: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

    def handle_tvs(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 2)
        self.press_enter()
        self.sleep(3)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de TVS: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

    def handle_sky(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 3)
        self.press_enter()
        self.sleep(3)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de SKY: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

    def handle_vetv(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(3)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 5)
        self.press_enter()
        self.sleep(3)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de VETV: {cantidad}")
        self.sleep(3)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(3)
        self.handle_error_click()

class GEAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica
        self.csv_file = CSV_FILE
        self.reference_image = "img/textoAdicional.PNG"
        self.ventana_archivo_img = "img/cargarArchivo.png"
        self.ventana_error_img = "img/ventanaError.png"
        self.is_running = False
        
        self.ahk_writer = AHKWriter()
        self.ahk_manager = AHKManager()
        self.enter = EnterAHKManager()
        self.ahk_click_down = AHKClickDown()
        
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.8

        self.nombre=""
        
        self.coords = {
            'agregar_ruta': (327, 381),
            'archivo': (1396, 608),
            'abrir': (1406, 634),
            'seleccionar_mapa': (168, 188),
            'anotar': (1366, 384),
            'agregar_texto_adicional': (1449, 452),
            'limpiar_trazo': (360, 980),
            'lote_again': (70, 266),
            'cerrar_ventana_archivo': (1530, 555)
        }
        
        self.coords_texto_relativas = {
            'campo_texto': (230, 66),
            'agregar_texto': (64, 100),
            'cerrar_ventana_texto': (139, 98)
        }

    def encontrar_ventana_archivo(self):
        intentos = 1
        confianza_minima = 0.6
        tiempo_espera_base = 1.5
        tiempo_espera_largo = 12
        
        template = cv2.imread(self.ventana_archivo_img)
        if template is None:
            logger.error(f"No se pudo cargar la imagen '{self.ventana_archivo_img}'")
            return None
        
        while self.is_running: 
            try:
                if estado_global.esperar_si_pausado():
                    return None
                    
                screenshot = pyautogui.screenshot()
                pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confianza_minima:
                    logger.info(f"Ventana encontrada con confianza: {max_val:.2f}")
                    return max_loc
                else:
                    if intentos % 10 == 0 and intentos > 0:
                        logger.info(f"Intento {intentos}: Mejor coincidencia: {max_val:.2f}")
                        logger.info("Esperando 12 segundos...")
                        for _ in range(12):
                            if estado_global.esperar_si_pausado():
                                return None
                            time.sleep(1)
                    else:
                        for _ in range(int(tiempo_espera_base)):
                            if estado_global.esperar_si_pausado():
                                return None
                            time.sleep(1)
                    intentos += 1
                    
            except Exception as e:
                logger.error(f"Error durante la búsqueda: {e}")
                for _ in range(int(tiempo_espera_base)):
                    if estado_global.esperar_si_pausado():
                        return None
                    time.sleep(1)
                intentos += 1

        return None

    def detectar_ventana_error(self):
        try:
            template = cv2.imread(self.ventana_error_img) 
            if template is None:
                logger.error(f"No se pudo cargar la imagen '{self.ventana_error_img}'")
                return False
            
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            confianza_minima = 0.6
            
            if max_val >= confianza_minima:
                logger.info(f"Ventana de error detectada con confianza: {max_val:.2f}")
                
                if not self.enter.start_ahk():
                    logger.error("No se pudo iniciar AutoHotkey")
                    return False
                    
                if self.enter.presionar_enter(1):
                    for _ in range(3):
                        if estado_global.esperar_si_pausado():
                            self.enter.stop_ahk()
                            return True
                        time.sleep(1)
                else:
                    logger.error("Error enviando comando a AHK")
                    return False
                
                logger.info("Ventana de error detectada y cerrada")
                self.enter.stop_ahk()
                return True
            else:
                return False
                
        except Exception as e:
            logger.error(f"Error al detectar ventana de error: {e}")
            return False

    def handle_archivo_special_behavior(self, nombre_archivo):
        coordenadas_ventana = self.encontrar_ventana_archivo()

        if coordenadas_ventana:
            x_ventana, y_ventana = coordenadas_ventana
            logger.info(f"Coordenadas ventana: x={x_ventana}, y={y_ventana}")
            
            x_campo = x_ventana + 294
            y_campo = y_ventana + 500
            logger.info(f"Coordenadas campo texto: x={x_campo}, y={y_campo}")
            
            if not self.ahk_manager.start_ahk():
                logger.error("No se pudo iniciar AutoHotkey")
                return False
            
            if self.ahk_manager.ejecutar_acciones_ahk(x_campo, y_campo, nombre_archivo):
                for _ in range(2):
                    if estado_global.esperar_si_pausado():
                        self.ahk_manager.stop_ahk()
                        return False
                    time.sleep(1)
            else:
                logger.error("Error enviando comando a AHK")
                return False
            
            self.ahk_manager.stop_ahk()
            return True
        else:
            logger.error("No se pudo encontrar la ventana de archivo.")
            return False

    def escribir_texto_adicional_ahk(self, x, y, texto):
        if pd.isna(texto) or texto is None or str(texto).strip() == '' or str(texto).strip().lower() == 'nan':
            print("⚠️  Texto adicional vacío, saltando escritura")
            return True
            
        texto_str = str(texto).strip()
        print(f"📝 Intentando escribir texto: '{texto_str}' en coordenadas ({x}, {y})")
        
        if x <= 0 or y <= 0:
            print(f"❌ Coordenadas inválidas: ({x}, {y})")
            return False
        
        if not self.ahk_writer.start_ahk():
            logger.error("No se pudo iniciar AHK Writer")
            print("❌ Falló al iniciar AHK Writer")
            return False
        
        print("🔄 AHK Writer iniciado, enviando comando...")
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, texto_str)
        self.ahk_writer.stop_ahk()
        
        if success:
            print(f"✅ Texto escrito exitosamente: '{texto_str}'")
        else:
            print(f"❌ Error al escribir texto: '{texto_str}'")
            print("🔄 Intentando método alternativo con pyautogui...")
            try:
                self.click(x, y)
                for _ in range(1):
                    if estado_global.esperar_si_pausado():
                        return False
                    time.sleep(1.5)
                pyautogui.hotkey('ctrl', 'a')
                for _ in range(1):
                    if estado_global.esperar_si_pausado():
                        return False
                    time.sleep(1)
                pyautogui.press('delete')
                for _ in range(1):
                    if estado_global.esperar_si_pausado():
                        return False
                    time.sleep(1)
                pyautogui.write(texto_str, interval=0.1)
                print(f"✅ Texto escrito con pyautogui: '{texto_str}'")
                success = True
            except Exception as e:
                print(f"❌ También falló pyautogui: {e}")
                
        return success

    def presionar_flecha_abajo_ahk(self, x,y,veces=1):
        if not self.ahk_click_down.start_ahk():
            logger.error("No se pudo iniciar AutoHotkey para flecha abajo")
            return False
        
        try:
            self.ahk_click_down.ejecutar_click_down(x, y, veces)
            return True
        except Exception as e:
            logger.error(f"Error presionando flecha abajo: {e}")
            return False
        finally:
            self.ahk_click_down.stop_ahk()

    def presionar_enter_ahk(self, veces=1):
        if not self.enter.start_ahk():
            logger.error("No se pudo iniciar AutoHotkey para Enter")
            return False
        
        success = self.enter.presionar_enter(veces)
        self.enter.stop_ahk()
        return success

    def click(self, x, y, duration=0.2):
        pyautogui.click(x, y, duration=duration)
        for _ in range(1):
            if estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def sleep(self, seconds):
        for _ in range(int(seconds * 1.5)):
            if estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def detect_image_with_cv2(self, image_path, confidence=0.7):
        try:
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            template = cv2.imread(image_path)
            if template is None:
                print(f"Error: No se pudo cargar la imagen {image_path}")
                return False, None
            
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            if max_val < confidence:
                print(f"Imagen no encontrada. Mejor coincidencia: {max_val:.2f}")
                return False, None
            
            print(f"✅ Imagen encontrada con confianza: {max_val:.2f}")
            return True, max_loc
        except Exception as e:
            print(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.7):
        print(f"🔍 Buscando imagen: {image_path}")
        
        for attempt in range(1, max_attempts + 1):
            if estado_global.esperar_si_pausado():
                return False, None
                
            found, location = self.detect_image_with_cv2(image_path, confidence)
            
            if found:
                print(f"✅ Imagen detectada en el intento {attempt} en coordenadas: {location}")
                return True, location
            
            print(f"⏳ Intento {attempt}/{max_attempts} - Imagen no encontrada")
            
            if attempt < max_attempts:
                if attempt % 10 == 0:
                    print("⏰ Espera prolongada de 12 segundos...")
                    for _ in range(12):
                        if estado_global.esperar_si_pausado():
                            return False, None
                        time.sleep(1)
                else:
                    for _ in range(3):
                        if estado_global.esperar_si_pausado():
                            return False, None
                        time.sleep(1)
        
        print("❌ Imagen no encontrada después de 30 intentos. Terminando proceso.")
        return False, None

    def verificar_valores_csv(self, df, row_index):
        try:
            if row_index >= len(df):
                print(f"❌ Fila {row_index} no existe en el CSV")
                return False
            
            row = df.iloc[row_index]
            if len(row) <= 27 or pd.isna(row.iloc[27]) or row.iloc[27] != 1:
                print(f"⚠️  Columna 27 vacía, no es 1 o no existe en fila {row_index}, saltando...")
                if not self.presionar_flecha_abajo_ahk(70, 266,1):
                    print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    print("✅ Flecha abajo presionada con AHK")
                self.sleep(3)
                self.presionar_enter_ahk(1)
                self.sleep(1)
                return False
            
            if len(row) <= 28 or pd.isna(row.iloc[28]):
                print(f"⚠️  Columna 28 vacía o no existe en fila {row_index}, saltando...")
                if not self.presionar_flecha_abajo_ahk(70, 266,1):
                    print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    print("✅ Flecha abajo presionada con AHK")
                self.sleep(2)
                self.presionar_enter_ahk(1)
                self.sleep(1)
                return False
                
            if len(row) <= 29:
                print(f"⚠️  Columna 29 no existe en fila {row_index}, saltando...")
                if not self.presionar_flecha_abajo_ahk(70, 266,1):
                    print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    print("✅ Flecha abajo presionada con AHK")
                self.sleep(2)
                self.presionar_enter_ahk(1)
                self.sleep(1)
                return False
                
            return True
            
        except Exception as e:
            print(f"❌ Error verificando valores CSV: {e}")
            return False

    def perform_actions(self):
        if not self.ahk_writer.start_ahk():
            print("❌ No se pudo iniciar AHKWriter")
            return False
            
        try:
            if not os.path.exists(self.csv_file):
                print(f"❌ El archivo CSV no existe: {self.csv_file}")
                return False

            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            if total_lines < 1:
                print("❌ No hay suficientes datos en el archivo CSV")
                return False

            print(f"📊 Total de líneas en CSV: {total_lines}")

            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False

            row_index = self.linea_especifica - 1
                
            if not self.verificar_valores_csv(df, row_index):
                print(f"⚠️  Valores inválidos en fila {row_index}. Línea {self.linea_especifica} saltada.")
                return True
                    
            print(f"🔄 Procesando línea {self.linea_especifica}/{total_lines}")
            success = self.process_single_iteration(df, self.linea_especifica, total_lines)
                
            if not success:
                print(f"⚠️  Línea {self.linea_especifica} falló")
                return False
                
            return True
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            return False
        finally:
            self.ahk_writer.stop_ahk()
            self.ahk_manager.stop_ahk()
            self.enter.stop_ahk()
            self.ahk_click_down.stop_ahk()

    def process_single_iteration(self, df, linea_especifica, total_lines):
        row_index = linea_especifica - 1
        row = df.iloc[row_index]
        
        try:
            num_txt_type = str(int(row.iloc[28])) if not pd.isna(row.iloc[28]) else None
            texto_adicional = str(row.iloc[29]) if not pd.isna(row.iloc[29]) else ""
        except (ValueError, IndexError) as e:
            print(f"❌ Error obteniendo valores del CSV: {e}")
            return False

        if not num_txt_type:
            print(f"⚠️  num_txt_type vacío en línea {linea_especifica}, saltando...")
            return False
        self.nombre="NN "+num_txt_type+".kml"

        print(f"📁 Archivo a cargar: {self.nombre}")
        print(f"📝 Texto adicional: '{texto_adicional}'")

        try:
            self.click(*self.coords['agregar_ruta'])
            self.sleep(3)
            self.click(*self.coords['archivo'])
            self.sleep(3)
            self.click(*self.coords['abrir'])
            self.sleep(3)
            
            nombre_archivo = self.nombre
            success = self.handle_archivo_special_behavior(nombre_archivo)
            
            if not success:
                print("❌ No se pudo cargar el archivo. Regresando a agregar_ruta...")
                self.click(*self.coords['agregar_ruta'])
                self.sleep(3)
                return False
            
            if not self.presionar_enter_ahk(1):
                print("⚠️  No se pudo presionar Enter con AHK, usando pyautogui")
                pyautogui.press('enter')
            
            self.sleep(4)

            self.click(*self.coords['agregar_ruta'])
            self.sleep(3)

            self.click(1406, 675)
            self.sleep(3)

            self.click(70, 266)
            self.sleep(3)

            self.click(*self.coords['cerrar_ventana_archivo'])
            self.sleep(4)

            self.click(*self.coords['seleccionar_mapa'])
            self.sleep(3)
            
            self.click(*self.coords['anotar'])
            self.sleep(3)
            
            self.click(*self.coords['agregar_texto_adicional'])
            self.sleep(3)
            
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=10)
            
            if image_found:
                x_campo = base_location[0] + self.coords_texto_relativas['campo_texto'][0]
                y_campo = base_location[1] + self.coords_texto_relativas['campo_texto'][1]
                x_agregar = base_location[0] + self.coords_texto_relativas['agregar_texto'][0]
                y_agregar = base_location[1] + self.coords_texto_relativas['agregar_texto'][1]
                x_cerrar = base_location[0] + self.coords_texto_relativas['cerrar_ventana_texto'][0]
                y_cerrar = base_location[1] + self.coords_texto_relativas['cerrar_ventana_texto'][1]
                
                if texto_adicional and texto_adicional.strip():
                    writing_success = self.escribir_texto_adicional_ahk(x_campo, y_campo, texto_adicional)
                    if not writing_success:
                        print("⚠️  Falló la escritura con AHK, intentando con pyautogui...")
                        pyautogui.write(texto_adicional, interval=0.1)
                else:
                    print("ℹ️  Texto adicional vacío, no se escribe nada")
                
                self.sleep(3)

                self.click(x_agregar, y_agregar)
                self.sleep(4)
            
                self.click(x_cerrar, y_cerrar)
                self.sleep(3)
            else:
                print("❌ No se pudo detectar la imagen del campo de texto")
                return False
                        
            self.click(*self.coords['limpiar_trazo'])
            self.sleep(2)
            
            self.click(*self.coords['lote_again'])
            self.sleep(3)
            
            if not self.presionar_flecha_abajo_ahk(*self.coords['lote_again'],1):
                print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                pyautogui.press('down')
            else:
                print("✅ Flecha abajo presionada con AHK")
            
            self.sleep(3)
            
            if self.detectar_ventana_error():
                print("✅ Ventana de error detectada y cerrada")
        
            self.presionar_enter_ahk(1)
            self.sleep(1)
            print(f"✅ Línea {linea_especifica} completada exitosamente")

            return True
            
        except Exception as e:
            print(f"❌ Error en línea {linea_especifica}: {e}")
            self.detectar_ventana_error()
            return False

def main():
    root = tk.Tk()
    app = InterfazAutomation(root)
    root.mainloop()

if __name__ == "__main__":
    main()