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
from tkinter import ttk, filedialog, messagebox, simpledialog
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
        logging.StreamHandler(sys.stdout)  # Usar stdout con encoding UTF-8
    ]
)
logger = logging.getLogger(__name__)

# Variables globales
LINEA_A_PROCESAR = None
CSV_FILE = ""
KML_FILENAME = "NN"  # Nombre por defecto para archivos KML
EJECUTANDO = False
PAUSADO = False
LINEA_ACTUAL = 0
LINEA_MAXIMA = 0
HILO_EJECUCION = None

class InterfazAutomation:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Automatización NSE/GE")
        self.root.geometry("800x600")
        
        # Variables
        self.csv_file = tk.StringVar()
        self.linea_maxima = tk.IntVar(value=1)
        self.estado = tk.StringVar(value="Listo")
        self.linea_actual = tk.StringVar(value="0")
        self.lineas_restantes = tk.StringVar(value="0")
        
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
        csv_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        csv_frame.columnconfigure(1, weight=1)
        
        ttk.Label(csv_frame, text="Archivo CSV:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Entry(csv_frame, textvariable=self.csv_file, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(csv_frame, text="Seleccionar", command=self.seleccionar_csv).grid(row=0, column=2, padx=5)
        
        # Sección ejecución
        exec_frame = ttk.LabelFrame(main_frame, text="Control de Ejecución", padding="5")
        exec_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(exec_frame, text="Línea máxima a ejecutar:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Spinbox(exec_frame, from_=1, to=10000, textvariable=self.linea_maxima, width=10).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(exec_frame, text="Línea actual:").grid(row=1, column=0, sticky=tk.W, padx=5)
        ttk.Label(exec_frame, textvariable=self.linea_actual).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        ttk.Label(exec_frame, text="Líneas restantes:").grid(row=2, column=0, sticky=tk.W, padx=5)
        ttk.Label(exec_frame, textvariable=self.lineas_restantes).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Botones de control
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        self.btn_iniciar = ttk.Button(control_frame, text="Iniciar Proceso", command=self.iniciar_proceso)
        self.btn_iniciar.grid(row=0, column=0, padx=5)
        
        self.btn_pausar = ttk.Button(control_frame, text="Pausar (F2)", command=self.pausar_proceso, state=tk.DISABLED)
        self.btn_pausar.grid(row=0, column=1, padx=5)
        
        self.btn_reanudar = ttk.Button(control_frame, text="Reanudar (F3)", command=self.reanudar_proceso, state=tk.DISABLED)
        self.btn_reanudar.grid(row=0, column=2, padx=5)
        
        self.btn_detener = ttk.Button(control_frame, text="Detener (F4)", command=self.detener_proceso, state=tk.DISABLED)
        self.btn_detener.grid(row=0, column=3, padx=5)
        
        # Botones adicionales
        extra_frame = ttk.Frame(main_frame)
        extra_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        ttk.Button(extra_frame, text="Escribir PRUEBA A", command=self.escribir_prueba_a).grid(row=0, column=0, padx=5)
        ttk.Button(extra_frame, text="Configurar KML", command=self.configurar_kml).grid(row=0, column=1, padx=5)
        
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
        
    def setup_bindings(self):
        # Configurar teclas globales
        keyboard.add_hotkey('esc', self.mostrar_estado_actual)
        keyboard.add_hotkey('f2', self.pausar_proceso)
        keyboard.add_hotkey('f3', self.reanudar_proceso)
        keyboard.add_hotkey('f4', self.detener_proceso)
        
    def log(self, mensaje):
        """Agregar mensaje al log"""
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {mensaje}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def seleccionar_csv(self):
        """Seleccionar archivo CSV"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if archivo:
            self.csv_file.set(archivo)
            global CSV_FILE
            CSV_FILE = archivo
            self.log(f"CSV seleccionado: {archivo}")
            
            # Calcular número máximo de líneas
            try:
                df = pd.read_csv(archivo)
                self.linea_maxima.set(len(df))
                self.actualizar_estado_lineas()
            except Exception as e:
                self.log(f"Error al leer CSV: {e}")
    
    def escribir_prueba_a(self):
        """Escribir PRUEBA A desde la última columna de la primera fila"""
        if not self.csv_file.get():
            messagebox.showerror("Error", "Primero seleccione un archivo CSV")
            return
            
        try:
            df = pd.read_csv(self.csv_file.get())
            if len(df) == 0:
                messagebox.showerror("Error", "El CSV está vacío")
                return
                
            # Obtener última columna de la primera fila
            ultima_columna = df.iloc[0, -1]
            texto_a_escribir = str(ultima_columna)
            
            self.log(f"Escribiendo: {texto_a_escribir}")
            
            # Usar AHKWriter para escribir
            ahk_writer = AHKWriter()
            if ahk_writer.start_ahk():
                # Coordenadas donde se debe escribir (ajustar según necesidad)
                exito = ahk_writer.ejecutar_escritura_ahk(100, 100, texto_a_escribir)
                ahk_writer.stop_ahk()
                
                if exito:
                    self.log("✅ Texto escrito exitosamente")
                else:
                    self.log("❌ Error al escribir texto")
            else:
                self.log("❌ No se pudo iniciar AHKWriter")
                
        except Exception as e:
            self.log(f"❌ Error al escribir PRUEBA A: {e}")
    
    def configurar_kml(self):
        """Configurar el nombre de los archivos KML"""
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
        """Actualizar display de líneas actuales y restantes"""
        global LINEA_ACTUAL, LINEA_MAXIMA
        self.linea_actual.set(str(LINEA_ACTUAL))
        lineas_rest = max(0, LINEA_MAXIMA - LINEA_ACTUAL)
        self.lineas_restantes.set(str(lineas_rest))
    
    def mostrar_estado_actual(self):
        """Mostrar estado actual al presionar ESC"""
        global LINEA_ACTUAL, LINEA_MAXIMA
        if EJECUTANDO:
            lineas_restantes = LINEA_MAXIMA - LINEA_ACTUAL
            mensaje = f"Línea actual: {LINEA_ACTUAL}\nLíneas restantes: {lineas_restantes}"
            messagebox.showinfo("Estado Actual", mensaje)
    
    def iniciar_proceso(self):
        """Iniciar el proceso completo"""
        global EJECUTANDO, PAUSADO, LINEA_ACTUAL, LINEA_MAXIMA, CSV_FILE
        
        if not self.csv_file.get():
            messagebox.showerror("Error", "Seleccione un archivo CSV primero")
            return
            
        CSV_FILE = self.csv_file.get()
        LINEA_MAXIMA = self.linea_maxima.get()
        LINEA_ACTUAL = 1
        
        if LINEA_ACTUAL > LINEA_MAXIMA:
            messagebox.showerror("Error", "La línea actual no puede ser mayor que la línea máxima")
            return
            
        EJECUTANDO = True
        PAUSADO = False
        
        self.actualizar_estado_botones()
        self.estado.set("Ejecutando...")
        self.estado_label.configure(foreground="green")
        
        # Iniciar en hilo separado
        self.hilo_ejecucion = threading.Thread(target=self.ejecutar_procesos)
        self.hilo_ejecucion.daemon = True
        self.hilo_ejecucion.start()
    
    def pausar_proceso(self):
        """Pausar el proceso"""
        global PAUSADO
        if EJECUTANDO and not PAUSADO:
            PAUSADO = True
            self.estado.set("Pausado")
            self.estado_label.configure(foreground="orange")
            self.log("⏸️ Proceso pausado")
            self.actualizar_estado_botones()
    
    def reanudar_proceso(self):
        """Reanudar el proceso después de 5 segundos"""
        global PAUSADO
        if EJECUTANDO and PAUSADO:
            self.estado.set("Reanudando en 5 segundos...")
            self.estado_label.configure(foreground="blue")
            
            for i in range(5, 0, -1):
                self.estado.set(f"Reanudando en {i} segundos...")
                time.sleep(1)
                
            PAUSADO = False
            self.estado.set("Ejecutando...")
            self.estado_label.configure(foreground="green")
            self.log("▶️ Proceso reanudado")
            self.actualizar_estado_botones()
    
    def detener_proceso(self):
        """Detener completamente el proceso"""
        global EJECUTANDO, PAUSADO, LINEA_ACTUAL
        EJECUTANDO = False
        PAUSADO = False
        LINEA_ACTUAL = 0
        
        self.estado.set("Detenido")
        self.estado_label.configure(foreground="red")
        self.log("⏹️ Proceso detenido")
        self.actualizar_estado_botones()
        self.actualizar_estado_lineas()
    
    def actualizar_estado_botones(self):
        """Actualizar estado de los botones según el estado del proceso"""
        global EJECUTANDO, PAUSADO
        
        if EJECUTANDO:
            self.btn_iniciar.config(state=tk.DISABLED)
            self.btn_detener.config(state=tk.NORMAL)
            
            if PAUSADO:
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
        """Ejecutar los procesos secuencialmente para cada línea"""
        # MOVER LA DECLARACIÓN GLOBAL AL INICIO DE LA FUNCIÓN
        global EJECUTANDO, PAUSADO, LINEA_ACTUAL, LINEA_MAXIMA, KML_FILENAME
        
        try:
            while LINEA_ACTUAL <= LINEA_MAXIMA and EJECUTANDO:
                # Verificar pausa
                while PAUSADO and EJECUTANDO:
                    time.sleep(0.1)
                    
                if not EJECUTANDO:
                    break
                    
                self.log(f"🔄 Procesando línea {LINEA_ACTUAL}/{LINEA_MAXIMA}")
                self.actualizar_estado_lineas()
                
                # Ejecutar Programa 1
                self.log("Iniciando Programa 1 - Procesador CSV")
                resultado1, linea_procesada = ejecutar_programa1_interfaz(LINEA_ACTUAL, self.log)
                
                if not resultado1 or not EJECUTANDO:
                    if not EJECUTANDO:
                        break
                    self.log(f"❌ Programa 1 falló en línea {LINEA_ACTUAL}")
                    LINEA_ACTUAL += 1
                    continue
                
                # Ejecutar Programa 2
                self.log("Iniciando Programa 2 - Automatización NSE")
                resultado2 = ejecutar_programa2_interfaz(linea_procesada, self.log)
                
                if not resultado2 or not EJECUTANDO:
                    if not EJECUTANDO:
                        break
                    self.log(f"❌ Programa 2 falló en línea {LINEA_ACTUAL}")
                    LINEA_ACTUAL += 1
                    continue
                
                # Ejecutar Programa 3
                self.log("Iniciando Programa 3 - Servicios NSE")
                resultado3 = ejecutar_programa3_interfaz(linea_procesada, self.log)
                
                if not resultado3 or not EJECUTANDO:
                    if not EJECUTANDO:
                        break
                    self.log(f"❌ Programa 3 falló en línea {LINEA_ACTUAL}")
                    LINEA_ACTUAL += 1
                    continue
                
                # Ejecutar Programa 4
                self.log("Iniciando Programa 4 - Automatización GE")
                resultado4 = ejecutar_programa4_interfaz(linea_procesada, KML_FILENAME, self.log)
                
                if resultado4:
                    self.log(f"✅ Línea {LINEA_ACTUAL} procesada exitosamente")
                else:
                    self.log(f"⚠️ Línea {LINEA_ACTUAL} completada con advertencias")
                
                LINEA_ACTUAL += 1
                self.actualizar_estado_lineas()
                
                # Pequeña pausa entre líneas
                time.sleep(2)
            
            if EJECUTANDO and LINEA_ACTUAL > LINEA_MAXIMA:
                self.log("🎉 Proceso completado exitosamente")
                self.estado.set("Completado")
                self.estado_label.configure(foreground="green")
            elif not EJECUTANDO:
                self.log("Proceso detenido por el usuario")
                
        except Exception as e:
            self.log(f"❌ Error en ejecución: {e}")
            self.estado.set("Error")
            self.estado_label.configure(foreground="red")
        
        finally:
            # LAS VARIABLES GLOBALES YA ESTÁN DECLARADAS AL INICIO
            EJECUTANDO = False
            PAUSADO = False
            self.actualizar_estado_botones()
# Funciones de interfaz para los programas existentes
def ejecutar_programa1_interfaz(linea_especifica, log_func):
    """Versión del Programa 1 para la interfaz"""
    try:
        log_func("=" * 60)
        log_func("INICIANDO PROGRAMA 1 - PROCESADOR CSV")
        log_func("=" * 60)
        
        archivo_csv = CSV_FILE
        procesador = ProcesadorCSV(archivo_csv)
        
        log_func("Iniciando procesamiento automático del Programa 1...")
        time.sleep(3)
        
        # Cargar CSV y configurar línea específica
        if not procesador.cargar_csv():
            return False, None
            
        # Modificar para usar línea específica
        procesador.df = procesador.df.iloc[linea_especifica-1:linea_especifica]
        
        resultado, linea_procesada = procesador.procesar_todo()
        
        if resultado and linea_procesada:
            log_func(f"✅ Programa 1 completado. Línea procesada: {linea_procesada}")
            return True, linea_especifica
        else:
            log_func("❌ Programa 1 falló")
            return False, None
            
    except Exception as e:
        log_func(f"❌ Error en Programa 1: {e}")
        return False, None

def ejecutar_programa2_interfaz(linea_especifica, log_func):
    """Versión del Programa 2 para la interfaz"""
    try:
        log_func("\n" + "=" * 60)
        log_func("INICIANDO PROGRAMA 2 - AUTOMATIZACIÓN NSE")
        log_func("=" * 60)
        
        nse = NSEAutomation(linea_especifica=linea_especifica)
        nse.is_running = True
        
        if not os.path.exists(nse.csv_file):
            log_func(f"❌ ERROR: Archivo CSV no encontrado: {nse.csv_file}")
            return False
        
        log_func(f"🎯 Procesando línea: {linea_especifica}")
        time.sleep(3)
        
        resultado = nse.execute_nse_script()
        
        if resultado:
            log_func("✅ Programa 2 finalizado exitosamente")
        else:
            log_func("❌ Programa 2 falló")
            
        return resultado
        
    except Exception as e:
        log_func(f"❌ Error en Programa 2: {e}")
        return False

def ejecutar_programa3_interfaz(linea_especifica, log_func):
    """Versión del Programa 3 para la interfaz"""
    try:
        log_func("\n" + "=" * 60)
        log_func("INICIANDO PROGRAMA 3 - SERVICIOS NSE")
        log_func("=" * 60)
        
        nse_services = NSEServicesAutomation(linea_especifica=linea_especifica)
        
        if not os.path.exists(nse_services.csv_file):
            log_func(f"❌ ERROR: Archivo CSV no encontrado: {nse_services.csv_file}")
            return False
        
        log_func(f"🎯 Procesando línea: {linea_especifica}")
        
        if not nse_services.iniciar_ahk():
            log_func("❌ No se pudieron iniciar los servicios AHK")
            return False
        
        nse_services.is_running = True
        resultado = nse_services.procesar_linea_especifica()
        
        if resultado:
            log_func(f"✅ Programa 3 completado exitosamente")
        else:
            log_func(f"❌ Programa 3 falló")
        
        nse_services.detener_ahk()
        return resultado
        
    except Exception as e:
        log_func(f"❌ Error en Programa 3: {e}")
        return False

def ejecutar_programa4_interfaz(linea_especifica, kml_filename, log_func):
    """Versión del Programa 4 para la interfaz"""
    try:
        log_func("\n" + "=" * 60)
        log_func("INICIANDO PROGRAMA 4 - AUTOMATIZACIÓN GE")
        log_func("=" * 60)
        
        ge_auto = GEAutomation(linea_especifica=linea_especifica)
        ge_auto.is_running = True
        
        # Actualizar nombre KML
        ge_auto.nombre = f"{kml_filename}.kml"
        
        if not os.path.exists(ge_auto.csv_file):
            log_func(f"❌ ERROR: Archivo CSV no encontrado: {ge_auto.csv_file}")
            return False
        
        log_func(f"🎯 Procesando línea: {linea_especifica}")
        log_func(f"📁 Archivo KML: {ge_auto.nombre}")
        
        time.sleep(3)
        
        success = ge_auto.perform_actions()
        
        if success:
            log_func("✅ Programa 4 finalizado exitosamente")
        else:
            log_func("❌ Programa 4 falló")
            
        return success
        
    except Exception as e:
        log_func(f"❌ Error en Programa 4: {e}")
        return False

# Clases originales (sin cambios)
class ProcesadorCSV:
    def __init__(self, archivo_csv):
        self.archivo_csv = archivo_csv
        self.df = None
        self.ahk_manager = AHKManagerCD()
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        
    def cargar_csv(self):
        """Carga el archivo CSV"""
        try:
            self.df = pd.read_csv(self.archivo_csv)
            logger.info(f"CSV cargado correctamente: {len(self.df)} registros")
            return True
        except Exception as e:
            logger.error(f"Error cargando CSV: {e}")
            return False
    
    def iniciar_ahk(self):
        """Inicia todos los procesos AHK"""
        logger.info("Iniciando procesos AHK...")
        return (self.ahk_manager.start_ahk() and 
                self.ahk_writer.start_ahk() and 
                self.ahk_click_down.start_ahk())
    
    def detener_ahk(self):
        """Detiene todos los procesos AHK"""
        logger.info("Deteniendo procesos AHK...")
        self.ahk_manager.stop_ahk()
        self.ahk_writer.stop_ahk()
        self.ahk_click_down.stop_ahk()
    
    def buscar_por_id(self, id_buscar):
        """Busca un ID en la primera columna del CSV"""
        if self.df is None:
            logger.error("CSV no cargado")
            return None
            
        # Buscar en la primera columna (asumimos que es la columna 0)
        resultado = self.df[self.df.iloc[:, 0] == id_buscar]
        
        if len(resultado) == 0:
            logger.warning(f"ID {id_buscar} no encontrado en el CSV")
            return None
        
        logger.info(f"ID {id_buscar} encontrado, datos: {resultado.iloc[0].tolist()}")
        return resultado.iloc[0]
    
    def procesar_registro(self):
        """Ejecuta el flujo completo para un registro"""
        try:
            # Paso 2: Click en (89, 263)
            logger.info("Paso 2: Click en (89, 263)")
            pyautogui.click(89, 263)
            time.sleep(1)
            
            # Paso 3: Usar AHKManager en (1483, 519) para obtener ID
            logger.info("Paso 3: Obteniendo ID con AHKManager en (1483, 519)")
            id_obtenido = self.ahk_manager.ejecutar_acciones_ahk(1483, 519)
            
            if not id_obtenido:
                logger.error("No se pudo obtener el ID")
                return False, None
            
            id_obtenido = int(id_obtenido)
            logger.info(f"ID obtenido: {id_obtenido}")

            # Escribir de nuevo el ID obtenido (1483,519)
            self.ahk_writer.ejecutar_escritura_ahk(1483, 519, str(id_obtenido))
            time.sleep(1)
            
            id_obtenido = self.ahk_manager.ejecutar_acciones_ahk(1483, 519)
            
            if not id_obtenido:
                logger.error("No se pudo obtener el ID")
                return False, None
            
            id_obtenido = int(id_obtenido)
            logger.info(f"ID obtenido: {id_obtenido}")
            # Paso 4: Buscar el ID en el CSV
            logger.info(f"Paso 4: Buscando ID {id_obtenido} en CSV")
            registro = self.buscar_por_id(id_obtenido)
            
            if registro is None:
                logger.error(f"ID {id_obtenido} no encontrado en CSV")
                return False, None
            
            # Obtener el número de línea del registro encontrado
            linea_procesada = None
            for idx in range(len(self.df)):
                if self.df.iloc[idx, 0] == id_obtenido:
                    linea_procesada = idx + 1  # +1 porque las líneas empiezan en 1
                    break
            
            # Paso 5: Escribir valor de columna 2 en (1483, 519)
            if len(registro) >= 2:  # Verificar que existe columna 2
                valor_columna_2 = str(registro.iloc[1])
                logger.info(f"Paso 5: Escribiendo valor '{valor_columna_2}' en (1483, 519)")
                
                exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, valor_columna_2)
                if not exito_escritura:
                    logger.error("Error en la escritura")
                    return False, linea_procesada
            else:
                logger.warning("No hay columna 2 en el registro")
            
            # Paso 6: Revisar si columna 4 es mayor a 0
            if len(registro) >= 4:  # Verificar que existe columna 4
                valor_columna_4 = registro.iloc[3]
                logger.info(f"Paso 6: Valor columna 4 = {valor_columna_4}")
                
                # Paso 7: Si es mayor a 0, usar AHKClickDown
                if pd.notna(valor_columna_4) and float(valor_columna_4) > 0:
                    veces_down = int(float(valor_columna_4))
                    logger.info(f"Paso 7: Ejecutando {veces_down} veces DOWN en (1507, 636)")
                    
                    exito_down = self.ahk_click_down.ejecutar_click_down(1507, 636, veces_down)
                    if not exito_down:
                        logger.error("Error en click + down")
                        return False, linea_procesada
                else:
                    logger.info("Paso 7: Saltado (columna 4 <= 0)")
            else:
                logger.warning("No hay columna 4 en el registro")
            
            # Paso 8: Click en (1290, 349)
            logger.info("Paso 8: Click en (1290, 349)")
            pyautogui.click(1290, 349)
            time.sleep(1)
            
            logger.info("Procesamiento completado exitosamente")
            return True, linea_procesada
            
        except Exception as e:
            logger.error(f"Error en procesar_registro: {e}")
            return False, None
    
    def procesar_todo(self, pausa_entre_registros=2):
        """Procesa múltiples registros (si es necesario)"""
        if not self.cargar_csv():
            return False, None
            
        if not self.iniciar_ahk():
            return False, None
        
        try:
            # Este método procesa un registro por ejecución
            logger.info("Iniciando procesamiento de registro...")
            exito, linea_procesada = self.procesar_registro()
            
            if exito:
                logger.info(f"Procesamiento completado. Línea procesada: {linea_procesada}")
            else:
                logger.error("Procesamiento falló")
                
            return exito, linea_procesada
            
        finally:
            # Siempre detener AHK al finalizar
            self.detener_ahk()

class NSEAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica  # Línea específica a procesar (1-indexed)
        self.csv_file = CSV_FILE
        self.reference_image = "img/VentanaAsignar.png"
        self.is_running = False
        
        # Inicializar AHKWriter
        self.ahk_writer = AHKWriter()
        
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # COORDENADAS RELATIVAS (de la tabla verde) - AJUSTADAS PARA COLUMNAS 7-17
        # Estas coordenadas serán sumadas a la posición de la imagen detectada
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
        
        # Coordenadas para botones (relativas)
        self.coords_asignar = [446, 281]  # Botón asignar en la ventana
        self.coords_cerrar = [396, 352]   # Botón cerrar

    def click(self, x, y, duration=0.1):
        """Hacer clic en coordenadas específicas"""
        pyautogui.click(x, y, duration=duration)
        time.sleep(0.5)

    def write_with_ahk(self, x, y, text):
        """Escribir texto usando AHKWriter"""
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, text)
        if not success:
            print(f"❌ Error al escribir con AHK en ({x}, {y}): {text}")
        return success

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

    def detect_image_with_cv2(self, image_path, confidence=0.6):
        """Detectar imagen en pantalla usando template matching con OpenCV"""
        try:
            # Capturar pantalla completa
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # Cargar template de la imagen de referencia
            template = cv2.imread(image_path)
            if template is None:
                print(f"Error: No se pudo cargar la imagen {image_path}")
                return False, None
            
            # Realizar template matching
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            # Umbral de confianza
            if max_val < confidence:
                print(f"Imagen no encontrada. Mejor coincidencia: {max_val:.2f}")
                return False, None
            
            print(f"Imagen encontrada con confianza: {max_val:.2f}")
            return True, max_loc  # Devuelve las coordenadas (x, y) de la esquina superior izquierda
        except Exception as e:
            print(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.6):
        """Esperar a que aparezca una imagen con múltiples intentos usando OpenCV"""
        print(f"🔍 Buscando imagen: {image_path}")
        
        for attempt in range(1, max_attempts + 1):
            found, location = self.detect_image_with_cv2(image_path, confidence)
            
            if found:
                print(f"✅ Imagen detectada en el intento {attempt} en coordenadas: {location}")
                return True, location
            
            print(f"⏳ Intento {attempt}/{max_attempts} - Imagen no encontrada")
            
            # Espera normal de 2 segundos entre intentos
            if attempt < max_attempts:
                if attempt % 10 == 0:
                    # Cada 10 intentos, esperar 10 segundos
                    print("⏰ Espera prolongada de 10 segundos...")
                    time.sleep(10)
                else:
                    # Espera normal de 2 segundos
                    time.sleep(2)
        
        print("❌ Imagen no encontrada después de 30 intentos. Terminando proceso.")
        return False, None

    def should_skip_process(self, row):
        """Determina si se debe saltar el proceso basado en la columna 6"""
        # Columna 6 es el índice 5 en base 0
        if pd.notna(row.iloc[5]):
            col_value = str(row.iloc[5]).strip()
            # Si la columna 6 tiene algún valor (no vacío y no NaN), se salta el proceso
            if col_value and col_value != "" and col_value != "nan":
                return True
        return False

    def execute_nse_script(self):
        """Función principal de ejecución NSE - Proceso único"""
        # Iniciar AHKWriter
        if not self.ahk_writer.start_ahk():
            print("❌ No se pudo iniciar AHKWriter")
            return False
            
        try:
            # Leer CSV
            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            # Validar línea específica
            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False
            
            # Procesar solo la línea específica (start_count - 1)
            row = df.iloc[self.linea_especifica - 1]
            print(f"🔄 Procesando línea {self.linea_especifica}/{total_lines}")
            
            # Verificar si se debe saltar el proceso (columna 6 tiene valor)
            if self.should_skip_process(row):
                print(f"⏭️  Saltando línea {self.linea_especifica} - Columna 6 tiene valor: {row.iloc[5]}")
                return True
            
            # Verificar que sea tipo V
            if str(row.iloc[4]).strip().upper() != "V":
                print(f"⚠️  Saltando línea {self.linea_especifica} - No es tipo V: {row.iloc[4]}")
                return True
            
            # click en el boton seleccionar lote 
            self.click(169, 189)
            self.sleep(2)
            # click en el boton asignar nse
            self.click(1491, 386)
            self.sleep(2)
            
            # ESPACIO PARA DETECCIÓN DE IMAGEN CON REINTENTOS
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=30)
            
            if not image_found:
                print("❌ No se puede continuar sin detectar la imagen de referencia.")
                return False
            
            # Si se encontró la imagen, continuar con el proceso usando las coordenadas base
            print("🎯 Imagen detectada, procediendo con tipo V")
            self.handle_type_v(row, base_location)
            
            print(f"✅ Línea {self.linea_especifica} completada (hasta CERRAR)")
            return True
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            return False
        finally:
            # Detener AHKWriter
            self.ahk_writer.stop_ahk()

    def handle_type_v(self, row, base_location):
        """Manejar tipo V con coordenadas relativas - COLUMNAS 7-17"""
        # Calcular coordenadas absolutas sumando las relativas a la posición base
        base_x, base_y = base_location
        
        # Lógica V para columnas 7-17 con coordenadas relativas
        # Nota: row[6] a row[16] corresponden a columnas 7-17 (índices base 0)
        for col_index in range(7, 18):  # 7 a 17 inclusive
            if pd.notna(row.iloc[col_index-1]) and row.iloc[col_index-1] > 0:
                # Usar coordenadas relativas de la tabla verde, sumando a la base
                x_cs_rel, y_cs_rel = self.coords_select[col_index]
                x_ct_rel, y_ct_rel = self.coords_type[col_index]
                
                # Calcular coordenadas absolutas
                x_cs_abs = base_x + x_cs_rel
                y_cs_abs = base_y + y_cs_rel
                x_ct_abs = base_x + x_ct_rel
                y_ct_abs = base_y + y_ct_rel
                
                self.click(x_cs_abs, y_cs_abs)
                self.sleep(2)
                
                # Usar AHKWriter para escribir en lugar de pyautogui
                texto = str(int(row.iloc[col_index-1]))
                self.write_with_ahk(x_ct_abs, y_ct_abs, texto)
                self.sleep(2)
        
        # Botón ASIGNAR antes de cerrar (coordenadas absolutas)
        x_asignar_rel, y_asignar_rel = self.coords_asignar
        x_asignar_abs = base_x + x_asignar_rel
        y_asignar_abs = base_y + y_asignar_rel
        self.click(x_asignar_abs, y_asignar_abs)
        self.sleep(2)
        
        # Botón CERRAR (coordenadas absolutas)
        x_cerrar_rel, y_cerrar_rel = self.coords_cerrar
        x_cerrar_abs = base_x + x_cerrar_rel
        y_cerrar_abs = base_y + y_cerrar_rel
        self.click(x_cerrar_abs, y_cerrar_abs)
        self.sleep(2)

class NSEServicesAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica
        self.csv_file = CSV_FILE
        self.current_line = 0
        self.is_running = False
        self.reference_point = None
        
        # Inicializar controladores AHK
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        self.ahk_enter = EnterAHKManager()
        
        # Configurar coordenadas base
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
        """Busca una imagen en la pantalla usando OpenCV"""
        logging.info(f"🔍 Buscando imagen: {imagen_path}")
        
        try:
            # Cargar la imagen template
            template = cv2.imread(imagen_path)
            if template is None:
                logging.error(f"❌ No se pudo cargar la imagen: {imagen_path}")
                return None
            
            template_height, template_width = template.shape[:2]
            
            for intento in range(timeout):
                # Capturar screenshot de toda la pantalla
                screenshot = ImageGrab.grab()
                screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                # Realizar la búsqueda de la plantilla
                result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confidence:
                    # Encontrado - retornar coordenadas de la esquina superior izquierda
                    x, y = max_loc
                    logging.info(f"✅ Imagen encontrada en intento {intento + 1} - Coordenadas: ({x}, {y}) - Confianza: {max_val:.2f}")
                    return (x, y)
                
                logging.info(f"⏳ Intento {intento + 1}/{timeout} - Confianza máxima: {max_val:.2f}")
                time.sleep(1)
            
            logging.error(f"❌ No se encontró la imagen después de {timeout} intentos")
            return None
            
        except Exception as e:
            logging.error(f"❌ Error en búsqueda de imagen: {e}")
            return None

    def actualizar_coordenadas_relativas(self, referencia):
        """Actualiza todas las coordenadas para que sean relativas al punto de referencia"""
        if referencia is None:
            logging.error("❌ No se puede actualizar coordenadas: referencia es None")
            return False
        
        ref_x, ref_y = referencia
        
        # Actualizar coordenadas relativas
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
        
        # Mantener coordenadas que no cambian
        self.coords_relativas['boton_error'] = self.coords['boton_error']
        self.coords_relativas['inicio_servicios'] = self.coords['inicio_servicios']
        
        self.reference_point = referencia
        logging.info("✅ Coordenadas actualizadas a relativas")
        return True

    def iniciar_ahk(self):
        """Iniciar todos los servicios AHK"""
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
        """Detener todos los servicios AHK"""
        try:
            self.ahk_writer.stop_ahk()
            self.ahk_click_down.stop_ahk()
            self.ahk_enter.stop_ahk()
            logging.info("✅ Todos los servicios AHK detenidos correctamente")
        except Exception as e:
            logging.error(f"Error deteniendo servicios AHK: {e}")

    def click(self, x, y, duration=0.1):
        """Hacer clic en coordenadas específicas"""
        pyautogui.click(x, y, duration=duration)
        time.sleep(0.5)

    def write(self, text):
        """Escribir texto usando AHK Writer"""
        try:
            # Usar coordenadas relativas si están disponibles
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                campo_coords = self.coords_relativas['campo_cantidad']
                logger.info(f"Usando coordenadas relativas para escribir: {campo_coords}")
            else:
                campo_coords = self.coords['campo_cantidad']
                logger.info(f"Usando coordenadas absolutas para escribir: {campo_coords}")

            # Usar AHKWriter para escribir
            success = self.ahk_writer.ejecutar_escritura_ahk(
                campo_coords[0],
                campo_coords[1],
                str(text)
            )
            if success:
                logging.info(f"✅ Texto escrito exitosamente: {text}")
            else:
                logging.error(f"❌ Error al escribir texto: {text}")
            return success
        
        except Exception as e:
            logging.error(f"Error escribiendo texto '{text}': {e}")
            return False

    def press_down(self, x, y, times=1):
        """Presionar flecha down usando AHK"""
        try:
            # Usar coordenadas relativas si están disponibles
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                click_coords = (x, y)
            else:
                click_coords = (x, y)
                
            # Usamos AHK Click Down con las veces especificadas
            return self.ahk_click_down.ejecutar_click_down(click_coords[0], click_coords[1], times)
        except Exception as e:
            logging.error(f"Error presionando DOWN {times} veces: {e}")
            return False

    def press_enter(self):
        """Presionar enter usando AHK"""
        try:                
            return self.ahk_enter.presionar_enter(1)
        except Exception as e:
            logging.error(f"Error presionando enter: {e}")
            return False

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

    def handle_error_click(self):
        """Manejar clics de error"""
        for _ in range(5):
            # Usar coordenadas relativas si están disponibles para boton_error
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                self.click(*self.coords_relativas['boton_error'])
            else:
                self.click(*self.coords['boton_error'])
            self.sleep(2)

    def procesar_linea_especifica(self):
        """Procesar solo una línea específica del CSV"""
        try:
            # Leer CSV
            df = pd.read_csv(self.csv_file, encoding='utf-8')
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            # Validar línea específica
            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica <= 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False
            
            # Obtener la línea específica (ajustar índice ya que CSV empieza en 0 para datos)
            linea_idx = self.linea_especifica - 1  # Convertir a índice base 0
            self.current_line = self.linea_especifica
            
            print(f"🎯 PROCESANDO LÍNEA ESPECÍFICA: {linea_idx}/{total_lines}")
            logger.info(f"😭 PROCESANDO LÍNEA ESPECÍFICA: {linea_idx}/{total_lines}")
            
            row = df.iloc[linea_idx]
            
            # Solo procesar servicios si la columna 18 tiene valor > 0
            if pd.notna(row.iloc[17]) and row.iloc[17] > 0:
                print(f"✅ Línea {self.current_line} tiene servicios para procesar")
                
                self.click(*self.coords['inicio_servicios'])
                self.sleep(2)
                
                # BUSCAR IMAGEN Y ACTUALIZAR COORDENADAS
                print("🔍 Buscando ventana de servicios...")
                referencia = self.buscar_imagen("img/ventanaAdministracion4.PNG", timeout=30)
                
                if referencia is None:
                    print("❌ ERROR: No se pudo encontrar la ventana de servicios")
                    return False
                
                # Actualizar coordenadas relativas
                if not self.actualizar_coordenadas_relativas(referencia):
                    print("❌ ERROR: No se pudieron actualizar las coordenadas relativas")
                    return False
                
                # Continuar con el procesamiento normal usando coordenadas relativas
                self.click(*self.coords_relativas['menu_principal'])
                self.sleep(2)
                    
                # Llamar a funciones de servicios
                servicios_procesados = 0
                logger.info(f"Procesando servicios para línea {self.current_line}")
                logger.info(f"😭Datos de la línea {self.current_line}: {row.iloc[18:27].tolist()}")
                
                if pd.notna(row.iloc[18]) and row.iloc[18] > 0:  # VOZ COBRE TELMEX
                    logger.info(f"  └─ Procesando VOZ COBRE TELMEX: {row.iloc[18]}")
                    self.handle_voz_cobre(int(row.iloc[18]))
                    servicios_procesados += 1
                    logger.info(f"  └─ VOZ COBRE TELMEX procesado")
                    
                if pd.notna(row.iloc[19]) and row.iloc[19] > 0:  # Datos s/dom
                    print(f"  └─ Procesando DATOS S/DOM: {row.iloc[19]}")
                    logger.info(f"  └─ Procesando DATOS S/DOM: {row.iloc[19]}")
                    self.handle_datos_sdom(row.iloc[19])
                    servicios_procesados += 1
                    
                if pd.notna(row.iloc[20]) and row.iloc[20] > 0:  # Datos-cobre-telmex-inf
                    print(f"  └─ Procesando DATOS COBRE TELMEX: {row.iloc[20]}")
                    logger.info(f"  └─ Procesando DATOS COBRE TELMEX: {row.iloc[20]}")
                    self.handle_datos_cobre_telmex(row.iloc[20])
                    servicios_procesados += 1
                    
                if pd.notna(row.iloc[21]) and row.iloc[21] > 0:  # Datos-fibra-telmex-inf
                    print(f"  └─ Procesando DATOS FIBRA TELMEX: {row.iloc[21]}")
                    logger.info(f"  └─ Procesando DATOS FIBRA TELMEX: {row.iloc[21]}")
                    self.handle_datos_fibra_telmex(row.iloc[21])
                    servicios_procesados += 1
                    
                if pd.notna(row.iloc[22]) and row.iloc[22] > 0:  # TV cable otros
                    print(f"  └─ Procesando TV CABLE OTROS: {row.iloc[22]}")
                    logger.info(f"  └─ Procesando TV CABLE OTROS: {row.iloc[22]}")
                    self.handle_tv_cable_otros(row.iloc[22])
                    servicios_procesados += 1
                    
                if pd.notna(row.iloc[23]) and row.iloc[23] > 0:  # Dish
                    print(f"  └─ Procesando DISH: {row.iloc[23]}")
                    logger.info(f"  └─ Procesando DISH: {row.iloc[23]}")
                    self.handle_dish(row.iloc[23])
                    servicios_procesados += 1
                    
                if pd.notna(row.iloc[24]) and row.iloc[24] > 0:  # TVS
                    print(f"  └─ Procesando TVS: {row.iloc[24]}")
                    logger.info(f"  └─ Procesando TVS: {row.iloc[24]}")
                    self.handle_tvs(row.iloc[24])
                    servicios_procesados += 1
                    
                if pd.notna(row.iloc[25]) and row.iloc[25] > 0:  # SKY
                    print(f"  └─ Procesando SKY: {row.iloc[25]}")
                    logger.info(f"  └─ Procesando SKY: {row.iloc[25]}")
                    self.handle_sky(row.iloc[25])
                    servicios_procesados += 1
                    
                if pd.notna(row.iloc[26]) and row.iloc[26] > 0:  # VETV
                    print(f"  └─ Procesando VETV: {row.iloc[26]}")
                    logger.info(f"  └─ Procesando VETV: {row.iloc[26]}")
                    self.handle_vetv(row.iloc[26])
                    servicios_procesados += 1
                
                # Usar coordenadas relativas para el cierre
                self.click(*self.coords_relativas['cierre'])
                self.sleep(5)
                
                print(f"✅ Línea {self.current_line} completada: {servicios_procesados} servicios procesados")
                return True
            else:
                print(f"⏭️  Línea {self.current_line} no tiene servicios para procesar")
                return True  # Consideramos éxito si no hay servicios para procesar
            
        except Exception as e:
            print(f"❌ Error procesando línea {self.current_line}: {e}")
            logging.error(f"Error en procesar_linea_especifica: {e}")
            return False

    def handle_voz_cobre(self, cantidad):
        """Manejar servicio VOZ COBRE TELMEX"""
        # Usar coordenadas relativas
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.write(cantidad)
        logger.info(f"Escribio cantidad de VOZ COBRE TELMEX: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_sdom(self, cantidad):
        """Manejar servicio DATOS S/DOM"""
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de DATOS S/DOM: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_cobre_telmex(self, cantidad):
        """Manejar servicio DATOS COBRE TELMEX"""
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_producto'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de DATOS COBRE TELMEX: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_fibra_telmex(self, cantidad):
        """Manejar servicio DATOS FIBRA TELMEX"""
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 1)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de DATOS FIBRA TELMEX: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_tv_cable_otros(self, cantidad):
        """Manejar servicio TV CABLE OTROS"""
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 4)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de TV CABLE OTROS: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_dish(self, cantidad):
        """Manejar servicio DISH"""
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de DISH: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_tvs(self, cantidad):
        """Manejar servicio TVS"""
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 2)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de TVS: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_sky(self, cantidad):
        """Manejar servicio SKY"""
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 3)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de SKY: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_vetv(self, cantidad):
        """Manejar servicio VETV"""
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 5)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        logger.info(f"Escribio cantidad de VETV: {cantidad}")
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

class GEAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica  # Línea específica a procesar
        self.csv_file = CSV_FILE
        self.reference_image = "img/textoAdicional.PNG"
        self.ventana_archivo_img = "img/cargarArchivo.png"
        self.ventana_error_img = "img/ventanaError.png"
        self.is_running = False
        
        # Inicializar todos los manejadores AHK
        self.ahk_writer = AHKWriter()  # Para escritura de texto
        self.ahk_manager = AHKManager()  # Para manejar ventanas de archivo
        self.enter = EnterAHKManager()  # Para presionar Enter
        self.ahk_click_down = AHKClickDown()  # Para flechas abajo
        
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5

        self.nombre=""
        
        # COORDENADAS ABSOLUTAS (solo las necesarias)
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
        
        # Coordenadas relativas para detección de imagen (campo de texto)
        self.coords_texto_relativas = {
            'campo_texto': (230, 66),
            'agregar_texto': (64, 100),
            'cerrar_ventana_texto': (139, 98)
        }

    def encontrar_ventana_archivo(self):
        """Busca la ventana de archivo usando template matching con reintentos inteligentes"""
        intentos = 1
        confianza_minima = 0.6
        tiempo_espera_base = 1
        tiempo_espera_largo = 10
        
        # Cargar template una sola vez fuera del bucle
        template = cv2.imread(self.ventana_archivo_img)
        if template is None:
            logger.error(f"No se pudo cargar la imagen '{self.ventana_archivo_img}'")
            return None
        
        while self.is_running: 
            try:
                # Capturar pantalla completa
                screenshot = pyautogui.screenshot()
                pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                # Realizar template matching
                result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confianza_minima:
                    logger.info(f"Ventana encontrada con confianza: {max_val:.2f}")
                    # Devolver tupla (x, y)
                    return max_loc
                else:
                    # Estrategia de espera progresiva
                    if intentos % 10 == 0 and intentos > 0:
                        logger.info(f"Intento {intentos}: Mejor coincidencia: {max_val:.2f}")
                        logger.info("Esperando 10 segundos...")
                        time.sleep(tiempo_espera_largo)
                    else:
                        time.sleep(tiempo_espera_base)
                    intentos += 1
                    
            except Exception as e:
                logger.error(f"Error durante la búsqueda: {e}")
                time.sleep(tiempo_espera_base)
                intentos += 1

        return None

    def detectar_ventana_error(self):
        """
        Detecta la ventana de error y presiona Enter para cerrarla
        Returns:
            bool: True si encontró la ventana de error, False en caso contrario
        """
        try:
            # Cargar template de la ventana de error
            template = cv2.imread(self.ventana_error_img) 
            if template is None:
                logger.error(f"No se pudo cargar la imagen '{self.ventana_error_img}'")
                return False
            
            # Capturar pantalla completa
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # Realizar template matching
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            confianza_minima = 0.6
            
            if max_val >= confianza_minima:
                logger.info(f"Ventana de error detectada con confianza: {max_val:.2f}")
                
                # Presionar Enter para cerrar la ventana de error usando AHK
                if not self.enter.start_ahk():
                    logger.error("No se pudo iniciar AutoHotkey")
                    return False
                    
                # Enviar comandos a AHK
                if self.enter.presionar_enter(1):
                    time.sleep(2.5)
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
        """Maneja el comportamiento especial para cargar archivos usando AHK Manager"""
        # Buscar la ventana de archivo
        coordenadas_ventana = self.encontrar_ventana_archivo()

        if coordenadas_ventana:
            x_ventana, y_ventana = coordenadas_ventana
            logger.info(f"Coordenadas ventana: x={x_ventana}, y={y_ventana}")
            
            # Calcular coordenadas del campo de texto
            x_campo = x_ventana + 294
            y_campo = y_ventana + 500
            logger.info(f"Coordenadas campo texto: x={x_campo}, y={y_campo}")
            
            # Iniciar AHK Manager para escribir el nombre del archivo
            if not self.ahk_manager.start_ahk():
                logger.error("No se pudo iniciar AutoHotkey")
                return False
            
            # Enviar comandos a AHK Manager para escribir el nombre del archivo
            if self.ahk_manager.ejecutar_acciones_ahk(x_campo, y_campo, nombre_archivo):
                time.sleep(1.5)  # Esperar a que AHK termine
            else:
                logger.error("Error enviando comando a AHK")
                return False
            
            self.ahk_manager.stop_ahk()
            return True
        else:
            logger.error("No se pudo encontrar la ventana de archivo.")
            return False

    def escribir_texto_adicional_ahk(self, x, y, texto):
        """Escribe texto adicional usando AHK Writer"""
        if not texto or pd.isna(texto) or str(texto).strip() == '':
            print("⚠️  Texto adicional vacío, saltando escritura")
            return True
            
        print(f"📝 Intentando escribir texto: '{texto}' en coordenadas ({x}, {y})")
        
        # Verificar que las coordenadas son válidas
        if x <= 0 or y <= 0:
            print(f"❌ Coordenadas inválidas: ({x}, {y})")
            return False
        
        if not self.ahk_writer.start_ahk():
            logger.error("No se pudo iniciar AHK Writer")
            print("❌ Falló al iniciar AHK Writer")
            return False
        
        print("🔄 AHK Writer iniciado, enviando comando...")
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, texto)
        self.ahk_writer.stop_ahk()
        
        if success:
            print(f"✅ Texto escrito exitosamente: '{texto}'")
        else:
            print(f"❌ Error al escribir texto: '{texto}'")
            print("🔄 Intentando método alternativo con pyautogui...")
            try:
                # Método de fallback
                self.click(x, y)
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'a')  # Seleccionar todo
                time.sleep(0.5)
                pyautogui.press('delete')  # Borrar
                time.sleep(0.5)
                pyautogui.write(texto, interval=0.05)  # Escribir
                print(f"✅ Texto escrito con pyautogui: '{texto}'")
                success = True
            except Exception as e:
                print(f"❌ También falló pyautogui: {e}")
                
        return success

    def presionar_flecha_abajo_ahk(self, x,y,veces=1):
        """Presiona flecha abajo usando AHK"""
        if not self.ahk_click_down.start_ahk():
            logger.error("No se pudo iniciar AutoHotkey para flecha abajo")
            return False
        
        try:
            self.ahk_click_down.ejecutar_click_down(x, y, veces)  # Coordenadas dummy
            return True
        except Exception as e:
            logger.error(f"Error presionando flecha abajo: {e}")
            return False
        finally:
            self.ahk_click_down.stop_ahk()

    def presionar_enter_ahk(self, veces=1):
        """Presiona Enter usando AHK"""
        if not self.enter.start_ahk():
            logger.error("No se pudo iniciar AutoHotkey para Enter")
            return False
        
        success = self.enter.presionar_enter(veces)
        self.enter.stop_ahk()
        return success

    def click(self, x, y, duration=0.1):
        """Hacer clic en coordenadas específicas"""
        pyautogui.click(x, y, duration=duration)
        time.sleep(0.5)

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

    def detect_image_with_cv2(self, image_path, confidence=0.7):
        """Detectar imagen en pantalla usando template matching con OpenCV"""
        try:
            # Capturar pantalla completa
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # Cargar template de la imagen de referencia
            template = cv2.imread(image_path)
            if template is None:
                print(f"Error: No se pudo cargar la imagen {image_path}")
                return False, None
            
            # Realizar template matching
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            # Umbral de confianza
            if max_val < confidence:
                print(f"Imagen no encontrada. Mejor coincidencia: {max_val:.2f}")
                return False, None
            
            print(f"✅ Imagen encontrada con confianza: {max_val:.2f}")
            return True, max_loc
        except Exception as e:
            print(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.7):
        """Esperar a que aparezca una imagen con múltiples intentos usando OpenCV"""
        print(f"🔍 Buscando imagen: {image_path}")
        
        for attempt in range(1, max_attempts + 1):
            found, location = self.detect_image_with_cv2(image_path, confidence)
            
            if found:
                print(f"✅ Imagen detectada en el intento {attempt} en coordenadas: {location}")
                return True, location
            
            print(f"⏳ Intento {attempt}/{max_attempts} - Imagen no encontrada")
            
            if attempt < max_attempts:
                if attempt % 10 == 0:
                    print("⏰ Espera prolongada de 10 segundos...")
                    time.sleep(10)
                else:
                    time.sleep(2)
        
        print("❌ Imagen no encontrada después de 30 intentos. Terminando proceso.")
        return False, None

    def verificar_valores_csv(self, df, row_index):
        """Verifica si los valores necesarios del CSV existen y son válidos"""
        try:
            # Verificar que la fila existe
            if row_index >= len(df):
                print(f"❌ Fila {row_index} no existe en el CSV")
                return False
            
            row = df.iloc[row_index]
            #Verificar columna 27 (QTY Nom Neg) si es 1 hace el proceso si no lo salta 
            if len(row) <= 27 or pd.isna(row.iloc[27]) or row.iloc[27] != 1:
                print(f"⚠️  Columna 27 vacía, no es 1 o no existe en fila {row_index}, saltando...")
                #Ejecuta ultimos pasos para pasar a la siguiente linea pasos 12
                # 12. Presionar flecha abajo con AHK
                if not self.presionar_flecha_abajo_ahk(70, 266,1):
                    print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    print("✅ Flecha abajo presionada con AHK")
                self.sleep(2)
                return False
            
            # Verificar columna 28 (num_txt_type)
            if len(row) <= 28 or pd.isna(row.iloc[28]):
                print(f"⚠️  Columna 28 vacía o no existe en fila {row_index}, saltando...")
                #Ejecuta ultimos pasos para pasar a la siguiente linea pasos 12
                # 12. Presionar flecha abajo con AHK
                if not self.presionar_flecha_abajo_ahk(70, 266,1):
                    print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    print("✅ Flecha abajo presionada con AHK")
                self.sleep(2)
                return False
                return False
                
            # Verificar columna 29 (texto_adicional) - puede estar vacía pero debe existir
            if len(row) <= 29:
                print(f"⚠️  Columna 29 no existe en fila {row_index}, saltando...")
                #Ejecuta ultimos pasos para pasar a la siguiente linea pasos 12
                # 12. Presionar flecha abajo con AHK
                if not self.presionar_flecha_abajo_ahk(70, 266,1):
                    print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    print("✅ Flecha abajo presionada con AHK")
                self.sleep(2)
                return False
                return False
                
            return True
            
        except Exception as e:
            print(f"❌ Error verificando valores CSV: {e}")
            return False

    def perform_actions(self):
        """Función principal que realiza todas las acciones"""
        if not self.ahk_writer.start_ahk():
            print("❌ No se pudo iniciar AHKWriter")
            return False
            
        try:
            # Verificar si el archivo CSV existe
            if not os.path.exists(self.csv_file):
                print(f"❌ El archivo CSV no existe: {self.csv_file}")
                return False

            # Leer el archivo CSV
            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            if total_lines < 1:
                print("❌ No hay suficientes datos en el archivo CSV")
                return False

            print(f"📊 Total de líneas en CSV: {total_lines}")

            # Usar la línea específica determinada por el Programa 1
            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False

            row_index = self.linea_especifica - 1
                
            # Verificar si los valores del CSV son válidos
            if not self.verificar_valores_csv(df, row_index):
                print(f"⚠️  Valores inválidos en fila {row_index}. Línea {self.linea_especifica} saltada.")
                return True  # Consideramos éxito si no hay valores válidos
                    
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
            # Detener todos los manejadores AHK
            self.ahk_writer.stop_ahk()
            self.ahk_manager.stop_ahk()
            self.enter.stop_ahk()
            self.ahk_click_down.stop_ahk()

    def process_single_iteration(self, df, linea_especifica, total_lines):
        """Procesar una sola iteración del bucle"""
        # Obtener la fila correspondiente (0-indexed)
        row_index = linea_especifica - 1
        row = df.iloc[row_index]
        
        # Obtener valores del CSV con verificación
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

        # SECUENCIA DE ACCIONES
        try:
            # 1. Seleccionar Agregar ruta de GE
            self.click(*self.coords['agregar_ruta'])
            self.sleep(2)
            self.click(*self.coords['archivo'])
            self.sleep(2)
            self.click(*self.coords['abrir'])
            self.sleep(2) 
            
            # 2. Usar detección de ventana de archivo para cargar el archivo con AHK Manager
            nombre_archivo = self.nombre
            success = self.handle_archivo_special_behavior(nombre_archivo)
            
            if not success:
                print("❌ No se pudo cargar el archivo. Regresando a agregar_ruta...")
                self.click(*self.coords['agregar_ruta'])
                self.sleep(2)
                return False
            
            # 3. Presionar Enter con AHK para confirmar la carga del archivo
            if not self.presionar_enter_ahk(1):
                print("⚠️  No se pudo presionar Enter con AHK, usando pyautogui")
                pyautogui.press('enter')
            
            self.sleep(3)

            self.click(*self.coords['agregar_ruta'])
            self.sleep(2)

            self.click(1406, 675) #cargar ruta
            self.sleep(2)

            self.click(70, 266)#seleccionar lote
            self.sleep(2)
            
            # 4. Seleccionar en el mapa
            self.click(*self.coords['seleccionar_mapa'])
            self.sleep(2)
            
            # 5. Anotar
            self.click(*self.coords['anotar'])
            self.sleep(2)
            
            # 6. Agregar texto adicional
            self.click(*self.coords['agregar_texto_adicional'])
            self.sleep(2)
            
            # 7. DETECCIÓN DE IMAGEN para el campo de texto
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=10)
            
            if image_found:
                # Calcular coordenadas absolutas del campo de texto basado en la detección
                x_campo = base_location[0] + self.coords_texto_relativas['campo_texto'][0]
                y_campo = base_location[1] + self.coords_texto_relativas['campo_texto'][1]
                x_agregar = base_location[0] + self.coords_texto_relativas['agregar_texto'][0]
                y_agregar = base_location[1] + self.coords_texto_relativas['agregar_texto'][1]
                x_cerrar = base_location[0] + self.coords_texto_relativas['cerrar_ventana_texto'][0]
                y_cerrar = base_location[1] + self.coords_texto_relativas['cerrar_ventana_texto'][1]
                
                # Escribir el texto adicional usando AHK Writer SOLO si hay texto
                if texto_adicional and texto_adicional.strip():
                    writing_success = self.escribir_texto_adicional_ahk(x_campo, y_campo, texto_adicional)
                    if not writing_success:
                        print("⚠️  Falló la escritura con AHK, intentando con pyautogui...")
                        pyautogui.write(texto_adicional, interval=0.05)
                else:
                    print("ℹ️  Texto adicional vacío, no se escribe nada")
                
                self.sleep(2)

                # 8. Agregar de texto adicional
                self.click(x_agregar, y_agregar)
                self.sleep(3)
            
                # 9. Cerrar ventana de texto adicional
                self.click(x_cerrar, y_cerrar)
                self.sleep(2)
            else:
                print("❌ No se pudo detectar la imagen del campo de texto")
                return False
            
            # 10. Limpiar trazo
            self.click(*self.coords['limpiar_trazo'])
            self.sleep(1)
            
            # 11. Seleccionar Lote nuevamente
            self.click(*self.coords['lote_again'])
            self.sleep(2)
            
            # 12. Presionar flecha abajo con AHK
            if not self.presionar_flecha_abajo_ahk(*self.coords['lote_again'],1):
                print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                pyautogui.press('down')
            else:
                print("✅ Flecha abajo presionada con AHK")
            
            self.sleep(2)
            
            # 13. Detectar ventana de error después de cada iteración
            if self.detectar_ventana_error():
                print("✅ Ventana de error detectada y cerrada")
            
            #14. Cerrar_ventana_archivo
            self.click(*self.coords['cerrar_ventana_archivo'])
            self.sleep(2)

            print(f"✅ Línea {linea_especifica} completada exitosamente")

            return True
            
        except Exception as e:
            print(f"❌ Error en línea {linea_especifica}: {e}")
            # Intentar cerrar ventana de error en caso de excepción
            self.detectar_ventana_error()
            return False

    def save_progress(self):
        """Guardar progreso con Ctrl + S"""
        print("💾 Guardando progreso...")
        pyautogui.hotkey('ctrl', 's')
        self.sleep(6)

def main():
    """Función principal para ejecutar la interfaz gráfica"""
    root = tk.Tk()
    app = InterfazAutomation(root)
    root.mainloop()

if __name__ == "__main__":
    main()