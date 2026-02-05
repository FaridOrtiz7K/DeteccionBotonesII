import threading
import time
import pandas as pd
import keyboard
from tkinter import filedialog, messagebox, simpledialog
import pyautogui

from model.estado import EstadoEjecucion
from model.datos_globales import DatosGlobales
from model.procesador_csv import ProcesadorCSV
from model.nse_automation import NSEAutomation
from model.nse_services import NSEServicesAutomation
from model.ge_automation import GEAutomation
from view.ventana_pausa import PauseWindow
from utils.ahk_manager_save import AHKSaveManager

class ControladorPrincipal:
    def __init__(self, vista, modelo_estado, datos_globales):
        self.vista = vista
        self.estado_global = modelo_estado
        self.datos_globales = datos_globales
        
        # Manager de guardado AHK
        self.save_manager = AHKSaveManager()
        
        # Referencia a ventana de pausa
        self.pause_window = None
        
        # Configurar bindings
        self.setup_bindings()
    
    def setup_bindings(self):
        keyboard.add_hotkey('esc', self.pausar_proceso)
        keyboard.add_hotkey('f2', self.pausar_proceso)
        keyboard.add_hotkey('f3', self.reanudar_proceso)
        keyboard.add_hotkey('f4', self.detener_proceso)
    
    def seleccionar_csv(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if archivo:
            self.datos_globales.CSV_FILE = archivo
            self.vista.csv_file.set(archivo)
            self.vista.log(f"CSV seleccionado: {archivo}")
            
            try:
                df = pd.read_csv(archivo)
                self.vista.linea_maxima.set(len(df))
                self.vista.lote_fin.set(len(df))
                self.vista.actualizar_estado_lineas()
                self.vista.log(f"CSV cargado: {len(df)} registros encontrados")
            except Exception as e:
                self.vista.log(f"Error al leer CSV: {e}")
    
    def consultar_id(self):
        if not self.datos_globales.CSV_FILE:
            messagebox.showerror("Error", "Primero seleccione un archivo CSV")
            return
            
        id_buscar = self.vista.id_consulta.get()
        if not id_buscar:
            messagebox.showwarning("Advertencia", "Ingrese un ID para consultar")
            return
            
        try:
            df = pd.read_csv(self.datos_globales.CSV_FILE)
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
                self.vista.log(f"ID {id_buscar} consultado: {len(resultado)} coincidencias")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al consultar ID: {e}")
    
    def actualizar_info_lote(self, linea_actual, datos):
        self.datos_globales.INFO_LOTE_ACTUAL = {
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
            
            self.vista.info_lote.set(info_text)
        else:
            self.vista.info_lote.set(f"Línea {linea_actual}: Datos vacíos o no encontrados")
    
    def escribir_prueba_a(self):
        if not self.datos_globales.CSV_FILE:
            messagebox.showerror("Error", "Primero seleccione un archivo CSV")
            return
            
        # ... (resto del método)
    
    def configurar_kml(self):
        nuevo_nombre = simpledialog.askstring(
            "Configurar KML", 
            "Ingrese el nuevo nombre para archivos KML:",
            initialvalue=self.datos_globales.KML_FILENAME
        )
        if nuevo_nombre:
            self.datos_globales.KML_FILENAME = nuevo_nombre
            self.vista.log(f"Nombre KML configurado a: {self.datos_globales.KML_FILENAME}")
    
    def iniciar_proceso(self):
        self.datos_globales.CSV_FILE = self.vista.csv_file.get()
        self.datos_globales.LOTE_INICIO = self.vista.lote_inicio.get()
        self.datos_globales.LOTE_FIN = self.vista.lote_fin.get()
        
        if self.datos_globales.LOTE_INICIO > self.datos_globales.LOTE_FIN:
            messagebox.showerror("Error", "El lote inicio no puede ser mayor que el lote fin")
            return
            
        try:
            df = pd.read_csv(self.datos_globales.CSV_FILE)
            if len(df) == 0:
                self.vista.log("⚠️ CSV está vacío. Continuando sin procesar datos...")
                self.datos_globales.LINEA_MAXIMA = 0
                self.datos_globales.LINEA_ACTUAL = 0
            else:
                self.datos_globales.LINEA_MAXIMA = min(self.datos_globales.LOTE_FIN, len(df))
                self.datos_globales.LINEA_ACTUAL = max(1, self.datos_globales.LOTE_INICIO)
                
                if self.datos_globales.LINEA_ACTUAL > self.datos_globales.LINEA_MAXIMA:
                    messagebox.showerror("Error", "El lote inicio está fuera del rango disponible")
                    return
        except Exception as e:
            self.vista.log(f"❌ Error al leer CSV: {e}")
            return
        
        # Iniciar manager de guardado AHK
        if not self.save_manager.start_ahk():
            self.vista.log("⚠️ No se pudo iniciar AutoHotkey para guardar. Continuando sin funcionalidad de guardado...")
        else:
            self.vista.log("✅ AutoHotkey para guardar iniciado correctamente")
        
        self.estado_global.set_ejecutando(True)
        self.estado_global.set_pausado(False)
        self.estado_global.set_detener_inmediato(False)
        self.estado_global.set_linea_en_proceso(False)
        
        self.vista.actualizar_estado_botones(self.estado_global)
        self.vista.estado.set("Ejecutando...")
        self.vista.estado_label.configure(foreground="green")
        
        self.hilo_ejecucion = threading.Thread(target=self.ejecutar_procesos)
        self.hilo_ejecucion.daemon = True
        self.hilo_ejecucion.start()
    
    def ejecutar_procesos(self):
        lotes_desde_ultimo_guardado = 0
        
        try:
            while (self.datos_globales.LINEA_ACTUAL <= self.datos_globales.LINEA_MAXIMA and 
                   self.estado_global.verificar_continuar()):
                
                if self.estado_global.esperar_si_pausado():
                    break
                
                self.estado_global.set_linea_en_proceso(True)
                    
                self.vista.log(f"🔄 Procesando línea {self.datos_globales.LINEA_ACTUAL}/{self.datos_globales.LINEA_MAXIMA}")
                self.vista.actualizar_estado_lineas()
                
                # Ejecutar Programa 1
                resultado1, linea_procesada, datos_lote = self.ejecutar_programa1_interfaz(
                    self.datos_globales.LINEA_ACTUAL
                )
                
                if not self.estado_global.verificar_continuar():
                    break
                    
                self.actualizar_info_lote(self.datos_globales.LINEA_ACTUAL, datos_lote)
                
                if not resultado1 or linea_procesada is None:
                    self.vista.log(f"⚠️ Programa 1 no procesó la línea {self.datos_globales.LINEA_ACTUAL}. Saltando al siguiente lote...")
                    self.estado_global.set_linea_en_proceso(False)
                    self.datos_globales.LINEA_ACTUAL += 1
                    lotes_desde_ultimo_guardado += 1
                    continue
                
                # ... (resto del proceso)
                
        except Exception as e:
            self.vista.log(f"❌ Error en ejecución: {e}")
            self.vista.estado.set("Error")
            self.vista.estado_label.configure(foreground="red")
            self.estado_global.set_linea_en_proceso(False)
        
        finally:
            self.estado_global.set_ejecutando(False)
            self.save_manager.stop_ahk()
            self.vista.actualizar_estado_botones(self.estado_global)
            if self.pause_window is not None and self.pause_window.winfo_exists():
                self.pause_window.destroy()
                self.pause_window = None
    
    def ejecutar_programa1_interfaz(self, linea_especifica):
        try:
            self.vista.log("=" * 60)
            self.vista.log("INICIANDO PROGRAMA 1 - PROCESADOR CSV")
            self.vista.log("=" * 60)
            
            procesador = ProcesadorCSV(self.datos_globales.CSV_FILE, self.estado_global)
            
            self.vista.log("Iniciando procesamiento automático del Programa 1...")
            
            if self.estado_global.esperar_si_pausado():
                return False, linea_especifica, None
            time.sleep(0.5)
            
            if not procesador.cargar_csv():
                self.vista.log("⚠️ CSV vacío detectado. Continuando sin procesar...")
                return True, None, None
            
            if linea_especifica > len(procesador.df):
                self.vista.log(f"⚠️ Línea {linea_especifica} no existe en CSV. Saltando...")
                return True, None, None
            
            procesador.df = procesador.df.iloc[linea_especifica-1:linea_especifica]
            datos_lote = procesador.df.iloc[0] if len(procesador.df) > 0 else None
            
            resultado, linea_procesada = procesador.procesar_todo()
            
            if resultado and linea_procesada:
                self.vista.log(f"✅ Programa 1 completado. Línea procesada: {linea_procesada}")
                return True, linea_procesada, datos_lote
            else:
                self.vista.log("❌ Programa 1 falló o no encontró ID")
                return False, None, datos_lote
                
        except Exception as e:
            self.vista.log(f"❌ Error en Programa 1: {e}")
            return False, None, None
    
    # ... (métodos para pausar, reanudar, detener, etc.)