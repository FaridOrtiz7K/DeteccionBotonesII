import os
import threading
import time
import pandas as pd
import keyboard
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pyautogui
import logging

from models.estado import EstadoEjecucion
from models.datos_globales import DatosGlobales
from models.procesador_csv import ProcesadorCSV
from models.nse_automation import NSEAutomation
from models.nse_services import NSEServicesAutomation
from models.ge_automation import GEAutomation
from views.ventana_pausa import PauseWindow
from utils.ahk_manager_save import AHKSaveManager
from utils.ahk_writer import AHKWriter

logger = logging.getLogger(__name__)

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
        
        # Configurar comandos de la vista
        self.configurar_comandos_vista()
    
    def configurar_comandos_vista(self):
        """Configura los comandos de los botones de la vista"""
        # Botones de control
        self.vista.btn_iniciar.config(command=self.iniciar_proceso)
        self.vista.btn_pausar.config(command=self.pausar_proceso)
        self.vista.btn_reanudar.config(command=self.reanudar_proceso)
        self.vista.btn_detener.config(command=self.detener_proceso)
        
        # Botones adicionales
        self.vista.btn_prueba.config(command=self.escribir_prueba_a)
        self.vista.btn_kml.config(command=self.configurar_kml)
        
        # Reemplazar métodos de la vista
        self.vista.seleccionar_csv = self.seleccionar_csv
        self.vista.consultar_id = self.consultar_id
        self.vista.escribir_prueba_a = self.escribir_prueba_a
        self.vista.configurar_kml = self.configurar_kml
        self.vista.actualizar_estado_lineas = self.actualizar_estado_lineas
        self.vista.actualizar_estado_botones = self.actualizar_estado_botones_vista
    
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
                self.actualizar_estado_lineas()
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
            
        try:
            df = pd.read_csv(self.datos_globales.CSV_FILE)
            if len(df) == 0:
                messagebox.showerror("Error", "El CSV está vacío")
                return
                
            ultima_columna = df.columns[-1]
            if pd.isna(ultima_columna):
                texto_a_escribir = ""
            else:
                texto_a_escribir = str(ultima_columna)
            
            self.vista.log(f"Escribiendo: '{texto_a_escribir}'")
            self.vista.log("⚠️ Coloque el cursor en la posición deseada - escribiendo en 5 segundos...")
            
            for i in range(5, 0, -1):
                self.vista.estado.set(f"Escribiendo en {i} segundos...")
                time.sleep(1)
            
            x, y = pyautogui.position()
            self.vista.log(f"Posición del cursor: ({x}, {y})")
            
            ahk_writer = AHKWriter()
            if ahk_writer.start_ahk():
                exito = ahk_writer.ejecutar_escritura_ahk(x, y, texto_a_escribir)
                ahk_writer.stop_ahk()
                
                if exito:
                    self.vista.log("✅ Texto escrito exitosamente")
                    self.vista.estado.set("Listo")
                else:
                    self.vista.log("❌ Error al escribir texto")
                    self.vista.estado.set("Error")
            else:
                self.vista.log("❌ No se pudo iniciar AHKWriter")
                self.vista.estado.set("Error")
                
        except Exception as e:
            self.vista.log(f"❌ Error al escribir PRUEBA A: {e}")
            self.vista.estado.set("Error")
    
    def configurar_kml(self):
        nuevo_nombre = simpledialog.askstring(
            "Configurar KML", 
            "Ingrese el nuevo nombre para archivos KML:",
            initialvalue=self.datos_globales.KML_FILENAME
        )
        if nuevo_nombre:
            self.datos_globales.KML_FILENAME = nuevo_nombre
            self.vista.log(f"Nombre KML configurado a: {self.datos_globales.KML_FILENAME}")
    
    def actualizar_estado_lineas(self):
        self.vista.linea_actual.set(str(self.datos_globales.LINEA_ACTUAL))
        lineas_rest = max(0, self.datos_globales.LINEA_MAXIMA - self.datos_globales.LINEA_ACTUAL)
        self.vista.lineas_restantes.set(str(lineas_rest))
    
    def actualizar_estado_botones_vista(self):
        estado_actual = self.vista.estado.get()
        if "Reanudando en" in estado_actual:
            self.vista.btn_iniciar.config(state=tk.DISABLED)
            self.vista.btn_pausar.config(state=tk.DISABLED)
            self.vista.btn_reanudar.config(state=tk.DISABLED)
            self.vista.btn_detener.config(state=tk.NORMAL)
            return
        
        if self.estado_global.ejecutando:
            self.vista.btn_iniciar.config(state=tk.DISABLED)
            self.vista.btn_detener.config(state=tk.NORMAL)
            
            if self.estado_global.pausado:
                self.vista.btn_pausar.config(state=tk.DISABLED)
                self.vista.btn_reanudar.config(state=tk.NORMAL)
            else:
                self.vista.btn_pausar.config(state=tk.NORMAL)
                self.vista.btn_reanudar.config(state=tk.DISABLED)
        else:
            self.vista.btn_iniciar.config(state=tk.NORMAL)
            self.vista.btn_pausar.config(state=tk.DISABLED)
            self.vista.btn_reanudar.config(state=tk.DISABLED)
            self.vista.btn_detener.config(state=tk.DISABLED)
    
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
        
        self.actualizar_estado_botones_vista()
        self.vista.estado.set("Ejecutando...")
        self.vista.estado_label.configure(foreground="green")
        
        self.hilo_ejecucion = threading.Thread(target=self.ejecutar_procesos)
        self.hilo_ejecucion.daemon = True
        self.hilo_ejecucion.start()
    
    def pausar_proceso(self):
        if self.estado_global.ejecutando and not self.estado_global.pausado:
            if self.estado_global.en_cuenta_regresiva:
                self.estado_global.set_en_cuenta_regresiva(False)
                self.vista.log("⏸️ Cuenta regresiva cancelada - Proceso permanece pausado")
            else:
                self.estado_global.set_pausado(True)
                self.vista.estado.set("Pausado")
                self.vista.estado_label.configure(foreground="orange")
                self.vista.log("⏸️ Proceso pausado")
                
                # Mostrar ventana emergente de pausa
                if self.pause_window is None or not self.pause_window.winfo_exists():
                    self.pause_window = PauseWindow(
                        self.vista.root, 
                        self, 
                        self.datos_globales.LINEA_ACTUAL,
                        self.datos_globales.LINEA_MAXIMA
                    )
                
                self.actualizar_estado_botones_vista()
    
    def reanudar_desde_ventana(self):
        """Método para reanudar desde la ventana de pausa"""
        self.estado_global.set_pausado(False)
        self.vista.estado.set("Ejecutando...")
        self.vista.estado_label.configure(foreground="green")
        self.vista.log("▶️ Proceso reanudado desde ventana de pausa")
        self.actualizar_estado_botones_vista()
    
    def reanudar_proceso(self):
        if self.estado_global.ejecutando and self.estado_global.pausado:
            # Si hay ventana de pausa abierta, cerrarla
            if self.pause_window is not None and self.pause_window.winfo_exists():
                self.pause_window.destroy()
                self.pause_window = None
            
            self.vista.btn_reanudar.config(state=tk.DISABLED)
            self.estado_global.set_en_cuenta_regresiva(True)
            
            self.vista.estado.set("Reanudando en 3 segundos...")
            self.vista.estado_label.configure(foreground="blue")
            self.vista.log("Reanudando en 3 segundos...")
            
            threading.Thread(target=self._cuenta_regresiva_reanudacion, daemon=True).start()

    def _cuenta_regresiva_reanudacion(self):
        try:
            for i in range(3, 0, -1):
                if self.estado_global.detener_inmediato:
                    self.vista.root.after(0, self._cancelar_reanudacion, "Reanudación cancelada (proceso detenido)")
                    return
                    
                self.vista.root.after(0, lambda x=i: self.vista.estado.set(f"Reanudando en {x} segundos..."))
                self.vista.root.after(0, self.vista.root.update)
                time.sleep(1)
            
            if not self.estado_global.detener_inmediato:
                self.vista.root.after(0, self._completar_reanudacion)
            else:
                self.vista.root.after(0, self._cancelar_reanudacion, 
                            "Reanudación cancelada (proceso detenido durante la cuenta regresiva)")
        except Exception as e:
            self.vista.root.after(0, self._cancelar_reanudacion, f"Error en cuenta regresiva: {e}")

    def _completar_reanudacion(self):
        self.estado_global.set_en_cuenta_regresiva(False)
        self.estado_global.set_pausado(False)
        self.vista.estado.set("Ejecutando...")
        self.vista.estado_label.configure(foreground="green")
        self.vista.log("▶️ Proceso reanudado")
        self.actualizar_estado_botones_vista()

    def _cancelar_reanudacion(self, mensaje):
        self.estado_global.set_en_cuenta_regresiva(False)
        self.vista.log(mensaje)
        self.actualizar_estado_botones_vista()
        
    def detener_proceso(self):
        # Cerrar ventana de pausa si está abierta
        if self.pause_window is not None and self.pause_window.winfo_exists():
            self.pause_window.destroy()
            self.pause_window = None
        
        self.estado_global.set_detener_inmediato(True)
        self.estado_global.set_ejecutando(False)
        
        # Detener manager de guardado AHK
        self.save_manager.stop_ahk()
        
        self.vista.estado.set("Detenido")
        self.vista.estado_label.configure(foreground="red")
        self.vista.log("⏹️ Proceso detenido inmediatamente")
        self.actualizar_estado_botones_vista()
        self.actualizar_estado_lineas()
    
    def ejecutar_procesos(self):
        lotes_desde_ultimo_guardado = 0
        
        try:
            while (self.datos_globales.LINEA_ACTUAL <= self.datos_globales.LINEA_MAXIMA and 
                   self.estado_global.verificar_continuar()):
                
                if self.estado_global.esperar_si_pausado():
                    break
                
                self.estado_global.set_linea_en_proceso(True)
                    
                self.vista.log(f"🔄 Procesando línea {self.datos_globales.LINEA_ACTUAL}/{self.datos_globales.LINEA_MAXIMA}")
                self.actualizar_estado_lineas()
                
                # Ejecutar Programa 1
                resultado1, linea_procesada, datos_lote = self.ejecutar_programa1_interfaz(
                    self.datos_globales.LINEA_ACTUAL
                )
                
                if not self.estado_global.verificar_continuar():
                    break
                    
                self.actualizar_info_lote(self.datos_globales.LINEA_ACTUAL, datos_lote)
                
                if not resultado1 or linea_procesada is None:
                    self.vista.log(f"⚠️ Programa 1 no procesó la línea {self.datos_globales.LINEA_ACTUAL} (ID no encontrado después de 2 intentos). Saltando al siguiente lote...")
                    self.estado_global.set_linea_en_proceso(False)
                    self.datos_globales.LINEA_ACTUAL += 1
                    lotes_desde_ultimo_guardado += 1
                    continue
                
                for _ in range(3):
                    if self.estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)
                
                if not self.estado_global.verificar_continuar():
                    break
                    
                # Ejecutar Programa 2
                self.vista.log("Iniciando Programa 2 - Automatización NSE")
                resultado2 = self.ejecutar_programa2_interfaz(linea_procesada)
                
                if not resultado2:
                    self.vista.log(f"❌ Programa 2 falló en línea {self.datos_globales.LINEA_ACTUAL}")
                    self.estado_global.set_linea_en_proceso(False)
                    self.datos_globales.LINEA_ACTUAL += 1
                    lotes_desde_ultimo_guardado += 1
                    continue
                
                for _ in range(3):
                    if self.estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)
                
                if not self.estado_global.verificar_continuar():
                    break
                    
                # Ejecutar Programa 3
                self.vista.log("Iniciando Programa 3 - Servicios NSE")
                resultado3 = self.ejecutar_programa3_interfaz(linea_procesada)
                
                if not resultado3:
                    self.vista.log(f"❌ Programa 3 falló en línea {self.datos_globales.LINEA_ACTUAL}")
                    self.estado_global.set_linea_en_proceso(False)
                    self.datos_globales.LINEA_ACTUAL += 1
                    lotes_desde_ultimo_guardado += 1
                    continue
                
                for _ in range(3):
                    if self.estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)
                
                if not self.estado_global.verificar_continuar():
                    break
                    
                # Ejecutar Programa 4
                self.vista.log("Iniciando Programa 4 - Automatización GE")
                resultado4 = self.ejecutar_programa4_interfaz(linea_procesada)
                
                if resultado4:
                    self.vista.log(f"✅ Línea {self.datos_globales.LINEA_ACTUAL} procesada exitosamente")
                else:
                    self.vista.log(f"⚠️ Línea {self.datos_globales.LINEA_ACTUAL} completada con advertencias")
                
                self.estado_global.set_linea_en_proceso(False)
                lotes_desde_ultimo_guardado += 1
                
                # GUARDAR CADA 10 LOTES CON AHK
                if lotes_desde_ultimo_guardado >= 10:
                    self.vista.log("📁 Guardando progreso después de 10 lotes...")
                    if self.save_manager.is_running:
                        if not self.save_manager.trigger_save():
                            self.vista.log("⚠️ No se pudo guardar con AHK. Usando pyautogui como respaldo...")
                            pyautogui.hotkey('ctrl', 's')
                    else:
                        self.vista.log("⚠️ AHK no está corriendo. Usando pyautogui...")
                        pyautogui.hotkey('ctrl', 's')
                    
                    for _ in range(6):
                        if self.estado_global.esperar_si_pausado():
                            break
                        time.sleep(1)
                    self.vista.log("✅ Progreso guardado exitosamente")
                    lotes_desde_ultimo_guardado = 0
                
                self.datos_globales.LINEA_ACTUAL += 1
                self.actualizar_estado_lineas()
                
                for _ in range(4):
                    if self.estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)
            
            # GUARDADO FINAL CON AHK
            if not self.estado_global.detener_inmediato and self.datos_globales.LINEA_ACTUAL > self.datos_globales.LINEA_MAXIMA:
                self.vista.log("📁 Guardando progreso final al completar todos los lotes...")
                if self.save_manager.is_running:
                    if not self.save_manager.trigger_save():
                        self.vista.log("⚠️ No se pudo guardar con AHK. Usando pyautogui como respaldo...")
                        pyautogui.hotkey('ctrl', 's')
                else:
                    self.vista.log("⚠️ AHK no está corriendo. Usando pyautogui...")
                    pyautogui.hotkey('ctrl', 's')
                
                time.sleep(6)
                self.vista.log("✅ Progreso final guardado exitosamente")
            
            if self.estado_global.detener_inmediato:
                self.vista.log("🛑 Proceso detenido inmediatamente por usuario")
                self.vista.estado.set("Detenido")
                self.vista.estado_label.configure(foreground="red")
            elif self.estado_global.ejecutando and self.datos_globales.LINEA_ACTUAL > self.datos_globales.LINEA_MAXIMA:
                self.vista.log("🎉 Proceso completado exitosamente")
                self.vista.estado.set("Completado")
                self.vista.estado_label.configure(foreground="green")
            elif not self.estado_global.ejecutando:
                self.vista.log("Proceso detenido por el usuario")
                
        except Exception as e:
            self.vista.log(f"❌ Error en ejecución: {e}")
            self.vista.estado.set("Error")
            self.vista.estado_label.configure(foreground="red")
            self.estado_global.set_linea_en_proceso(False)
        
        finally:
            self.estado_global.set_ejecutando(False)
            self.save_manager.stop_ahk()  # Detener manager de guardado
            self.actualizar_estado_botones_vista()
            # Cerrar ventana de pausa si está abierta
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
    
    def ejecutar_programa2_interfaz(self, linea_especifica):
        try:
            self.vista.log("\n" + "=" * 60)
            self.vista.log("INICIANDO PROGRAMA 2 - AUTOMATIZACIÓN NSE")
            self.vista.log("=" * 60)
            
            nse = NSEAutomation(
                linea_especifica=linea_especifica,
                csv_file=self.datos_globales.CSV_FILE,
                estado_global=self.estado_global
            )
            nse.is_running = True
            
            if not os.path.exists(nse.csv_file):
                self.vista.log(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
                return True
            
            try:
                df = pd.read_csv(nse.csv_file)
                if len(df) == 0:
                    self.vista.log("⚠️ CSV vacío. Saltando Programa 2...")
                    return True
            except:
                pass
            
            self.vista.log(f"🎯 Procesando línea: {linea_especifica}")
            
            for _ in range(1):
                if self.estado_global.esperar_si_pausado():
                    return False
                time.sleep(0.5)
            
            resultado = nse.execute_nse_script()
            
            if resultado:
                self.vista.log("✅ Programa 2 finalizado exitosamente")
            else:
                self.vista.log("⚠️ Programa 2 completado con advertencias")
                
            return resultado
            
        except Exception as e:
            self.vista.log(f"❌ Error en Programa 2: {e}")
            return False
    
    def ejecutar_programa3_interfaz(self, linea_especifica):
        try:
            self.vista.log("\n" + "=" * 60)
            self.vista.log("INICIANDO PROGRAMA 3 - SERVICIOS NSE")
            self.vista.log("=" * 60)
            
            nse_services = NSEServicesAutomation(
                linea_especifica=linea_especifica,
                csv_file=self.datos_globales.CSV_FILE,
                estado_global=self.estado_global
            )
            
            if not os.path.exists(nse_services.csv_file):
                self.vista.log(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
                return True
            
            self.vista.log(f"🎯 Procesando línea: {linea_especifica}")
            
            if not nse_services.iniciar_ahk():
                self.vista.log("⚠️ No se pudieron iniciar los servicios AHK. Continuando...")
                return True
            
            nse_services.is_running = True
            resultado = nse_services.procesar_linea_especifica()
            
            if resultado:
                self.vista.log(f"✅ Programa 3 completado exitosamente")
            else:
                self.vista.log(f"⚠️ Programa 3 completado con advertencias")
            
            nse_services.detener_ahk()
            return resultado
            
        except Exception as e:
            self.vista.log(f"❌ Error en Programa 3: {e}")
            return False
    
    def ejecutar_programa4_interfaz(self, linea_especifica):
        try:
            self.vista.log("\n" + "=" * 60)
            self.vista.log("INICIANDO PROGRAMA 4 - AUTOMATIZACIÓN GE")
            self.vista.log("=" * 60)
            
            ge_auto = GEAutomation(
                linea_especifica=linea_especifica,
                csv_file=self.datos_globales.CSV_FILE,
                kml_filename=self.datos_globales.KML_FILENAME,
                estado_global=self.estado_global
            )
            ge_auto.is_running = True
            
            if not os.path.exists(ge_auto.csv_file):
                self.vista.log(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
                return True
            
            try:
                df = pd.read_csv(ge_auto.csv_file)
                if len(df) == 0:
                    self.vista.log("⚠️ CSV vacío. Saltando Programa 4...")
                    return True
            except:
                pass
            
            self.vista.log(f"🎯 Procesando línea: {linea_especifica}")
            self.vista.log(f"📁 Archivo KML: {ge_auto.nombre}")
            
            for _ in range(1):
                if self.estado_global.esperar_si_pausado():
                    return False
                time.sleep(0.5)
            
            success = ge_auto.perform_actions()
            
            if success:
                self.vista.log("✅ Programa 4 finalizado exitosamente")
            else:
                self.vista.log("⚠️ Programa 4 completado con advertencias")
                
            return success
            
        except Exception as e:
            self.vista.log(f"❌ Error en Programa 4: {e}")
            return False