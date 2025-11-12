# controller.py
import threading
import time
import pandas as pd
import keyboard
from models.modelo import EstadoPrograma, ProcesadorCSV, NSEAutomation, NSEServicesAutomation, GEAutomation
from utils.ahk_writer import AHKWriter

class ControladorAutomation:
    def __init__(self):
        self.modelo = EstadoPrograma()
        self.vista = None
        self.hilo_ejecucion = None
        
    def set_vista(self, vista):
        self.vista = vista
        self.setup_bindings()
    
    def setup_bindings(self):
        # Configurar teclas globales
        keyboard.add_hotkey('esc', self.mostrar_estado_actual)
        keyboard.add_hotkey('f2', self.pausar_proceso)
        keyboard.add_hotkey('f3', self.reanudar_proceso)
        keyboard.add_hotkey('f4', self.detener_proceso)
    
    def actualizar_vista(self):
        """Actualizar todos los elementos de la vista"""
        if self.vista:
            self.vista.actualizar_estado_botones(self.modelo)
            self.vista.actualizar_estado_lineas(self.modelo)
            self.vista.actualizar_estado_general(self.modelo)
    
    def log(self, mensaje):
        """Agregar mensaje al log"""
        if self.vista:
            self.vista.log(mensaje)
    
    def seleccionar_csv(self):
        """Seleccionar archivo CSV"""
        archivo = self.vista.pedir_seleccion_csv()
        if archivo:
            self.modelo.csv_file = archivo
            self.vista.csv_file.set(archivo)
            self.log(f"CSV seleccionado: {archivo}")
            
            # Calcular número máximo de líneas
            try:
                df = pd.read_csv(archivo)
                self.modelo.linea_maxima = len(df)
                self.vista.linea_maxima.set(len(df))
                self.actualizar_vista()
            except Exception as e:
                self.log(f"Error al leer CSV: {e}")
    
    def escribir_prueba_a(self):
        """Escribir PRUEBA A desde la última columna de la primera fila"""
        if not self.modelo.csv_file:
            self.vista.mostrar_mensaje("Error", "Primero seleccione un archivo CSV", "error")
            return
            
        try:
            df = pd.read_csv(self.modelo.csv_file)
            if len(df) == 0:
                self.vista.mostrar_mensaje("Error", "El CSV está vacío", "error")
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
        nuevo_nombre = self.vista.pedir_configuracion_kml(self.modelo.kml_filename)
        if nuevo_nombre:
            self.modelo.kml_filename = nuevo_nombre
            self.log(f"Nombre KML configurado a: {self.modelo.kml_filename}")
    
    def mostrar_estado_actual(self):
        """Mostrar estado actual al presionar ESC"""
        if self.modelo.ejecutando:
            lineas_restantes = self.modelo.linea_maxima - self.modelo.linea_actual
            mensaje = f"Línea actual: {self.modelo.linea_actual}\nLíneas restantes: {lineas_restantes}"
            self.vista.mostrar_mensaje("Estado Actual", mensaje)
    
    def iniciar_proceso(self):
        """Iniciar el proceso completo"""
        if not self.modelo.csv_file:
            self.vista.mostrar_mensaje("Error", "Seleccione un archivo CSV primero", "error")
            return
            
        self.modelo.linea_maxima = self.vista.linea_maxima.get()
        self.modelo.linea_actual = 1
        
        if self.modelo.linea_actual > self.modelo.linea_maxima:
            self.vista.mostrar_mensaje("Error", "La línea actual no puede ser mayor que la línea máxima", "error")
            return
            
        self.modelo.ejecutando = True
        self.modelo.pausado = False
        self.modelo.estado = "Ejecutando..."
        
        self.actualizar_vista()
        
        # Iniciar en hilo separado
        self.hilo_ejecucion = threading.Thread(target=self.ejecutar_procesos)
        self.hilo_ejecucion.daemon = True
        self.hilo_ejecucion.start()
    
    def pausar_proceso(self):
        """Pausar el proceso"""
        if self.modelo.ejecutando and not self.modelo.pausado:
            self.modelo.pausado = True
            self.modelo.estado = "Pausado"
            self.log("⏸️ Proceso pausado")
            self.actualizar_vista()
    
    def reanudar_proceso(self):
        """Reanudar el proceso después de 5 segundos"""
        if self.modelo.ejecutando and self.modelo.pausado:
            self.modelo.estado = "Reanudando en 5 segundos..."
            self.actualizar_vista()
            
            for i in range(5, 0, -1):
                self.modelo.estado = f"Reanudando en {i} segundos..."
                self.actualizar_vista()
                time.sleep(1)
                
            self.modelo.pausado = False
            self.modelo.estado = "Ejecutando..."
            self.log("▶️ Proceso reanudado")
            self.actualizar_vista()
    
    def detener_proceso(self):
        """Detener completamente el proceso"""
        self.modelo.ejecutando = False
        self.modelo.pausado = False
        self.modelo.linea_actual = 0
        self.modelo.estado = "Detenido"
        
        self.log("⏹️ Proceso detenido")
        self.actualizar_vista()
    
    def ejecutar_procesos(self):
        """Ejecutar los procesos secuencialmente para cada línea"""
        try:
            while self.modelo.linea_actual <= self.modelo.linea_maxima and self.modelo.ejecutando:
                # Verificar pausa
                while self.modelo.pausado and self.modelo.ejecutando:
                    time.sleep(0.1)
                    
                if not self.modelo.ejecutando:
                    break
                    
                self.log(f"🔄 Procesando línea {self.modelo.linea_actual}/{self.modelo.linea_maxima}")
                self.actualizar_vista()
                
                # Ejecutar Programa 1
                self.log("Iniciando Programa 1 - Procesador CSV")
                resultado1, linea_procesada = self.ejecutar_programa1(self.modelo.linea_actual)
                
                if not resultado1 or not self.modelo.ejecutando:
                    if not self.modelo.ejecutando:
                        break
                    self.log(f"❌ Programa 1 falló en línea {self.modelo.linea_actual}")
                    self.modelo.linea_actual += 1
                    continue
                
                # Ejecutar Programa 2
                self.log("Iniciando Programa 2 - Automatización NSE")
                resultado2 = self.ejecutar_programa2(linea_procesada)
                
                if not resultado2 or not self.modelo.ejecutando:
                    if not self.modelo.ejecutando:
                        break
                    self.log(f"❌ Programa 2 falló en línea {self.modelo.linea_actual}")
                    self.modelo.linea_actual += 1
                    continue
                
                # Ejecutar Programa 3
                self.log("Iniciando Programa 3 - Servicios NSE")
                resultado3 = self.ejecutar_programa3(linea_procesada)
                
                if not resultado3 or not self.modelo.ejecutando:
                    if not self.modelo.ejecutando:
                        break
                    self.log(f"❌ Programa 3 falló en línea {self.modelo.linea_actual}")
                    self.modelo.linea_actual += 1
                    continue
                
                # Ejecutar Programa 4
                self.log("Iniciando Programa 4 - Automatización GE")
                resultado4 = self.ejecutar_programa4(linea_procesada, self.modelo.kml_filename)
                
                if resultado4:
                    self.log(f"✅ Línea {self.modelo.linea_actual} procesada exitosamente")
                else:
                    self.log(f"⚠️ Línea {self.modelo.linea_actual} completada con advertencias")
                
                self.modelo.linea_actual += 1
                self.actualizar_vista()
                
                # Pequeña pausa entre líneas
                time.sleep(2)
            
            if self.modelo.ejecutando and self.modelo.linea_actual > self.modelo.linea_maxima:
                self.log("🎉 Proceso completado exitosamente")
                self.modelo.estado = "Completado"
                self.modelo.ejecutando = False
            elif not self.modelo.ejecutando:
                self.log("Proceso detenido por el usuario")
                
        except Exception as e:
            self.log(f"❌ Error en ejecución: {e}")
            self.modelo.estado = "Error"
        
        finally:
            self.modelo.ejecutando = False
            self.modelo.pausado = False
            self.actualizar_vista()
    
    def ejecutar_programa1(self, linea_especifica):
        """Ejecutar Programa 1 - Procesador CSV"""
        try:
            procesador = ProcesadorCSV(self.modelo.csv_file)
            
            # Cargar CSV y configurar línea específica
            if not procesador.cargar_csv():
                return False, None
                
            # Modificar para usar línea específica
            procesador.df = procesador.df.iloc[linea_especifica-1:linea_especifica]
            
            resultado, linea_procesada = procesador.procesar_todo()
            
            if resultado and linea_procesada:
                self.log(f"✅ Programa 1 completado. Línea procesada: {linea_procesada}")
                return True, linea_especifica
            else:
                self.log("❌ Programa 1 falló")
                return False, None
                
        except Exception as e:
            self.log(f"❌ Error en Programa 1: {e}")
            return False, None
    
    def ejecutar_programa2(self, linea_especifica):
        """Ejecutar Programa 2 - Automatización NSE"""
        try:
            nse = NSEAutomation(self.modelo.csv_file, linea_especifica=linea_especifica)
            nse.is_running = True
            
            if not self.modelo.csv_file:
                self.log(f"❌ ERROR: Archivo CSV no encontrado: {self.modelo.csv_file}")
                return False
            
            self.log(f"🎯 Procesando línea: {linea_especifica}")
            time.sleep(3)
            
            resultado = nse.execute_nse_script()
            
            if resultado:
                self.log("✅ Programa 2 finalizado exitosamente")
            else:
                self.log("❌ Programa 2 falló")
                
            return resultado
            
        except Exception as e:
            self.log(f"❌ Error en Programa 2: {e}")
            return False
    
    def ejecutar_programa3(self, linea_especifica):
        """Ejecutar Programa 3 - Servicios NSE"""
        try:
            nse_services = NSEServicesAutomation(self.modelo.csv_file, linea_especifica=linea_especifica)
            
            if not self.modelo.csv_file:
                self.log(f"❌ ERROR: Archivo CSV no encontrado: {self.modelo.csv_file}")
                return False
            
            self.log(f"🎯 Procesando línea: {linea_especifica}")
            
            if not nse_services.iniciar_ahk():
                self.log("❌ No se pudieron iniciar los servicios AHK")
                return False
            
            nse_services.is_running = True
            resultado = nse_services.procesar_linea_especifica()
            
            if resultado:
                self.log(f"✅ Programa 3 completado exitosamente")
            else:
                self.log(f"❌ Programa 3 falló")
            
            nse_services.detener_ahk()
            return resultado
            
        except Exception as e:
            self.log(f"❌ Error en Programa 3: {e}")
            return False
    
    def ejecutar_programa4(self, linea_especifica, kml_filename):
        """Ejecutar Programa 4 - Automatización GE"""
        try:
            ge_auto = GEAutomation(self.modelo.csv_file, linea_especifica=linea_especifica)
            ge_auto.is_running = True
            
            # Actualizar nombre KML
            ge_auto.set_kml_filename(kml_filename)
            
            if not self.modelo.csv_file:
                self.log(f"❌ ERROR: Archivo CSV no encontrado: {self.modelo.csv_file}")
                return False
            
            self.log(f"🎯 Procesando línea: {linea_especifica}")
            self.log(f"📁 Archivo KML: {kml_filename}")
            
            time.sleep(3)
            
            success = ge_auto.perform_actions()
            
            if success:
                self.log("✅ Programa 4 finalizado exitosamente")
            else:
                self.log("❌ Programa 4 falló")
                
            return success
            
        except Exception as e:
            self.log(f"❌ Error en Programa 4: {e}")
            return False