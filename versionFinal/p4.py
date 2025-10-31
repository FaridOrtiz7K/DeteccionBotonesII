import os
import time
import sys
import pyautogui
import pandas as pd
import cv2
import numpy as np
import threading
import logging
from utils.ahk_writer import AHKWriter
from utils.ahk_manager import AHKManager
from utils.ahk_enter import EnterAHKManager
from utils.ahk_click_down import AHKClickDownManager  # Nuevo manager para flechas abajo

logger = logging.getLogger(__name__)

class GEAutomation:
    def __init__(self):
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.reference_image = "img/textoAdicional.PNG"
        self.is_running = False
        self.is_paused = False
        self.pause_condition = threading.Condition()
        
        # Inicializar todos los AHK managers
        self.ahk_writer = AHKWriter()           # Para escritura de texto
        self.ahk_manager = AHKManager()         # Para ventana de archivo
        self.enter = EnterAHKManager()          # Para presionar Enter
        self.ahk_click_down = AHKClickDownManager()  # Para flechas abajo
        
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # COORDENADAS ABSOLUTAS (solo las necesarias)
        self.coords = {
            'agregar_ruta': (327, 381),
            'seleccionar_mapa': (168, 188),
            'anotar': (1366, 384),
            'agregar_texto_adicional': (1449, 452),
            'campo_texto': (1421, 526),
            'agregar_texto': (1263, 572),
            'cerrar_ventana_texto': (1338, 570),
            'limpiar_trazo': (360, 980),
            'lote_again': (70, 266)
        }
        
        # Coordenadas relativas para detección de imagen (campo de texto)
        self.coords_texto_relativas = {
            'campo_texto': (222, 54),
            'agregar_texto': (64, 100),
            'cerrar_ventana_texto': (139, 98)
        }

        # Bandera para control de errores
        self.b2_fallback_used = False

    def encontrar_ventana_archivo(self):
        """Busca la ventana de archivo usando template matching con reintentos inteligentes"""
        intentos = 1
        confianza_minima = 0.6
        tiempo_espera_base = 1
        tiempo_espera_largo = 10
        
        # Cargar template una sola vez fuera del bucle
        template = cv2.imread('img/cargarArchivo.png')
        if template is None:
            logger.error("No se pudo cargar la imagen 'cargarArchivo.png'")
            return None
        
        while self.is_running: 
            # Verificar si está pausado
            if self.is_paused:
                with self.pause_condition:
                    while self.is_paused and self.is_running:
                        self.pause_condition.wait()
                if not self.is_running:
                    return False
            
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
            template = cv2.imread('img/ventanaError.png') 
            if template is None:
                logger.error("No se pudo cargar la imagen 'ventanaError.png'")
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
                
                print("Ventana de error detectada y cerrada")
                self.enter.stop_ahk()
                return True
            else:
                return False
                
        except Exception as e:
            logger.error(f"Error al detectar ventana de error: {e}")
            return False

    def handle_b4_special_behavior(self, nombre_archivo):
        """Maneja el comportamiento especial para cargar archivos usando AHK"""
        # Hacer clic en agregar ruta (equivalente a b4)
        self.click(*self.coords['agregar_ruta'])
        
        # Esperar después de presionar el botón
        time.sleep(3)

        # Buscar la ventana de archivo
        coordenadas_ventana = self.encontrar_ventana_archivo()

        if coordenadas_ventana:
            x_ventana, y_ventana = coordenadas_ventana
            logger.info(f"Coordenadas ventana: x={x_ventana}, y={y_ventana}")
            
            # Calcular coordenadas del campo de texto
            x_campo = x_ventana + 294
            y_campo = y_ventana + 500
            logger.info(f"Coordenadas campo texto: x={x_campo}, y={y_campo}")
            
            # Iniciar AHK si no está corriendo
            if not self.ahk_manager.start_ahk():
                logger.error("No se pudo iniciar AutoHotkey")
                return False
            
            # Enviar comandos a AHK para escribir el nombre del archivo
            if self.ahk_manager.ejecutar_acciones_ahk(x_campo, y_campo, nombre_archivo):
                time.sleep(1.5)   # Esperar a que AHK termine
            else:
                logger.error("Error enviando comando a AHK")
                return False
            
            # Presionar Enter usando AHK para abrir el archivo
            if not self.enter.start_ahk():
                logger.error("No se pudo iniciar AutoHotkey para Enter")
                return False
                
            if self.enter.presionar_enter(1):
                time.sleep(2)
            else:
                logger.error("Error enviando comando Enter a AHK")
                return False
                
            self.enter.stop_ahk()
            self.ahk_manager.stop_ahk()
            
            # Hacer clic en agregar ruta nuevamente para cargar
            self.click(*self.coords['agregar_ruta'])
            time.sleep(2)
            
            return True
        else:
            logger.error("No se pudo encontrar la ventana de archivo.")
            return False

    def escribir_texto_adicional_con_ahk(self, x, y, texto):
        """Escribe texto adicional usando AHKWriter"""
        print(f"✍️ Escribiendo texto adicional con AHK: '{texto}'")
        
        if not self.ahk_writer.start_ahk():
            print("❌ No se pudo iniciar AHKWriter")
            return False
        
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, texto)
        
        if success:
            print("✅ Texto escrito exitosamente con AHK")
            time.sleep(1)
        else:
            print(f"❌ Error al escribir con AHK en ({x}, {y}): {texto}")
        
        self.ahk_writer.stop_ahk()
        return success

    def presionar_flecha_abajo_con_ahk(self):
        """Presiona flecha abajo usando AHK"""
        print("⬇️ Presionando flecha abajo con AHK")
        
        if not self.ahk_click_down.start_ahk():
            print("❌ No se pudo iniciar AHK para flecha abajo")
            return False
        
        success = self.ahk_click_down.presionar_flecha_abajo(1)
        
        if success:
            print("✅ Flecha abajo presionada exitosamente")
            time.sleep(1)
        else:
            print("❌ Error al presionar flecha abajo con AHK")
        
        self.ahk_click_down.stop_ahk()
        return success

    def handle_fallback_to_agregar_ruta(self):
        """Maneja el fallback a la coordenada agregar_ruta cuando hay errores"""
        print("Realizando fallback a coordenadas de agregar ruta...")
        self.click(*self.coords['agregar_ruta'])
        time.sleep(2)
        return True

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
            return True, max_loc  # Devuelve las coordenadas (x, y) de la esquina superior izquierda
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
            
            # Verificar si está pausado
            if self.is_paused:
                with self.pause_condition:
                    while self.is_paused and self.is_running:
                        self.pause_condition.wait()
                if not self.is_running:
                    return False, None
            
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

    def perform_actions(self):
        """Función principal que realiza todas las acciones"""
        # Iniciar todos los managers AHK
        managers = [self.ahk_writer, self.ahk_manager, self.enter, self.ahk_click_down]
        for manager in managers:
            if not manager.start_ahk():
                print(f"❌ No se pudo iniciar {manager.__class__.__name__}")
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

            # Realizar el bucle 9 veces (como en el código original)
            for iteration in range(1, 10):
                if iteration > total_lines:
                    print(f"⚠️  No hay más líneas en el CSV. Iteración {iteration} saltada.")
                    continue
                    
                print(f"🔄 Procesando iteración {iteration}/9")
                self.process_single_iteration(df, iteration, total_lines)
                
                # Verificar si está pausado entre iteraciones
                if self.is_paused:
                    with self.pause_condition:
                        while self.is_paused and self.is_running:
                            self.pause_condition.wait()
                    if not self.is_running:
                        break
                
                # Guardar cada 10 iteraciones (en este caso, solo al final del bucle)
                if iteration % 10 == 0:
                    self.save_progress()
                    
            # Guardar al final
            self.save_progress()
            print("✅ Script completado exitosamente!")
            return True
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            return False
        finally:
            # Detener todos los managers AHK
            for manager in managers:
                manager.stop_ahk()

    def process_single_iteration(self, df, iteration, total_lines):
        """Procesar una sola iteración del bucle"""
        # Obtener la fila correspondiente (0-indexed)
        row_index = iteration - 1
        row = df.iloc[row_index]
        
        # Obtener valores del CSV
        if len(row) >= 2:
            num_txt_type = str(row[0])  # Primera columna para el nombre del archivo
            texto_adicional = str(row[1])  # Segunda columna para el texto adicional
        else:
            print(f"⚠️  Fila {iteration} no tiene suficientes columnas")
            return

        # SECUENCIA DE ACCIONES
        try:
            # 1. Cargar archivo usando el método especial de b4 con AHK
            nombre_archivo = f"RA {num_txt_type}.kml"
            success = self.handle_b4_special_behavior(nombre_archivo)
            
            if not success:
                print("❌ Falló la carga del archivo, realizando fallback...")
                self.handle_fallback_to_agregar_ruta()
                return

            # 2. Seleccionar Lote
            self.click(*self.coords['lote_again'])
            self.sleep(2)
            
            # 3. Seleccionar en el mapa
            self.click(*self.coords['seleccionar_mapa'])
            self.sleep(2)
            
            # 4. Anotar
            self.click(*self.coords['anotar'])
            self.sleep(2)
            
            # 5. Agregar texto adicional
            self.click(*self.coords['agregar_texto_adicional'])
            self.sleep(2)
            
            # 6. DETECCIÓN DE IMAGEN para el campo de texto
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=10)
            
            if image_found:
                # Calcular coordenadas absolutas del campo de texto basado en la detección
                x_campo = base_location[0] + self.coords_texto_relativas['campo_texto'][0]
                y_campo = base_location[1] + self.coords_texto_relativas['campo_texto'][1]
                x_agregar = base_location[0] + self.coords_texto_relativas['agregar_texto'][0]
                y_agregar = base_location[1] + self.coords_texto_relativas['agregar_texto'][1]
                x_cerrar = base_location[0] + self.coords_texto_relativas['cerrar_ventana_texto'][0]
                y_cerrar = base_location[1] + self.coords_texto_relativas['cerrar_ventana_texto'][1]
                
                # Hacer clic en el campo de texto detectado
                self.click(x_campo, y_campo)
                self.sleep(2)
                
                # Escribir texto adicional con AHKWriter
                self.escribir_texto_adicional_con_ahk(x_campo, y_campo, texto_adicional)
                self.sleep(1)
                
                # 7. Agregar texto
                self.click(x_agregar, y_agregar)
                self.sleep(3)
            
                # 8. Cerrar ventana de texto
                self.click(x_cerrar, y_cerrar)
                self.sleep(2)
            else:
                # Fallback a coordenadas fijas si no se detecta la imagen
                print("⚠️  Usando coordenadas fijas para campo de texto")
                self.click(*self.coords['campo_texto'])
                self.sleep(2)
                
                # Limpiar campo con backspace usando AHK Enter
                if not self.enter.start_ahk():
                    print("❌ No se pudo iniciar AHK para backspace")
                else:
                    # Simular backspace múltiples veces
                    for _ in range(10):
                        self.enter.presionar_backspace(1)
                        time.sleep(0.1)
                    self.enter.stop_ahk()
                
                self.sleep(1)
                
                # Escribir texto adicional con AHKWriter
                self.escribir_texto_adicional_con_ahk(
                    self.coords['campo_texto'][0], 
                    self.coords['campo_texto'][1], 
                    texto_adicional
                )
                self.sleep(1)
                
                # 7. Agregar texto
                self.click(*self.coords['agregar_texto'])
                self.sleep(3)
            
                # 8. Cerrar ventana de texto
                self.click(*self.coords['cerrar_ventana_texto'])
                self.sleep(2)
            
            # 9. Limpiar trazo
            self.click(*self.coords['limpiar_trazo'])
            self.sleep(1)
            
            # 10. Seleccionar Lote nuevamente
            self.click(*self.coords['lote_again'])
            self.sleep(2)
            
            # 11. Presionar flecha abajo CON AHK
            self.presionar_flecha_abajo_con_ahk()
            self.sleep(2)
            
            # 12. Detectar ventana de error después de procesar
            if self.detectar_ventana_error():
                print("⚠️  Ventana de error detectada y manejada")
                # Realizar fallback a agregar_ruta
                self.handle_fallback_to_agregar_ruta()
            
            print(f"✅ Iteración {iteration} completada")
            
        except Exception as e:
            print(f"❌ Error en iteración {iteration}: {e}")
            # En caso de error, intentar fallback
            self.handle_fallback_to_agregar_ruta()

    def save_progress(self):
        """Guardar progreso con Ctrl + S"""
        print("💾 Guardando progreso...")
        pyautogui.hotkey('ctrl', 's')
        self.sleep(6)

    def pause_search(self):
        """Pausar la búsqueda"""
        if self.is_running and not self.is_paused:
            self.is_paused = True
            print("⏸️  Búsqueda pausada")

    def resume_search(self):
        """Reanudar la búsqueda"""
        if self.is_running and self.is_paused:
            self.is_paused = False
            with self.pause_condition:
                self.pause_condition.notify_all()
            print("▶️  Búsqueda reanudada")

    def stop_search(self):
        """Detener la búsqueda"""
        self.is_running = False
        self.is_paused = False
        with self.pause_condition:
            self.pause_condition.notify_all()
        
        # Detener todos los managers AHK
        managers = [self.enter, self.ahk_manager, self.ahk_writer, self.ahk_click_down]
        for manager in managers:
            manager.stop_ahk()
            
        print("⏹️  Proceso detenido")

def main():
    # Inicializar automatización
    ge_auto = GEAutomation()
    ge_auto.is_running = True
    
    # Verificar archivo CSV
    if not os.path.exists(ge_auto.csv_file):
        print(f"❌ ERROR: Archivo CSV no encontrado: {ge_auto.csv_file}")
        input("Presiona Enter para salir...")
        return
    
    print(f"✅ Archivo CSV encontrado: {ge_auto.csv_file}")
    
    # Verificar imagen de referencia
    if not os.path.exists(ge_auto.reference_image):
        print(f"⚠️  Advertencia: Imagen de referencia no encontrada: {ge_auto.reference_image}")
        print("   El proceso usará coordenadas fijas como fallback")
    else:
        print(f"✅ Imagen de referencia encontrada: {ge_auto.reference_image}")
    
    # Verificar imagen de ventana de archivo
    if not os.path.exists('img/cargarArchivo.png'):
        print(f"⚠️  Advertencia: Imagen de ventana de archivo no encontrada: img/cargarArchivo.png")
    else:
        print(f"✅ Imagen de ventana de archivo encontrada: img/cargarArchivo.png")
    
    # Verificar imagen de error
    if not os.path.exists('img/ventanaError.png'):
        print(f"⚠️  Advertencia: Imagen de error no encontrada: img/ventanaError.png")
    else:
        print(f"✅ Imagen de error encontrada: img/ventanaError.png")
    
    print()
    print("Configuración:")
    print(f"  - Archivo CSV: {ge_auto.csv_file}")
    print(f"  - Imagen de referencia: {ge_auto.reference_image}")
    print("  - Usando AHKWriter para escritura de texto")
    print("  - Usando AHKManager para ventana de archivo")
    print("  - Usando EnterAHKManager para tecla Enter")
    print("  - Usando AHKClickDown para flechas abajo")
    print("  - 9 iteraciones programadas")
    print("  - Usando detección inteligente de ventanas")
    print()
    
    try:
        input("Presiona Enter para INICIAR la automatización...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO AUTOMATIZACIÓN ...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar script
        success = ge_auto.perform_actions()
        
        if success:
            print("🎉 AUTOMATIZACIÓN COMPLETADA EXITOSAMENTE!")
        else:
            print("❌ La automatización encontró errores")
        
        print()
        input("Presiona Enter para salir...")
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
        ge_auto.stop_search()
        input("Presiona Enter para salir...")
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        ge_auto.stop_search()
        input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()