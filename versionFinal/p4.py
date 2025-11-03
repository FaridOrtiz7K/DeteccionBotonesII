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
from utils.ahk_click_down import AHKClickDown

logger = logging.getLogger(__name__)

class GEAutomation:
    def __init__(self):
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.reference_image = "img/textoAdicional.PNG"
        self.ventana_archivo_img = "img/cargarArchivo.png"
        self.ventana_error_img = "img/ventanaError.png"
        self.is_running = False
        
        # Inicializar todos los manejadores AHK
        self.ahk_writer = AHKWriter()  # Para escritura de texto
        self.ahk_manager = AHKManager()  # Para manejar ventanas de archivo
        self.enter = EnterAHKManager()  # Para presionar Enter
        self.ahk_click_down = AHKClickDown()  # Para flechas abajo (puede reutilizarse o crearse uno específico)
        
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # COORDENADAS ABSOLUTAS (solo las necesarias)
        self.coords = {
            'agregar_ruta': (327, 381),
            'archivo': (1396, 608),
            'abrir': (1406, 634),
            'seleccionar_mapa': (168, 188),
            'anotar': (1366, 384),
            'agregar_texto_adicional': (1449, 452),
            'limpiar_trazo': (360, 980),
            'lote_again': (70, 266)
        }
        
        # Coordenadas relativas para detección de imagen (campo de texto)
        self.coords_texto_relativas = {
            'campo_texto': (222, 54),
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
        if not self.ahk_writer.start_ahk():
            logger.error("No se pudo iniciar AHK Writer")
            return False
        
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, texto)
        self.ahk_writer.stop_ahk()
        return success

    def presionar_flecha_abajo_ahk(self, veces=1):
        """Presiona flecha abajo usando AHK"""
        # Para flecha abajo, podemos usar el mismo EnterAHKManager o uno específico
        if not self.ahk_click_down.start_ahk():
            logger.error("No se pudo iniciar AutoHotkey para flecha abajo")
            return False
        
        try:
            for i in range(veces):
                # Simular flecha abajo - esto depende de cómo esté implementado en tu AHK
                # Si no hay método específico, podrías necesitar extender la clase
                pyautogui.press('down')
                time.sleep(0.5)
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
        # Iniciar AHKWriter
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

            # Realizar el bucle 9 veces (como en el código original)
            for iteration in range(1, 10):
                if iteration > total_lines:
                    print(f"⚠️  No hay más líneas en el CSV. Iteración {iteration} saltada.")
                    continue
                    
                print(f"🔄 Procesando iteración {iteration}/9")
                self.process_single_iteration(df, iteration, total_lines)
                
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
            # Detener todos los manejadores AHK
            self.ahk_writer.stop_ahk()
            self.ahk_manager.stop_ahk()
            self.enter.stop_ahk()
            self.ahk_click_down.stop_ahk()

    def process_single_iteration(self, df, iteration, total_lines):
        """Procesar una sola iteración del bucle"""
        # Obtener la fila correspondiente (0-indexed)
        row_index = iteration - 1
        row = df.iloc[row_index]
        
        # Obtener valores del CSV
        if len(row) >= 2:
            num_txt_type = str(row.iloc[29])  # Primera columna para el nombre del archivo
            texto_adicional = str(row.iloc[30])  # Segunda columna para el texto adicional
        else:
            print(f"⚠️  Fila {iteration} no tiene suficientes columnas")
            return

        # SECUENCIA DE ACCIONES
        try:
            # 1. Seleccionar Agregar ruta de GE
            self.click(*self.coords['agregar_ruta'])
            self.sleep(2)
            
            # 2. Usar detección de ventana de archivo para cargar el archivo con AHK Manager
            nombre_archivo = f"NM {num_txt_type}.kml"
            success = self.handle_archivo_special_behavior(nombre_archivo)
            
            if not success:
                print("❌ No se pudo cargar el archivo. Regresando a agregar_ruta...")
                # Regresar a agregar_ruta en caso de error
                self.click(*self.coords['agregar_ruta'])
                self.sleep(2)
                return
            
            # 3. Presionar Enter con AHK para confirmar la carga del archivo
            if not self.presionar_enter_ahk(1):
                print("⚠️  No se pudo presionar Enter con AHK, usando pyautogui")
                pyautogui.press('enter')
            
            self.sleep(3)
            
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
                
                # Hacer clic en el campo de texto detectado
                self.click(x_campo, y_campo)
                self.sleep(2)
                
                # Escribir texto adicional con AHK Writer
                if not self.escribir_texto_adicional_ahk(x_campo, y_campo, texto_adicional):
                    print("⚠️  No se pudo escribir con AHK Writer, usando pyautogui")
                    pyautogui.write(texto_adicional, interval=0.05)
                
                self.sleep(1)
                
                # 8. Agregar texto
                self.click(x_agregar, y_agregar)
                self.sleep(3)
            
                # 9. Cerrar ventana de texto
                self.click(x_cerrar, y_cerrar)
                self.sleep(2)
            else:
                # Fallback a coordenadas fijas si no se detecta la imagen
                print("⚠️  Usando coordenadas fijas para campo de texto")
                # Nota: Las coordenadas fijas fueron eliminadas, usar detección es obligatorio
                print("❌ No se pudo detectar la imagen y no hay coordenadas de fallback")
                return
            
            # 10. Limpiar trazo
            self.click(*self.coords['limpiar_trazo'])
            self.sleep(1)
            
            # 11. Seleccionar Lote nuevamente
            self.click(*self.coords['lote_again'])
            self.sleep(2)
            
            # 12. Presionar flecha abajo con AHK
            if not self.presionar_flecha_abajo_ahk(1):
                print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                pyautogui.press('down')
            
            self.sleep(2)
            
            # 13. Detectar ventana de error después de cada iteración
            if self.detectar_ventana_error():
                print("✅ Ventana de error detectada y cerrada")
            
            print(f"✅ Iteración {iteration} completada")
            
        except Exception as e:
            print(f"❌ Error en iteración {iteration}: {e}")
            # Intentar cerrar ventana de error en caso de excepción
            self.detectar_ventana_error()

    def save_progress(self):
        """Guardar progreso con Ctrl + S"""
        print("💾 Guardando progreso...")
        pyautogui.hotkey('ctrl', 's')
        self.sleep(6)

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
    
    # Verificar imágenes de referencia
    images_to_check = [
        ge_auto.reference_image,
        ge_auto.ventana_archivo_img,
        ge_auto.ventana_error_img
    ]
    
    for image_path in images_to_check:
        if not os.path.exists(image_path):
            print(f"⚠️  Advertencia: Imagen no encontrada: {image_path}")
        else:
            print(f"✅ Imagen encontrada: {image_path}")
    
    print()
    print("Configuración:")
    print(f"  - Archivo CSV: {ge_auto.csv_file}")
    print(f"  - Imagen de referencia: {ge_auto.reference_image}")
    print(f"  - Imagen ventana archivo: {ge_auto.ventana_archivo_img}")
    print(f"  - Imagen ventana error: {ge_auto.ventana_error_img}")
    print("  - Usando AHK Writer para escritura de texto adicional")
    print("  - Usando AHK Manager para carga de archivos")
    print("  - Usando AHK Enter para presionar Enter")
    print("  - Usando AHK para flechas abajo")
    print("  - 9 iteraciones programadas")
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
        ge_auto.is_running = False
        input("Presiona Enter para salir...")
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        ge_auto.is_running = False
        input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()