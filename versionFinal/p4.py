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
            'lote_again': (70, 266)
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
            
            # Verificar columna 28 (num_txt_type)
            if len(row) <= 28 or pd.isna(row.iloc[28]):
                print(f"⚠️  Columna 28 vacía o no existe en fila {row_index}, saltando...")
                return False
                
            # Verificar columna 29 (texto_adicional) - puede estar vacía pero debe existir
            if len(row) <= 29:
                print(f"⚠️  Columna 29 no existe en fila {row_index}, saltando...")
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

            # Realizar el bucle 9 veces
            for iteration in range(1, 10):
                row_index = iteration - 1
                
                # Verificar si debemos saltar esta iteración
                if row_index >= total_lines:
                    print(f"⚠️  No hay más líneas en el CSV. Iteración {iteration} saltada.")
                    continue
                    
                # Verificar si los valores del CSV son válidos
                if not self.verificar_valores_csv(df, row_index):
                    print(f"⚠️  Valores inválidos en fila {row_index}. Iteración {iteration} saltada.")
                    continue
                    
                print(f"🔄 Procesando iteración {iteration}/9")
                success = self.process_single_iteration(df, 3, total_lines)
                
                if not success:
                    print(f"⚠️  Iteración {iteration} falló, continuando con la siguiente...")
                
                # Guardar cada 10 iteraciones
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
        
        # Obtener valores del CSV con verificación
        try:
            num_txt_type = str(int(row.iloc[28])) if not pd.isna(row.iloc[28]) else None
            texto_adicional = str(row.iloc[29]) if not pd.isna(row.iloc[29]) else ""
        except (ValueError, IndexError) as e:
            print(f"❌ Error obteniendo valores del CSV: {e}")
            return False

        if not num_txt_type:
            print(f"⚠️  num_txt_type vacío en iteración {iteration}, saltando...")
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
            
            print(f"✅ Iteración {iteration} completada exitosamente")
            return True
            
        except Exception as e:
            print(f"❌ Error en iteración {iteration}: {e}")
            # Intentar cerrar ventana de error en caso de excepción
            self.detectar_ventana_error()
            return False

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
    
    # Verificar que el CSV tiene datos
    try:
        df = pd.read_csv(ge_auto.csv_file)
        if len(df) == 0:
            print("❌ ERROR: El archivo CSV está vacío")
            input("Presiona Enter para salir...")
            return
        print(f"✅ CSV tiene {len(df)} filas de datos")
    except Exception as e:
        print(f"❌ ERROR: No se pudo leer el CSV: {e}")
        input("Presiona Enter para salir...")
        return
    
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
    print("  - Verificación de CSV: Se saltarán filas vacías o inválidas")
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