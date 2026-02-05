import pandas as pd
import pyautogui
import cv2
import numpy as np
import time
import os
import logging
from utils.ahk_writer import AHKWriter
from utils.ahk_manager import AHKManager
from utils.ahk_enter import EnterAHKManager
from utils.ahk_click_down import AHKClickDown

logger = logging.getLogger(__name__)

class GEAutomation:
    def __init__(self, linea_especifica=None, csv_file="", kml_filename="NN", estado_global=None):
        self.linea_especifica = linea_especifica
        self.csv_file = csv_file
        self.estado_global = estado_global
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

        self.nombre = kml_filename
        
        self.coords = {
            'agregar_ruta': (327, 381),
            'archivo': (1396, 608),
            'abrir': (1406, 634),
            'seleccionar_mapa': (168, 188),
            'anotar': (1366, 384),
            'agregar_texto_adicional': (1449, 452),
            'limpiar_trazo': (360, 980),
            'lote_again': (83, 266),
            'cerrar_ventana_archivo': (1530, 555),
            'boton_documentos': (1120, 666),
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
                if self.estado_global and self.estado_global.esperar_si_pausado():
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
                            if self.estado_global and self.estado_global.esperar_si_pausado():
                                return None
                            time.sleep(1)
                    else:
                        for _ in range(int(tiempo_espera_base)):
                            if self.estado_global and self.estado_global.esperar_si_pausado():
                                return None
                            time.sleep(1)
                    intentos += 1
                    
            except Exception as e:
                logger.error(f"Error durante la búsqueda: {e}")
                for _ in range(int(tiempo_espera_base)):
                    if self.estado_global and self.estado_global.esperar_si_pausado():
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
                        if self.estado_global and self.estado_global.esperar_si_pausado():
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
            x_documento = x_ventana + 64
            y_documento = y_ventana + 315
            if not self.click(x_documento, y_documento):
                logger.error("No se pudo hacer clic en el botón de documentos")
                return False
            
            x_campo = x_ventana + 294
            y_campo = y_ventana + 500
            logger.info(f"Coordenadas campo texto: x={x_campo}, y={y_campo}")
            
            if not self.ahk_manager.start_ahk():
                logger.error("No se pudo iniciar AutoHotkey")
                return False
            
            if self.ahk_manager.ejecutar_acciones_ahk(x_campo, y_campo, nombre_archivo):
                for _ in range(2):
                    if self.estado_global and self.estado_global.esperar_si_pausado():
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
            logger.info("Texto adicional vacío, saltando escritura")
            return True
            
        texto_str = str(texto).strip()
        logger.info(f"Intentando escribir texto: '{texto_str}' en coordenadas ({x}, {y})")
        
        if x <= 0 or y <= 0:
            logger.error(f"Coordenadas inválidas: ({x}, {y})")
            return False
        
        if not self.ahk_writer.start_ahk():
            logger.error("No se pudo iniciar AHK Writer")
            return False
        
        logger.info("AHK Writer iniciado, enviando comando...")
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, texto_str)
        self.ahk_writer.stop_ahk()
        
        if success:
            logger.info(f"Texto escrito exitosamente: '{texto_str}'")
        else:
            logger.error(f"Error al escribir texto: '{texto_str}'")
            logger.info("Intentando método alternativo con pyautogui...")
            try:
                self.click(x, y)
                for _ in range(1):
                    if self.estado_global and self.estado_global.esperar_si_pausado():
                        return False
                    time.sleep(1.5)
                pyautogui.hotkey('ctrl', 'a')
                for _ in range(1):
                    if self.estado_global and self.estado_global.esperar_si_pausado():
                        return False
                    time.sleep(1)
                pyautogui.press('delete')
                for _ in range(1):
                    if self.estado_global and self.estado_global.esperar_si_pausado():
                        return False
                    time.sleep(1)
                pyautogui.write(texto_str, interval=0.1)
                logger.info(f"Texto escrito con pyautogui: '{texto_str}'")
                success = True
            except Exception as e:
                logger.error(f"También falló pyautogui: {e}")
                
        return success

    def presionar_flecha_abajo_ahk(self, x, y, veces=1):
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
            if self.estado_global and self.estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def sleep(self, seconds):
        for _ in range(int(seconds * 1.5)):
            if self.estado_global and self.estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def detect_image_with_cv2(self, image_path, confidence=0.7):
        try:
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            template = cv2.imread(image_path)
            if template is None:
                logger.error(f"No se pudo cargar la imagen {image_path}")
                return False, None
            
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            if max_val < confidence:
                logger.info(f"Imagen no encontrada. Mejor coincidencia: {max_val:.2f}")
                return False, None
            
            logger.info(f"Imagen encontrada con confianza: {max_val:.2f}")
            return True, max_loc
        except Exception as e:
            logger.error(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.7):
        logger.info(f"Buscando imagen: {image_path}")
        
        for attempt in range(1, max_attempts + 1):
            if self.estado_global and self.estado_global.esperar_si_pausado():
                return False, None
                
            found, location = self.detect_image_with_cv2(image_path, confidence)
            
            if found:
                logger.info(f"Imagen detectada en el intento {attempt} en coordenadas: {location}")
                return True, location
            
            logger.info(f"Intento {attempt}/{max_attempts} - Imagen no encontrada")
            
            if attempt < max_attempts:
                if attempt % 10 == 0:
                    logger.info("Espera prolongada de 12 segundos...")
                    for _ in range(12):
                        if self.estado_global and self.estado_global.esperar_si_pausado():
                            return False, None
                        time.sleep(1)
                else:
                    for _ in range(3):
                        if self.estado_global and self.estado_global.esperar_si_pausado():
                            return False, None
                        time.sleep(1)
        
        logger.error("Imagen no encontrada después de 30 intentos. Terminando proceso.")
        return False, None

    def verificar_valores_csv(self, df, row_index):
        try:
            if row_index >= len(df):
                logger.error(f"Fila {row_index} no existe en el CSV")
                return False
            
            row = df.iloc[row_index]
            if len(row) <= 27 or pd.isna(row.iloc[27]) or row.iloc[27] != 1:
                logger.info(f"Columna 27 vacía, no es 1 o no existe en fila {row_index}, saltando...")
                if not self.presionar_flecha_abajo_ahk(83, 266, 1):
                    logger.info("No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    logger.info("Flecha abajo presionada con AHK")
                self.sleep(3)
                self.presionar_enter_ahk(1)
                self.sleep(1)
                return False
            
            if len(row) <= 28 or pd.isna(row.iloc[28]):
                logger.info(f"Columna 28 vacía o no existe en fila {row_index}, saltando...")
                if not self.presionar_flecha_abajo_ahk(83, 266, 1):
                    logger.info("No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    logger.info("Flecha abajo presionada con AHK")
                self.sleep(2)
                self.presionar_enter_ahk(1)
                self.sleep(1)
                return False
                
            if len(row) <= 29:
                logger.info(f"Columna 29 no existe en fila {row_index}, saltando...")
                if not self.presionar_flecha_abajo_ahk(83, 266, 1):
                    logger.info("No se pudo presionar flecha abajo con AHK, usando pyautogui")
                    pyautogui.press('down')
                else:
                    logger.info("Flecha abajo presionada con AHK")
                self.sleep(2)
                self.presionar_enter_ahk(1)
                self.sleep(1)
                return False
                
            return True
            
        except Exception as e:
            logger.error(f"Error verificando valores CSV: {e}")
            return False

    def perform_actions(self):
        if not self.ahk_writer.start_ahk():
            logger.error("No se pudo iniciar AHKWriter")
            return False
            
        try:
            if not os.path.exists(self.csv_file):
                logger.error(f"El archivo CSV no existe: {self.csv_file}")
                return False

            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            if total_lines < 1:
                logger.error("No hay suficientes datos en el archivo CSV")
                return False

            logger.info(f"Total de líneas en CSV: {total_lines}")

            if self.linea_especifica is None:
                logger.error("No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                logger.error(f"Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False

            row_index = self.linea_especifica - 1
                
            if not self.verificar_valores_csv(df, row_index):
                logger.info(f"Valores inválidos en fila {row_index}. Línea {self.linea_especifica} saltada.")
                return True
                    
            logger.info(f"Procesando línea {self.linea_especifica}/{total_lines}")
            success = self.process_single_iteration(df, self.linea_especifica, total_lines)
                
            if not success:
                logger.info(f"Línea {self.linea_especifica} falló")
                return False
                
            return True
            
        except Exception as e:
            logger.error(f"Error durante la ejecución: {e}")
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
            logger.error(f"Error obteniendo valores del CSV: {e}")
            return False

        if not num_txt_type:
            logger.info(f"num_txt_type vacío en línea {linea_especifica}, saltando...")
            return False
        self.nombre = f"{self.nombre} {num_txt_type}.kml"

        logger.info(f"Archivo a cargar: {self.nombre}")
        logger.info(f"Texto adicional: '{texto_adicional}'")

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
                logger.error("No se pudo cargar el archivo. Regresando a agregar_ruta...")
                self.click(*self.coords['agregar_ruta'])
                self.sleep(3)
                return False
            
            if not self.presionar_enter_ahk(1):
                logger.info("No se pudo presionar Enter con AHK, usando pyautogui")
                pyautogui.press('enter')
            
            self.sleep(4)

            self.click(*self.coords['agregar_ruta'])
            self.sleep(3)

            self.click(1406, 675)
            self.sleep(3)

            self.click(83, 266)
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
                        logger.info("Falló la escritura con AHK, intentando con pyautogui...")
                        pyautogui.write(texto_adicional, interval=0.1)
                else:
                    logger.info("Texto adicional vacío, no se escribe nada")
                
                self.sleep(3)

                self.click(x_agregar, y_agregar)
                self.sleep(4)
            
                self.click(x_cerrar, y_cerrar)
                self.sleep(3)
            else:
                logger.error("No se pudo detectar la imagen del campo de texto")
                return False
                        
            self.click(*self.coords['limpiar_trazo'])
            self.sleep(2)
            
            self.click(*self.coords['lote_again'])
            self.sleep(3)
            
            if not self.presionar_flecha_abajo_ahk(*self.coords['lote_again'], 1):
                logger.info("No se pudo presionar flecha abajo con AHK, usando pyautogui")
                pyautogui.press('down')
            else:
                logger.info("Flecha abajo presionada con AHK")
            
            self.sleep(3)
            
            if self.detectar_ventana_error():
                logger.info("Ventana de error detectada y cerrada")
        
            self.presionar_enter_ahk(1)
            self.sleep(1)
            logger.info(f"Línea {linea_especifica} completada exitosamente")

            return True
            
        except Exception as e:
            logger.error(f"Error en línea {linea_especifica}: {e}")
            self.detectar_ventana_error()
            return False