import pandas as pd
import pyautogui
import cv2
import numpy as np
import time
import os
import logging
from utils.ahk_writer import AHKWriter

logger = logging.getLogger(__name__)

class NSEAutomation:
    def __init__(self, linea_especifica=None, csv_file="", estado_global=None):
        self.linea_especifica = linea_especifica
        self.csv_file = csv_file
        self.estado_global = estado_global
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
            if self.estado_global and self.estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def write_with_ahk(self, x, y, text):
        if pd.isna(text) or text is None or str(text).strip() == "":
            return True
            
        text_str = str(text).strip()
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, text_str)
        if not success:
            logger.error(f"Error al escribir con AHK en ({x}, {y}): {text_str}")
        for _ in range(1):
            if self.estado_global and self.estado_global.esperar_si_pausado():
                return False
            time.sleep(1.5)
        return success

    def sleep(self, seconds):
        for _ in range(int(seconds * 1.5)):
            if self.estado_global and self.estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def detect_image_with_cv2(self, image_path, confidence=0.6):
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

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.6):
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
                    logger.info("Espera prolongada de 15 segundos...")
                    for _ in range(15):
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
            logger.error("No se pudo iniciar AHKWriter")
            return False
            
        try:
            try:
                df = pd.read_csv(self.csv_file)
                if len(df) == 0:
                    logger.info("CSV vacío. Saltando Programa 2...")
                    return True
            except:
                pass
            
            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            logger.info(f"Total de líneas en CSV: {total_lines}")
            
            if self.linea_especifica is None:
                logger.error("No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                logger.error(f"Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False
            
            row = df.iloc[self.linea_especifica - 1]
            logger.info(f"Procesando línea {self.linea_especifica}/{total_lines}")
            
            if self.should_skip_process(row):
                logger.info(f"Saltando línea {self.linea_especifica} - Columna 6 tiene valor: {row.iloc[5]}")
                return True
            
            if str(row.iloc[4]).strip().upper() != "V":
                logger.info(f"Saltando línea {self.linea_especifica} - No es tipo V: {row.iloc[4]}")
                return True
            
            self.click(169, 189)
            self.sleep(3)
            self.click(1491, 386)
            self.sleep(3)
            
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=30)
            
            if not image_found:
                logger.error("No se puede continuar sin detectar la imagen de referencia.")
                return False
            
            logger.info("Imagen detectada, procediendo con tipo V")
            self.handle_type_v(row, base_location)
            
            logger.info(f"Línea {self.linea_especifica} completada (hasta CERRAR)")
            return True
            
        except Exception as e:
            logger.error(f"Error durante la ejecución: {e}")
            return False
        finally:
            self.ahk_writer.stop_ahk()

    def handle_type_v(self, row, base_location):
        base_x, base_y = base_location
        
        for col_index in range(7, 18):
            if self.estado_global and self.estado_global.esperar_si_pausado():
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
        