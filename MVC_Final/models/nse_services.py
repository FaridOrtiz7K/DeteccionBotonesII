import pandas as pd
import pyautogui
import cv2
import numpy as np
import time
import logging
from PIL import ImageGrab
from utils.ahk_writer import AHKWriter
from utils.ahk_click_down import AHKClickDown
from utils.ahk_enter import EnterAHKManager

logger = logging.getLogger(__name__)

class NSEServicesAutomation:
    def __init__(self, linea_especifica=None, csv_file="", estado_global=None):
        self.linea_especifica = linea_especifica
        self.csv_file = csv_file
        self.estado_global = estado_global
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
        logger.info(f"Buscando imagen: {imagen_path}")
        
        try:
            template = cv2.imread(imagen_path)
            if template is None:
                logger.error(f"No se pudo cargar la imagen: {imagen_path}")
                return None
            
            template_height, template_width = template.shape[:2]
            
            for intento in range(timeout):
                if self.estado_global and self.estado_global.esperar_si_pausado():
                    return None
                    
                screenshot = ImageGrab.grab()
                screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confidence:
                    x, y = max_loc
                    logger.info(f"Imagen encontrada en intento {intento + 1} - Coordenadas: ({x}, {y}) - Confianza: {max_val:.2f}")
                    return (x, y)
                
                logger.info(f"Intento {intento + 1}/{timeout} - Confianza máxima: {max_val:.2f}")
                
                for _ in range(2):
                    if self.estado_global and self.estado_global.esperar_si_pausado():
                        return None
                    time.sleep(1)
            
            logger.error(f"No se encontró la imagen después de {timeout} intentos")
            return None
            
        except Exception as e:
            logger.error(f"Error en búsqueda de imagen: {e}")
            return None

    def actualizar_coordenadas_relativas(self, referencia):
        if referencia is None:
            logger.error("No se puede actualizar coordenadas: referencia es None")
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
        logger.info("Coordenadas actualizadas a relativas")
        return True

    def iniciar_ahk(self):
        try:
            if not self.ahk_writer.start_ahk():
                logger.error("No se pudo iniciar AHK Writer")
                return False
            if not self.ahk_click_down.start_ahk():
                logger.error("No se pudo iniciar AHK Click Down")
                return False
            if not self.ahk_enter.start_ahk():
                logger.error("No se pudo iniciar AHK Enter")
                return False
            logger.info("Todos los servicios AHK iniciados correctamente")
            return True
        except Exception as e:
            logger.error(f"Error iniciando servicios AHK: {e}")
            return False

    def detener_ahk(self):
        try:
            self.ahk_writer.stop_ahk()
            self.ahk_click_down.stop_ahk()
            self.ahk_enter.stop_ahk()
            logger.info("Todos los servicios AHK detenidos correctamente")
        except Exception as e:
            logger.error(f"Error deteniendo servicios AHK: {e}")

    def click(self, x, y, duration=0.2):
        pyautogui.click(x, y, duration=duration)
        for _ in range(1):
            if self.estado_global and self.estado_global.esperar_si_pausado():
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
                logger.info(f"Texto escrito exitosamente: {text}")
                for _ in range(2):
                    if self.estado_global and self.estado_global.esperar_si_pausado():
                        return False
                    time.sleep(1)
            else:
                logger.error(f"Error al escribir texto: {text}")
            return success
        
        except Exception as e:
            logger.error(f"Error escribiendo texto '{text}': {e}")
            return False

    def press_down(self, x, y, times=1):
        try:
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                click_coords = (x, y)
            else:
                click_coords = (x, y)
                
            return self.ahk_click_down.ejecutar_click_down(click_coords[0], click_coords[1], times)
        except Exception as e:
            logger.error(f"Error presionando DOWN {times} veces: {e}")
            return False

    def press_enter(self):
        try:                
            return self.ahk_enter.presionar_enter(1)
        except Exception as e:
            logger.error(f"Error presionando enter: {e}")
            return False

    def sleep(self, seconds):
        for _ in range(int(seconds * 1.5)):
            if self.estado_global and self.estado_global.esperar_si_pausado():
                return
            time.sleep(1)

    def handle_error_click(self):
        for _ in range(5):
            if self.estado_global and self.estado_global.esperar_si_pausado():
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
            
            logger.info(f"Total de líneas en CSV: {total_lines}")
            
            if self.linea_especifica is None:
                logger.error("No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                logger.error(f"Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False
            
            linea_idx = self.linea_especifica - 1
            self.current_line = self.linea_especifica
            
            logger.info(f"PROCESANDO LÍNEA ESPECÍFICA: {self.current_line}/{total_lines}")
            
            row = df.iloc[linea_idx]
            
            if pd.notna(row.iloc[17]) and row.iloc[17] > 0:
                logger.info(f"Línea {self.current_line} tiene servicios para procesar")
                
                self.click(*self.coords['inicio_servicios'])
                self.sleep(3)
                
                logger.info("Buscando ventana de servicios...")
                referencia = self.buscar_imagen("img/ventanaAdministracion4.PNG", timeout=30)
                
                if referencia is None:
                    logger.error("ERROR: No se pudo encontrar la ventana de servicios")
                    return False
                
                if not self.actualizar_coordenadas_relativas(referencia):
                    logger.error("ERROR: No se pudieron actualizar las coordenadas relativas")
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
                    if self.estado_global and self.estado_global.esperar_si_pausado():
                        return False
                        
                    if pd.notna(row.iloc[col_idx]) and row.iloc[col_idx] > 0:
                        logger.info(f"  └─ Procesando {nombre}: {row.iloc[col_idx]}")
                        metodo(str(int(row.iloc[col_idx])))
                        servicios_procesados += 1
                
                self.click(*self.coords_relativas['cierre'])
                self.sleep(5)
                
                logger.info(f"Línea {self.current_line} completada: {servicios_procesados} servicios procesados")
                return True
            else:
                logger.info(f"Línea {self.current_line} no tiene servicios para procesar")
                return True
            
        except Exception as e:
            logger.error(f"Error procesando línea {self.current_line}: {e}")
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
        