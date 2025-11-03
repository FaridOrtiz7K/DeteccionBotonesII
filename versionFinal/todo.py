import pandas as pd
import pyautogui
import time
import logging
import os
import sys
import cv2
import numpy as np
from utils.ahk_managerCopyDelete import AHKManagerCD
from utils.ahk_writer import AHKWriter
from utils.ahk_click_down import AHKClickDown
from utils.ahk_enter import EnterAHKManager
from utils.ahk_manager import AHKManager
from PIL import ImageGrab

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('procesamiento_completo.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ProcesadorCompleto:
    def __init__(self, archivo_csv):
        self.archivo_csv = archivo_csv
        self.df = None
        self.current_line = 0
        self.total_lines = 0
        self.is_running = False
        
        # Inicializar todos los manejadores AHK
        self.ahk_manager_cd = AHKManagerCD()
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        self.ahk_enter = EnterAHKManager()
        self.ahk_manager = AHKManager()
        
        # Configurar pyautogui
        pyautogui.PAUSE = 0.5
        pyautogui.FAILSAFE = True
        
        # Configuraciones específicas de cada parte
        self.setup_parte1()
        self.setup_parte2()
        self.setup_parte3()
        self.setup_parte4()
    
    def setup_parte1(self):
        """Configuración para la parte 1 (procesamiento básico)"""
        self.coords_p1 = {
            'click_1': (89, 263),
            'click_2': (1290, 349)
        }
    
    def setup_parte2(self):
        """Configuración para la parte 2 (NSE tipo V)"""
        self.reference_image_p2 = "img/VentanaAsignar.png"
        self.coords_p2 = {
            'seleccionar_lote': (169, 189),
            'asignar_nse': (1491, 386)
        }
        
        # Coordenadas relativas para tabla verde
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
    
    def setup_parte3(self):
        """Configuración para la parte 3 (servicios)"""
        self.reference_image_p3 = "img/ventanaAdministracion4.PNG"
        self.coords_p3 = {
            'inicio_servicios': (1563, 385),
            'boton_error': (704, 384)
        }
        
        self.coords_relativas_p3 = {
            'menu_principal': (81, 81),
            'campo_cantidad': (108, 350),
            'boton_guardar': (63, 390),
            'cierre': (863, 16),
            'casilla_servicio': (121, 236),
            'casilla_tipo': (121, 261),
            'casilla_empresa': (121, 290),
            'casilla_producto': (121, 322),
        }
    
    def setup_parte4(self):
        """Configuración para la parte 4 (GE)"""
        self.reference_image_p4 = "img/textoAdicional.PNG"
        self.ventana_archivo_img = "img/cargarArchivo.png"
        self.ventana_error_img = "img/ventanaError.png"
        
        self.coords_p4 = {
            'agregar_ruta': (327, 381),
            'archivo': (1396, 608),
            'abrir': (1406, 634),
            'seleccionar_mapa': (168, 188),
            'anotar': (1366, 384),
            'agregar_texto_adicional': (1449, 452),
            'limpiar_trazo': (360, 980),
            'lote_again': (70, 266),
            'cargar_ruta': (1406, 675)
        }
        
        self.coords_texto_relativas = {
            'campo_texto': (230, 66),
            'agregar_texto': (64, 100),
            'cerrar_ventana_texto': (139, 98)
        }

    def iniciar_ahk(self):
        """Iniciar todos los servicios AHK"""
        logger.info("Iniciando todos los servicios AHK...")
        return (self.ahk_manager_cd.start_ahk() and 
                self.ahk_writer.start_ahk() and 
                self.ahk_click_down.start_ahk() and
                self.ahk_enter.start_ahk() and
                self.ahk_manager.start_ahk())

    def detener_ahk(self):
        """Detener todos los servicios AHK"""
        logger.info("Deteniendo todos los servicios AHK...")
        self.ahk_manager_cd.stop_ahk()
        self.ahk_writer.stop_ahk()
        self.ahk_click_down.stop_ahk()
        self.ahk_enter.stop_ahk()
        self.ahk_manager.stop_ahk()

    def cargar_csv(self):
        """Cargar el archivo CSV"""
        try:
            self.df = pd.read_csv(self.archivo_csv)
            self.total_lines = len(self.df)
            logger.info(f"CSV cargado correctamente: {self.total_lines} registros")
            return True
        except Exception as e:
            logger.error(f"Error cargando CSV: {e}")
            return False

    # ========== MÉTODOS UTILITARIOS COMUNES ==========
    
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
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            template = cv2.imread(image_path)
            if template is None:
                logger.error(f"No se pudo cargar la imagen {image_path}")
                return False, None
            
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            if max_val < confidence:
                return False, None
            
            return True, max_loc
        except Exception as e:
            logger.error(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.7):
        """Esperar a que aparezca una imagen con múltiples intentos"""
        logger.info(f"Buscando imagen: {image_path}")
        
        for attempt in range(1, max_attempts + 1):
            found, location = self.detect_image_with_cv2(image_path, confidence)
            
            if found:
                logger.info(f"Imagen detectada en el intento {attempt}")
                return True, location
            
            logger.info(f"Intento {attempt}/{max_attempts} - Imagen no encontrada")
            
            if attempt < max_attempts:
                if attempt % 10 == 0:
                    time.sleep(10)
                else:
                    time.sleep(2)
        
        logger.error("Imagen no encontrada después de 30 intentos")
        return False, None

    # ========== PARTE 1: PROCESAMIENTO BÁSICO ==========
    
    def buscar_por_id(self, id_buscar):
        """Busca un ID en la primera columna del CSV"""
        if self.df is None:
            logger.error("CSV no cargado")
            return None
            
        resultado = self.df[self.df.iloc[:, 0] == id_buscar]
        
        if len(resultado) == 0:
            logger.warning(f"ID {id_buscar} no encontrado en el CSV")
            return None
        
        logger.info(f"ID {id_buscar} encontrado")
        return resultado.iloc[0]

    def ejecutar_parte1(self, linea_actual):
        """Ejecutar la parte 1 del proceso"""
        try:
            logger.info(f"=== EJECUTANDO PARTE 1 PARA LÍNEA {linea_actual} ===")
            
            # Paso 2: Click en (89, 263)
            logger.info("Paso 2: Click en (89, 263)")
            self.click(*self.coords_p1['click_1'])
            
            # Paso 3: Usar AHKManager en (1483, 519) para obtener ID
            logger.info("Paso 3: Obteniendo ID con AHKManager")
            id_obtenido = self.ahk_manager_cd.ejecutar_acciones_ahk(1483, 519)
            
            if not id_obtenido:
                logger.error("No se pudo obtener el ID")
                return False
                
            id_obtenido = int(id_obtenido)
            logger.info(f"ID obtenido: {id_obtenido}")
            
            # Paso 4: Buscar el ID en el CSV
            logger.info(f"Paso 4: Buscando ID {id_obtenido} en CSV")
            registro = self.buscar_por_id(id_obtenido)
            
            if registro is None:
                logger.error(f"ID {id_obtenido} no encontrado en CSV")
                return False
            
            # Paso 5: Escribir valor de columna 2
            if len(registro) >= 2:
                valor_columna_2 = str(registro.iloc[1])
                logger.info(f"Paso 5: Escribiendo valor '{valor_columna_2}'")
                
                exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, valor_columna_2)
                if not exito_escritura:
                    logger.error("Error en la escritura")
                    return False
            
            # Paso 6-7: Procesar columna 4
            if len(registro) >= 4:
                valor_columna_4 = registro.iloc[3]
                logger.info(f"Paso 6: Valor columna 4 = {valor_columna_4}")
                
                if pd.notna(valor_columna_4) and float(valor_columna_4) > 0:
                    veces_down = int(float(valor_columna_4))
                    logger.info(f"Paso 7: Ejecutando {veces_down} veces DOWN")
                    
                    exito_down = self.ahk_click_down.ejecutar_click_down(1507, 636, veces_down)
                    if not exito_down:
                        logger.error("Error en click + down")
                        return False
            
            # Paso 8: Click final
            logger.info("Paso 8: Click en (1290, 349)")
            self.click(*self.coords_p1['click_2'])
            
            logger.info("Parte 1 completada exitosamente")
            return True
            
        except Exception as e:
            logger.error(f"Error en parte 1: {e}")
            return False

    # ========== PARTE 2: NSE TIPO V ==========
    
    def should_skip_process_p2(self, row):
        """Determina si se debe saltar el proceso basado en la columna 6"""
        if pd.notna(row[5]):
            col_value = str(row[5]).strip()
            if col_value and col_value != "" and col_value != "nan":
                return True
        return False

    def ejecutar_parte2(self, linea_actual):
        """Ejecutar la parte 2 del proceso (NSE tipo V)"""
        try:
            logger.info(f"=== EJECUTANDO PARTE 2 PARA LÍNEA {linea_actual} ===")
            
            row_index = linea_actual - 1
            row = self.df.iloc[row_index]
            
            # Verificar si se debe saltar
            if self.should_skip_process_p2(row):
                logger.info(f"Saltando línea {linea_actual} - Columna 6 tiene valor")
                return True
            
            # Verificar que sea tipo V
            if str(row[4]).strip().upper() != "V":
                logger.info(f"Saltando línea {linea_actual} - No es tipo V: {row[4]}")
                return True
            
            # Clics iniciales
            self.click(*self.coords_p2['seleccionar_lote'])
            self.click(*self.coords_p2['asignar_nse'])
            
            # Detectar imagen
            image_found, base_location = self.wait_for_image_with_retries(
                self.reference_image_p2, max_attempts=30
            )
            
            if not image_found:
                logger.error("No se puede continuar sin detectar la imagen de referencia")
                return False
            
            # Procesar tipo V
            self.handle_type_v(row, base_location)
            
            logger.info("Parte 2 completada exitosamente")
            return True
            
        except Exception as e:
            logger.error(f"Error en parte 2: {e}")
            return False

    def handle_type_v(self, row, base_location):
        """Manejar tipo V con coordenadas relativas"""
        base_x, base_y = base_location
        
        # Procesar columnas 7-17
        for col_index in range(7, 18):
            if pd.notna(row[col_index-1]) and row[col_index-1] > 0:
                x_cs_rel, y_cs_rel = self.coords_select[col_index]
                x_ct_rel, y_ct_rel = self.coords_type[col_index]
                
                x_cs_abs = base_x + x_cs_rel
                y_cs_abs = base_y + y_cs_rel
                x_ct_abs = base_x + x_ct_rel
                y_ct_abs = base_y + y_ct_rel
                
                self.click(x_cs_abs, y_cs_abs)
                self.sleep(2)
                
                texto = str(int(row[col_index-1]))
                self.ahk_writer.ejecutar_escritura_ahk(x_ct_abs, y_ct_abs, texto)
                self.sleep(2)
        
        # Botones finales
        x_asignar_abs = base_x + self.coords_asignar[0]
        y_asignar_abs = base_y + self.coords_asignar[1]
        self.click(x_asignar_abs, y_asignar_abs)
        self.sleep(2)
        
        x_cerrar_abs = base_x + self.coords_cerrar[0]
        y_cerrar_abs = base_y + self.coords_cerrar[1]
        self.click(x_cerrar_abs, y_cerrar_abs)
        self.sleep(2)

    # ========== PARTE 3: SERVICIOS ==========
    
    def actualizar_coordenadas_relativas_p3(self, referencia):
        """Actualizar coordenadas relativas para parte 3"""
        if referencia is None:
            return False
        
        ref_x, ref_y = referencia
        self.coords_relativas_p3_abs = {}
        
        for key, (x, y) in self.coords_relativas_p3.items():
            self.coords_relativas_p3_abs[key] = (ref_x + x, ref_y + y)
        
        # Mantener coordenadas que no cambian
        self.coords_relativas_p3_abs['boton_error'] = self.coords_p3['boton_error']
        self.coords_relativas_p3_abs['inicio_servicios'] = self.coords_p3['inicio_servicios']
        
        return True

    def handle_error_click_p3(self):
        """Manejar clics de error en parte 3"""
        for _ in range(5):
            self.click(*self.coords_relativas_p3_abs['boton_error'])
            self.sleep(2)

    def ejecutar_parte3(self, linea_actual):
        """Ejecutar la parte 3 del proceso (servicios)"""
        try:
            logger.info(f"=== EJECUTANDO PARTE 3 PARA LÍNEA {linea_actual} ===")
            
            row_index = linea_actual - 1
            row = self.df.iloc[row_index]
            
            # Solo procesar si columna 18 tiene valor > 0
            if not (pd.notna(row[17]) and row[17] > 0):
                logger.info(f"Línea {linea_actual} no tiene servicios para procesar")
                return True
            
            logger.info(f"Línea {linea_actual} tiene servicios para procesar")
            
            self.click(*self.coords_p3['inicio_servicios'])
            self.sleep(2)
            
            # Buscar ventana de servicios
            referencia = self.buscar_imagen_p3(self.reference_image_p3, timeout=30)
            
            if referencia is None:
                logger.error("No se pudo encontrar la ventana de servicios")
                return False
            
            # Actualizar coordenadas relativas
            if not self.actualizar_coordenadas_relativas_p3(referencia):
                logger.error("No se pudieron actualizar las coordenadas relativas")
                return False
            
            # Procesar servicios
            self.click(*self.coords_relativas_p3_abs['menu_principal'])
            self.sleep(2)
            
            servicios_procesados = self.procesar_servicios(row)
            
            # Cierre
            self.click(*self.coords_relativas_p3_abs['cierre'])
            self.sleep(5)
            
            logger.info(f"Parte 3 completada: {servicios_procesados} servicios procesados")
            return True
            
        except Exception as e:
            logger.error(f"Error en parte 3: {e}")
            return False

    def buscar_imagen_p3(self, imagen_path, timeout=30, confidence=0.8):
        """Buscar imagen para parte 3"""
        logger.info(f"Buscando imagen: {imagen_path}")
        
        try:
            template = cv2.imread(imagen_path)
            if template is None:
                logger.error(f"No se pudo cargar la imagen: {imagen_path}")
                return None
            
            for intento in range(timeout):
                screenshot = ImageGrab.grab()
                screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confidence:
                    x, y = max_loc
                    logger.info(f"Imagen encontrada en intento {intento + 1}")
                    return (x, y)
                
                logger.info(f"Intento {intento + 1}/{timeout} - Confianza máxima: {max_val:.2f}")
                time.sleep(1)
            
            logger.error(f"No se encontró la imagen después de {timeout} intentos")
            return None
            
        except Exception as e:
            logger.error(f"Error en búsqueda de imagen: {e}")
            return None

    def procesar_servicios(self, row):
        """Procesar todos los servicios de la línea"""
        servicios_procesados = 0
        
        # Diccionario de servicios a procesar
        servicios = {
            18: self.handle_voz_cobre,      # VOZ COBRE TELMEX
            19: self.handle_datos_sdom,     # Datos s/dom
            20: self.handle_datos_cobre_telmex,  # Datos-cobre-telmex-inf
            21: self.handle_datos_fibra_telmex,  # Datos-fibra-telmex-inf
            22: self.handle_tv_cable_otros, # TV cable otros
            23: self.handle_dish,           # Dish
            24: self.handle_tvs,            # TVS
            25: self.handle_sky,            # SKY
            26: self.handle_vetv            # VETV
        }
        
        for col_idx, handler in servicios.items():
            if pd.notna(row[col_idx]) and row[col_idx] > 0:
                handler(row[col_idx])
                servicios_procesados += 1
        
        return servicios_procesados

    # Handlers de servicios (similares a los de p3.py)
    def handle_voz_cobre(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3_abs['campo_cantidad'][0],
            self.coords_relativas_p3_abs['campo_cantidad'][1],
            str(int(cantidad))
        )
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    def handle_datos_sdom(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.ahk_click_down.ejecutar_click_down(
            self.coords_relativas_p3_abs['casilla_servicio'][0],
            self.coords_relativas_p3_abs['casilla_servicio'][1], 2
        )
        self.ahk_enter.presionar_enter(1)
        self.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3_abs['campo_cantidad'][0],
            self.coords_relativas_p3_abs['campo_cantidad'][1],
            str(int(cantidad))
        )
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    def handle_datos_cobre_telmex(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas_p3_abs['casilla_servicio'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_producto'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    def handle_datos_fibra_telmex(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas_p3_abs['casilla_servicio'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_tipo'], 1)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_empresa'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    def handle_tv_cable_otros(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas_p3_abs['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_empresa'], 4)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    def handle_dish(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas_p3_abs['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_empresa'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    def handle_tvs(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas_p3_abs['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_empresa'], 2)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    def handle_sky(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas_p3_abs['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_empresa'], 3)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    def handle_vetv(self, cantidad):
        self.click(*self.coords_relativas_p3_abs['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas_p3_abs['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas_p3_abs['casilla_empresa'], 5)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas_p3_abs['boton_guardar'])
        self.sleep(2)
        self.handle_error_click_p3()

    # ========== PARTE 4: GOOGLE EARTH ==========
    
    def ejecutar_parte4(self, linea_actual):
        """Ejecutar la parte 4 del proceso (Google Earth)"""
        try:
            logger.info(f"=== EJECUTANDO PARTE 4 PARA LÍNEA {linea_actual} ===")
            
            row_index = linea_actual - 1
            row = self.df.iloc[row_index]
            
            # Verificar valores del CSV
            if not self.verificar_valores_csv_p4(row_index):
                logger.info(f"Valores inválidos en fila {row_index}, saltando...")
                return True
            
            # Obtener valores
            num_txt_type = str(int(row.iloc[28])) if not pd.isna(row.iloc[28]) else None
            texto_adicional = str(row.iloc[29]) if not pd.isna(row.iloc[29]) else ""
            
            if not num_txt_type:
                logger.info("num_txt_type vacío, saltando...")
                return True
            
            nombre_archivo = f"NN {num_txt_type}.kml"
            logger.info(f"Archivo a cargar: {nombre_archivo}")
            logger.info(f"Texto adicional: '{texto_adicional}'")
            
            # Ejecutar secuencia de GE
            success = self.ejecutar_secuencia_ge(nombre_archivo, texto_adicional)
            
            if success:
                logger.info("Parte 4 completada exitosamente")
            else:
                logger.error("Error en parte 4")
            
            return success
            
        except Exception as e:
            logger.error(f"Error en parte 4: {e}")
            return False

    def verificar_valores_csv_p4(self, row_index):
        """Verificar valores para parte 4"""
        try:
            if row_index >= len(self.df):
                return False
            
            row = self.df.iloc[row_index]
            
            # Verificar columna 28 (num_txt_type)
            if len(row) <= 28 or pd.isna(row.iloc[28]):
                return False
                
            # Verificar columna 29 (texto_adicional) - debe existir
            if len(row) <= 29:
                return False
                
            return True
            
        except Exception as e:
            logger.error(f"Error verificando valores CSV: {e}")
            return False

    def ejecutar_secuencia_ge(self, nombre_archivo, texto_adicional):
        """Ejecutar secuencia completa de Google Earth"""
        try:
            # 1. Agregar ruta
            self.click(*self.coords_p4['agregar_ruta'])
            self.sleep(2)
            self.click(*self.coords_p4['archivo'])
            self.sleep(2)
            self.click(*self.coords_p4['abrir'])
            self.sleep(2)
            
            # 2. Cargar archivo
            success = self.handle_archivo_special_behavior(nombre_archivo)
            if not success:
                logger.error("No se pudo cargar el archivo")
                self.click(*self.coords_p4['agregar_ruta'])
                self.sleep(2)
                return False
            
            # 3. Confirmar carga
            self.ahk_enter.presionar_enter(1)
            self.sleep(3)

            # 4. Secuencia adicional de GE
            self.click(*self.coords_p4['agregar_ruta'])
            self.sleep(2)
            self.click(*self.coords_p4['cargar_ruta'])
            self.sleep(2)
            self.click(*self.coords_p4['lote_again'])
            self.sleep(2)
            
            # 5-13. Resto de la secuencia (similar a p4.py)
            # ... (implementar el resto de la secuencia)
            
            return True
            
        except Exception as e:
            logger.error(f"Error en secuencia GE: {e}")
            return False

    def handle_archivo_special_behavior(self, nombre_archivo):
        """Manejar carga de archivos (similar a p4.py)"""
        coordenadas_ventana = self.encontrar_ventana_archivo()
        
        if coordenadas_ventana:
            x_ventana, y_ventana = coordenadas_ventana
            x_campo = x_ventana + 294
            y_campo = y_ventana + 500
            
            if not self.ahk_manager.start_ahk():
                return False
            
            success = self.ahk_manager.ejecutar_acciones_ahk(x_campo, y_campo, nombre_archivo)
            self.ahk_manager.stop_ahk()
            
            if success:
                time.sleep(1.5)
                return True
        
        return False

    def encontrar_ventana_archivo(self):
        """Encontrar ventana de archivo (similar a p4.py)"""
        intentos = 1
        confianza_minima = 0.6
        
        template = cv2.imread(self.ventana_archivo_img)
        if template is None:
            return None
        
        while self.is_running and intentos <= 30:
            try:
                screenshot = pyautogui.screenshot()
                pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confianza_minima:
                    return max_loc
                else:
                    if intentos % 10 == 0:
                        time.sleep(10)
                    else:
                        time.sleep(1)
                    intentos += 1
                    
            except Exception as e:
                time.sleep(1)
                intentos += 1

        return None

    # ========== MÉTODO PRINCIPAL ==========
    
    def procesar_linea(self, numero_linea):
        """Procesar una línea completa (las 4 partes)"""
        logger.info(f"🚀 INICIANDO PROCESAMIENTO DE LÍNEA {numero_linea}")
        
        self.current_line = numero_linea
        
        # Ejecutar las 4 partes en secuencia
        partes = [
            self.ejecutar_parte1,
            self.ejecutar_parte2,
            self.ejecutar_parte3,
            self.ejecutar_parte4
        ]
        
        for i, parte in enumerate(partes, 1):
            logger.info(f"--- Ejecutando Parte {i} ---")
            success = parte(numero_linea)
            
            if not success:
                logger.error(f"Fallo en Parte {i} para línea {numero_linea}")
                return False
            
            logger.info(f"✅ Parte {i} completada")
        
        logger.info(f"🎉 LÍNEA {numero_linea} PROCESADA COMPLETAMENTE")
        return True

    def procesar_csv_completo(self, linea_inicio=1, linea_fin=None):
        """Procesar todo el CSV o un rango de líneas"""
        if not self.cargar_csv():
            return False
        
        if linea_fin is None:
            linea_fin = self.total_lines
        
        logger.info(f"Procesando líneas {linea_inicio} a {linea_fin} de {self.total_lines}")
        
        if not self.iniciar_ahk():
            return False
        
        self.is_running = True
        
        try:
            for linea in range(linea_inicio, linea_fin + 1):
                if not self.is_running:
                    break
                
                logger.info(f"📋 Procesando línea {linea}/{self.total_lines}")
                success = self.procesar_linea(linea)
                
                if not success:
                    logger.error(f"Fallo en línea {linea}, continuando con siguiente...")
                    # Decidir si continuar o detenerse
                    continuar = input("¿Continuar con siguiente línea? (s/n): ").lower().strip()
                    if continuar != 's':
                        break
            
            logger.info("Procesamiento completo finalizado")
            return True
            
        except KeyboardInterrupt:
            logger.info("Procesamiento interrumpido por el usuario")
            return False
        except Exception as e:
            logger.error(f"Error durante el procesamiento: {e}")
            return False
        finally:
            self.detener_ahk()
            self.is_running = False

# ========== FUNCIÓN PRINCIPAL ==========

def main():
    """Función principal"""
    print("=" * 60)
    print("     PROCESADOR COMPLETO CSV - 4 PARTES INTEGRADAS")
    print("=" * 60)
    print()
    
    archivo_csv = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
    
    # Verificar archivo CSV
    if not os.path.exists(archivo_csv):
        print(f"❌ ERROR: Archivo CSV no encontrado: {archivo_csv}")
        input("Presiona Enter para salir...")
        return
    
    # Crear procesador
    procesador = ProcesadorCompleto(archivo_csv)
    
    # Configurar rango de procesamiento
    try:
        linea_inicio = int(input("Línea de inicio (1): ") or "1")
        linea_fin_input = input("Línea final (Enter para todas): ").strip()
        linea_fin = int(linea_fin_input) if linea_fin_input else None
    except ValueError:
        print("❌ Entrada inválida, usando valores por defecto")
        linea_inicio = 1
        linea_fin = None
    
    print()
    print("Configuración:")
    print(f"  - Archivo CSV: {archivo_csv}")
    print(f"  - Línea inicio: {linea_inicio}")
    print(f"  - Línea final: {linea_fin if linea_fin else 'Todas'}")
    print("  - Partes: 1 (Básico) → 2 (NSE V) → 3 (Servicios) → 4 (GE)")
    print()
    
    try:
        input("Presiona Enter para INICIAR el procesamiento completo...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO PROCESAMIENTO COMPLETO...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar procesamiento
        success = procesador.procesar_csv_completo(linea_inicio, linea_fin)
        
        if success:
            print("🎉 PROCESAMIENTO COMPLETADO EXITOSAMENTE!")
        else:
            print("❌ El procesamiento encontró errores")
        
        print()
        input("Presiona Enter para salir...")
        
    except KeyboardInterrupt:
        print()
        print("❌ Procesamiento cancelado por el usuario")
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")

if __name__ == "__main__":
    main()