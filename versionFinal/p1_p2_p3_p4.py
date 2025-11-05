import os
import time
import sys
import pandas as pd
import pyautogui
import cv2
import numpy as np
import logging
import threading
from PIL import ImageGrab
from utils.ahk_managerCopyDelete import AHKManagerCD
from utils.ahk_writer import AHKWriter
from utils.ahk_click_down import AHKClickDown
from utils.ahk_enter import EnterAHKManager
from utils.ahk_manager import AHKManager

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('nse_automation.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Variable global para almacenar la línea a procesar
LINEA_A_PROCESAR = None

class ProcesadorCSV:
    def __init__(self, archivo_csv):
        self.archivo_csv = archivo_csv
        self.df = None
        self.ahk_manager = AHKManagerCD()
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        
    def cargar_csv(self):
        """Carga el archivo CSV"""
        try:
            self.df = pd.read_csv(self.archivo_csv)
            logger.info(f"CSV cargado correctamente: {len(self.df)} registros")
            return True
        except Exception as e:
            logger.error(f"Error cargando CSV: {e}")
            return False
    
    def iniciar_ahk(self):
        """Inicia todos los procesos AHK"""
        logger.info("Iniciando procesos AHK...")
        return (self.ahk_manager.start_ahk() and 
                self.ahk_writer.start_ahk() and 
                self.ahk_click_down.start_ahk())
    
    def detener_ahk(self):
        """Detiene todos los procesos AHK"""
        logger.info("Deteniendo procesos AHK...")
        self.ahk_manager.stop_ahk()
        self.ahk_writer.stop_ahk()
        self.ahk_click_down.stop_ahk()
    
    def buscar_por_id(self, id_buscar):
        """Busca un ID en la primera columna del CSV"""
        if self.df is None:
            logger.error("CSV no cargado")
            return None
            
        # Buscar en la primera columna (asumimos que es la columna 0)
        resultado = self.df[self.df.iloc[:, 0] == id_buscar]
        
        if len(resultado) == 0:
            logger.warning(f"ID {id_buscar} no encontrado en el CSV")
            return None
        
        logger.info(f"ID {id_buscar} encontrado, datos: {resultado.iloc[0].tolist()}")
        return resultado.iloc[0]
    
    def procesar_registro(self):
        """Ejecuta el flujo completo para un registro"""
        try:
            # Paso 2: Click en (89, 263)
            logger.info("Paso 2: Click en (89, 263)")
            pyautogui.click(89, 263)
            time.sleep(1)
            
            # Paso 3: Usar AHKManager en (1483, 519) para obtener ID
            logger.info("Paso 3: Obteniendo ID con AHKManager en (1483, 519)")
            id_obtenido = self.ahk_manager.ejecutar_acciones_ahk(1483, 519)
            
            if not id_obtenido:
                logger.error("No se pudo obtener el ID")
                return False, None
                
            id_obtenido = int(id_obtenido)
            logger.info(f"ID obtenido: {id_obtenido}")
            
            # Paso 4: Buscar el ID en el CSV
            logger.info(f"Paso 4: Buscando ID {id_obtenido} en CSV")
            registro = self.buscar_por_id(id_obtenido)
            
            if registro is None:
                logger.error(f"ID {id_obtenido} no encontrado en CSV")
                return False, None
            
            # Obtener el número de línea del registro encontrado
            linea_procesada = None
            for idx in range(len(self.df)):
                if self.df.iloc[idx, 0] == id_obtenido:
                    linea_procesada = idx + 1  # +1 porque las líneas empiezan en 1
                    break
            
            # Paso 5: Escribir valor de columna 2 en (1483, 519)
            if len(registro) >= 2:  # Verificar que existe columna 2
                valor_columna_2 = str(registro.iloc[1])
                logger.info(f"Paso 5: Escribiendo valor '{valor_columna_2}' en (1483, 519)")
                
                exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, valor_columna_2)
                if not exito_escritura:
                    logger.error("Error en la escritura")
                    return False, linea_procesada
            else:
                logger.warning("No hay columna 2 en el registro")
            
            # Paso 6: Revisar si columna 4 es mayor a 0
            if len(registro) >= 4:  # Verificar que existe columna 4
                valor_columna_4 = registro.iloc[3]
                logger.info(f"Paso 6: Valor columna 4 = {valor_columna_4}")
                
                # Paso 7: Si es mayor a 0, usar AHKClickDown
                if pd.notna(valor_columna_4) and float(valor_columna_4) > 0:
                    veces_down = int(float(valor_columna_4))
                    logger.info(f"Paso 7: Ejecutando {veces_down} veces DOWN en (1507, 636)")
                    
                    exito_down = self.ahk_click_down.ejecutar_click_down(1507, 636, veces_down)
                    if not exito_down:
                        logger.error("Error en click + down")
                        return False, linea_procesada
                else:
                    logger.info("Paso 7: Saltado (columna 4 <= 0)")
            else:
                logger.warning("No hay columna 4 en el registro")
            
            # Paso 8: Click en (1290, 349)
            logger.info("Paso 8: Click en (1290, 349)")
            pyautogui.click(1290, 349)
            time.sleep(1)
            
            logger.info("Procesamiento completado exitosamente")
            return True, linea_procesada
            
        except Exception as e:
            logger.error(f"Error en procesar_registro: {e}")
            return False, None
    
    def procesar_todo(self, pausa_entre_registros=2):
        """Procesa múltiples registros (si es necesario)"""
        if not self.cargar_csv():
            return False, None
            
        if not self.iniciar_ahk():
            return False, None
        
        try:
            # Este método procesa un registro por ejecución
            logger.info("Iniciando procesamiento de registro...")
            exito, linea_procesada = self.procesar_registro()
            
            if exito:
                logger.info(f"Procesamiento completado. Línea procesada: {linea_procesada}")
            else:
                logger.error("Procesamiento falló")
                
            return exito, linea_procesada
            
        finally:
            # Siempre detener AHK al finalizar
            self.detener_ahk()

class NSEAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica  # Línea específica a procesar (1-indexed)
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.reference_image = "img/VentanaAsignar.png"
        self.is_running = False
        
        # Inicializar AHKWriter
        self.ahk_writer = AHKWriter()
        
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # COORDENADAS RELATIVAS (de la tabla verde) - AJUSTADAS PARA COLUMNAS 7-17
        # Estas coordenadas serán sumadas a la posición de la imagen detectada
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
        
        # Coordenadas para botones (relativas)
        self.coords_asignar = [446, 281]  # Botón asignar en la ventana
        self.coords_cerrar = [396, 352]   # Botón cerrar

    def click(self, x, y, duration=0.1):
        """Hacer clic en coordenadas específicas"""
        pyautogui.click(x, y, duration=duration)
        time.sleep(0.5)

    def write_with_ahk(self, x, y, text):
        """Escribir texto usando AHKWriter"""
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, text)
        if not success:
            print(f"❌ Error al escribir con AHK en ({x}, {y}): {text}")
        return success

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

    def detect_image_with_cv2(self, image_path, confidence=0.6):
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
            
            print(f"Imagen encontrada con confianza: {max_val:.2f}")
            return True, max_loc  # Devuelve las coordenadas (x, y) de la esquina superior izquierda
        except Exception as e:
            print(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.6):
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

    def should_skip_process(self, row):
        """Determina si se debe saltar el proceso basado en la columna 6"""
        # Columna 6 es el índice 5 en base 0
        if pd.notna(row[5]):
            col_value = str(row[5]).strip()
            # Si la columna 6 tiene algún valor (no vacío y no NaN), se salta el proceso
            if col_value and col_value != "" and col_value != "nan":
                return True
        return False

    def execute_nse_script(self):
        """Función principal de ejecución NSE - Proceso único"""
        # Iniciar AHKWriter
        if not self.ahk_writer.start_ahk():
            print("❌ No se pudo iniciar AHKWriter")
            return False
            
        try:
            # Leer CSV
            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            # Validar línea específica
            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False
            
            # Procesar solo la línea específica (start_count - 1)
            row = df.iloc[self.linea_especifica - 1]
            print(f"🔄 Procesando línea {self.linea_especifica}/{total_lines}")
            
            # Verificar si se debe saltar el proceso (columna 6 tiene valor)
            if self.should_skip_process(row):
                print(f"⏭️  Saltando línea {self.linea_especifica} - Columna 6 tiene valor: {row[5]}")
                return True
            
            # Verificar que sea tipo V
            if str(row[4]).strip().upper() != "V":
                print(f"⚠️  Saltando línea {self.linea_especifica} - No es tipo V: {row[4]}")
                return True
            
            # click en el boton seleccionar lote 
            self.click(169, 189)
            self.sleep(2)
            # click en el boton asignar nse
            self.click(1491, 386)
            self.sleep(2)
            
            # ESPACIO PARA DETECCIÓN DE IMAGEN CON REINTENTOS
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=30)
            
            if not image_found:
                print("❌ No se puede continuar sin detectar la imagen de referencia.")
                return False
            
            # Si se encontró la imagen, continuar con el proceso usando las coordenadas base
            print("🎯 Imagen detectada, procediendo con tipo V")
            self.handle_type_v(row, base_location)
            
            print(f"✅ Línea {self.linea_especifica} completada (hasta CERRAR)")
            return True
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            return False
        finally:
            # Detener AHKWriter
            self.ahk_writer.stop_ahk()

    def handle_type_v(self, row, base_location):
        """Manejar tipo V con coordenadas relativas - COLUMNAS 7-17"""
        # Calcular coordenadas absolutas sumando las relativas a la posición base
        base_x, base_y = base_location
        
        # Lógica V para columnas 7-17 con coordenadas relativas
        # Nota: row[6] a row[16] corresponden a columnas 7-17 (índices base 0)
        for col_index in range(7, 18):  # 7 a 17 inclusive
            if pd.notna(row[col_index-1]) and row[col_index-1] > 0:
                # Usar coordenadas relativas de la tabla verde, sumando a la base
                x_cs_rel, y_cs_rel = self.coords_select[col_index]
                x_ct_rel, y_ct_rel = self.coords_type[col_index]
                
                # Calcular coordenadas absolutas
                x_cs_abs = base_x + x_cs_rel
                y_cs_abs = base_y + y_cs_rel
                x_ct_abs = base_x + x_ct_rel
                y_ct_abs = base_y + y_ct_rel
                
                self.click(x_cs_abs, y_cs_abs)
                self.sleep(2)
                
                # Usar AHKWriter para escribir en lugar de pyautogui
                texto = str(int(row[col_index-1]))
                self.write_with_ahk(x_ct_abs, y_ct_abs, texto)
                self.sleep(2)
        
        # Botón ASIGNAR antes de cerrar (coordenadas absolutas)
        x_asignar_rel, y_asignar_rel = self.coords_asignar
        x_asignar_abs = base_x + x_asignar_rel
        y_asignar_abs = base_y + y_asignar_rel
        self.click(x_asignar_abs, y_asignar_abs)
        self.sleep(2)
        
        # Botón CERRAR (coordenadas absolutas)
        x_cerrar_rel, y_cerrar_rel = self.coords_cerrar
        x_cerrar_abs = base_x + x_cerrar_rel
        y_cerrar_abs = base_y + y_cerrar_rel
        self.click(x_cerrar_abs, y_cerrar_abs)
        self.sleep(2)

class NSEServicesAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica  # Línea específica a procesar (empezando desde 1)
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.current_line = 0
        self.is_running = False
        self.reference_point = None  # Punto de referencia para coordenadas relativas
        
        # Inicializar controladores AHK
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        self.ahk_enter = EnterAHKManager()
        
        # Configurar coordenadas base (serán actualizadas con coordenadas relativas)
        self.coords = {
            'menu_principal': (81, 81),
            'campo_cantidad': (108, 350),
            'boton_guardar': (63, 390),
            'boton_error': (704, 384),  # Esta no cambia ya que es global
            'cierre': (863, 16),
            'inicio_servicios': (1563, 385),  # Esta no cambia ya que es para iniciar
            'casilla_servicio': (121, 236),
            'casilla_tipo': (121, 261),
            'casilla_empresa': (121, 290),
            'casilla_producto': (121, 322),
        }

    def buscar_imagen(self, imagen_path, timeout=30, confidence=0.8):
        """
        Busca una imagen en la pantalla usando OpenCV
        Retorna las coordenadas de la esquina superior izquierda si la encuentra
        """
        logging.info(f"🔍 Buscando imagen: {imagen_path}")
        
        try:
            # Cargar la imagen template
            template = cv2.imread(imagen_path)
            if template is None:
                logging.error(f"❌ No se pudo cargar la imagen: {imagen_path}")
                return None
            
            template_height, template_width = template.shape[:2]
            
            for intento in range(timeout):
                # Capturar screenshot de toda la pantalla
                screenshot = ImageGrab.grab()
                screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                # Realizar la búsqueda de la plantilla
                result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confidence:
                    # Encontrado - retornar coordenadas de la esquina superior izquierda
                    x, y = max_loc
                    logging.info(f"✅ Imagen encontrada en intento {intento + 1} - Coordenadas: ({x}, {y}) - Confianza: {max_val:.2f}")
                    return (x, y)
                
                logging.info(f"⏳ Intento {intento + 1}/{timeout} - Confianza máxima: {max_val:.2f}")
                time.sleep(1)
            
            logging.error(f"❌ No se encontró la imagen después de {timeout} intentos")
            return None
            
        except Exception as e:
            logging.error(f"❌ Error en búsqueda de imagen: {e}")
            return None

    def actualizar_coordenadas_relativas(self, referencia):
        """
        Actualiza todas las coordenadas para que sean relativas al punto de referencia
        """
        if referencia is None:
            logging.error("❌ No se puede actualizar coordenadas: referencia es None")
            return False
        
        ref_x, ref_y = referencia
        
        # Actualizar coordenadas relativas
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
        
        # Mantener coordenadas que no cambian
        self.coords_relativas['boton_error'] = self.coords['boton_error']
        self.coords_relativas['inicio_servicios'] = self.coords['inicio_servicios']
        
        self.reference_point = referencia
        logging.info("✅ Coordenadas actualizadas a relativas")
        return True

    def iniciar_ahk(self):
        """Iniciar todos los servicios AHK"""
        try:
            if not self.ahk_writer.start_ahk():
                logging.error("No se pudo iniciar AHK Writer")
                return False
            if not self.ahk_click_down.start_ahk():
                logging.error("No se pudo iniciar AHK Click Down")
                return False
            if not self.ahk_enter.start_ahk():
                logging.error("No se pudo iniciar AHK Enter")
                return False
            logging.info("✅ Todos los servicios AHK iniciados correctamente")
            return True
        except Exception as e:
            logging.error(f"Error iniciando servicios AHK: {e}")
            return False

    def detener_ahk(self):
        """Detener todos los servicios AHK"""
        try:
            self.ahk_writer.stop_ahk()
            self.ahk_click_down.stop_ahk()
            self.ahk_enter.stop_ahk()
            logging.info("✅ Todos los servicios AHK detenidos correctamente")
        except Exception as e:
            logging.error(f"Error deteniendo servicios AHK: {e}")

    def click(self, x, y, duration=0.1):
        """Hacer clic en coordenadas específicas"""
        pyautogui.click(x, y, duration=duration)
        time.sleep(0.5)

    def write(self, text):
        """Escribir texto usando AHK Writer"""
        try:
            # Usar coordenadas relativas si están disponibles
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                campo_coords = self.coords_relativas['campo_cantidad']
            else:
                campo_coords = self.coords['campo_cantidad']
                
            # Primero hacer clic en el campo de cantidad, luego escribir
            if self.click(*campo_coords):
                return self.ahk_writer.ejecutar_escritura_ahk(
                    campo_coords[0],
                    campo_coords[1],
                    str(text)
                )
            return False
        except Exception as e:
            logging.error(f"Error escribiendo texto '{text}': {e}")
            return False

    def press_down(self, x, y, times=1):
        """Presionar flecha down usando AHK"""
        try:
            # Usar coordenadas relativas si están disponibles
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                click_coords = (x, y)
            else:
                click_coords = (x, y)
                
            # Usamos AHK Click Down con las veces especificadas
            return self.ahk_click_down.ejecutar_click_down(click_coords[0], click_coords[1], times)
        except Exception as e:
            logging.error(f"Error presionando DOWN {times} veces: {e}")
            return False

    def press_enter(self):
        """Presionar flecha down usando AHK"""
        try:                
            # Usamos AHK Click Down con las veces especificadas
            return self.ahk_enter.presionar_enter(1)
        except Exception as e:
            logging.error(f"Error presionando enter")
            return False

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

    def handle_error_click(self):
        """Manejar clics de error"""
        for _ in range(5):
            # Usar coordenadas relativas si están disponibles para boton_error
            if hasattr(self, 'coords_relativas') and self.coords_relativas:
                self.click(*self.coords_relativas['boton_error'])
            else:
                self.click(*self.coords['boton_error'])
            self.sleep(2)

    def procesar_linea_especifica(self):
        """Procesar solo una línea específica del CSV"""
        try:
            # Leer CSV
            df = pd.read_csv(self.csv_file, encoding='utf-8')
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            # Validar línea específica
            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False
            
            # Obtener la línea específica (ajustar índice ya que CSV empieza en 0 para datos)
            linea_idx = self.linea_especifica - 1  # Convertir a índice base 0
            self.current_line = self.linea_especifica
            
            print(f"🎯 PROCESANDO LÍNEA ESPECÍFICA: {self.current_line}/{total_lines}")
            
            row = df.iloc[linea_idx]
            
            # Solo procesar servicios si la columna 18 tiene valor > 0
            if pd.notna(row[17]) and row[17] > 0:
                print(f"✅ Línea {self.current_line} tiene servicios para procesar")
                
                self.click(*self.coords['inicio_servicios'])
                self.sleep(2)
                
                # BUSCAR IMAGEN Y ACTUALIZAR COORDENADAS
                print("🔍 Buscando ventana de servicios...")
                referencia = self.buscar_imagen("img/ventanaAdministracion4.PNG", timeout=30)
                
                if referencia is None:
                    print("❌ ERROR: No se pudo encontrar la ventana de servicios")
                    return False
                
                # Actualizar coordenadas relativas
                if not self.actualizar_coordenadas_relativas(referencia):
                    print("❌ ERROR: No se pudieron actualizar las coordenadas relativas")
                    return False
                
                # Continuar con el procesamiento normal usando coordenadas relativas
                self.click(*self.coords_relativas['menu_principal'])
                self.sleep(2)
                    
                # Llamar a funciones de servicios
                servicios_procesados = 0
                
                if pd.notna(row[18]) and row[18] > 0:  # VOZ COBRE TELMEX
                    print(f"  └─ Procesando VOZ COBRE TELMEX: {row[18]}")
                    self.handle_voz_cobre(row[18])
                    servicios_procesados += 1
                    
                if pd.notna(row[19]) and row[19] > 0:  # Datos s/dom
                    print(f"  └─ Procesando DATOS S/DOM: {row[19]}")
                    self.handle_datos_sdom(row[19])
                    servicios_procesados += 1
                    
                if pd.notna(row[20]) and row[20] > 0:  # Datos-cobre-telmex-inf
                    print(f"  └─ Procesando DATOS COBRE TELMEX: {row[20]}")
                    self.handle_datos_cobre_telmex(row[20])
                    servicios_procesados += 1
                    
                if pd.notna(row[21]) and row[21] > 0:  # Datos-fibra-telmex-inf
                    print(f"  └─ Procesando DATOS FIBRA TELMEX: {row[21]}")
                    self.handle_datos_fibra_telmex(row[21])
                    servicios_procesados += 1
                    
                if pd.notna(row[22]) and row[22] > 0:  # TV cable otros
                    print(f"  └─ Procesando TV CABLE OTROS: {row[22]}")
                    self.handle_tv_cable_otros(row[22])
                    servicios_procesados += 1
                    
                if pd.notna(row[23]) and row[23] > 0:  # Dish
                    print(f"  └─ Procesando DISH: {row[23]}")
                    self.handle_dish(row[23])
                    servicios_procesados += 1
                    
                if pd.notna(row[24]) and row[24] > 0:  # TVS
                    print(f"  └─ Procesando TVS: {row[24]}")
                    self.handle_tvs(row[24])
                    servicios_procesados += 1
                    
                if pd.notna(row[25]) and row[25] > 0:  # SKY
                    print(f"  └─ Procesando SKY: {row[25]}")
                    self.handle_sky(row[25])
                    servicios_procesados += 1
                    
                if pd.notna(row[26]) and row[26] > 0:  # VETV
                    print(f"  └─ Procesando VETV: {row[26]}")
                    self.handle_vetv(row[26])
                    servicios_procesados += 1
                
                # Usar coordenadas relativas para el cierre
                self.click(*self.coords_relativas['cierre'])
                self.sleep(5)
                
                print(f"✅ Línea {self.current_line} completada: {servicios_procesados} servicios procesados")
                return True
            else:
                print(f"⏭️  Línea {self.current_line} no tiene servicios para procesar")
                return True  # Consideramos éxito si no hay servicios para procesar
            
        except Exception as e:
            print(f"❌ Error procesando línea {self.current_line}: {e}")
            logging.error(f"Error en procesar_linea_especifica: {e}")
            return False

    def handle_voz_cobre(self, cantidad):
        # Usar coordenadas relativas
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_sdom(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_cobre_telmex(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_producto'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_fibra_telmex(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 1)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_tv_cable_otros(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 4)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_dish(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 1)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_tvs(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 2)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_sky(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 3)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_vetv(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter()
        self.press_down(*self.coords_relativas['casilla_empresa'], 5)
        self.press_enter()
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

class GEAutomation:
    def __init__(self, linea_especifica=None):
        self.linea_especifica = linea_especifica  # Línea específica a procesar
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

            # Usar la línea específica determinada por el Programa 1
            if self.linea_especifica is None:
                print("❌ No se especificó línea a procesar")
                return False
                
            if self.linea_especifica < 1 or self.linea_especifica > total_lines:
                print(f"❌ Línea {self.linea_especifica} fuera de rango (1-{total_lines})")
                return False

            row_index = self.linea_especifica - 1
                
            # Verificar si los valores del CSV son válidos
            if not self.verificar_valores_csv(df, row_index):
                print(f"⚠️  Valores inválidos en fila {row_index}. Línea {self.linea_especifica} saltada.")
                return True  # Consideramos éxito si no hay valores válidos
                    
            print(f"🔄 Procesando línea {self.linea_especifica}/{total_lines}")
            success = self.process_single_iteration(df, self.linea_especifica, total_lines)
                
            if not success:
                print(f"⚠️  Línea {self.linea_especifica} falló")
                return False
                
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

    def process_single_iteration(self, df, linea_especifica, total_lines):
        """Procesar una sola iteración del bucle"""
        # Obtener la fila correspondiente (0-indexed)
        row_index = linea_especifica - 1
        row = df.iloc[row_index]
        
        # Obtener valores del CSV con verificación
        try:
            num_txt_type = str(int(row.iloc[28])) if not pd.isna(row.iloc[28]) else None
            texto_adicional = str(row.iloc[29]) if not pd.isna(row.iloc[29]) else ""
        except (ValueError, IndexError) as e:
            print(f"❌ Error obteniendo valores del CSV: {e}")
            return False

        if not num_txt_type:
            print(f"⚠️  num_txt_type vacío en línea {linea_especifica}, saltando...")
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
            
            print(f"✅ Línea {linea_especifica} completada exitosamente")
            return True
            
        except Exception as e:
            print(f"❌ Error en línea {linea_especifica}: {e}")
            # Intentar cerrar ventana de error en caso de excepción
            self.detectar_ventana_error()
            return False

    def save_progress(self):
        """Guardar progreso con Ctrl + S"""
        print("💾 Guardando progreso...")
        pyautogui.hotkey('ctrl', 's')
        self.sleep(6)

def ejecutar_programa1():
    """Ejecuta el primer programa (ProcesadorCSV) automáticamente"""
    global LINEA_A_PROCESAR
    
    print("=" * 60)
    print("INICIANDO PROGRAMA 1 - PROCESADOR CSV")
    print("=" * 60)
    
    # Configurar pyautogui
    pyautogui.PAUSE = 0.5
    pyautogui.FAILSAFE = True
    
    # Nombre del archivo CSV
    archivo_csv = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
    
    # Crear procesador
    procesador = ProcesadorCSV(archivo_csv)
    
    # Ejecutar procesamiento automáticamente
    print("Iniciando procesamiento automático del Programa 1...")
    print("Asegúrate de que la ventana objetivo esté activa")
    print("Presiona Ctrl+C para cancelar")
    
    try:
        print("Iniciando en 3 segundos...")
        time.sleep(3)  # Tiempo para cambiar a la ventana correcta
        resultado, linea_procesada = procesador.procesar_todo()
        
        if resultado and linea_procesada:
            print(f"✅ Programa 1 completado exitosamente. Línea procesada: {linea_procesada}")
            LINEA_A_PROCESAR = linea_procesada
            return True, linea_procesada
        else:
            print("❌ Programa 1 falló o no se pudo determinar la línea procesada")
            return False, None
    except KeyboardInterrupt:
        print("\n❌ Proceso cancelado por el usuario")
        return False, None
    except Exception as e:
        print(f"❌ Error en Programa 1: {e}")
        return False, None

def ejecutar_programa2(linea_especifica):
    """Ejecuta el segundo programa (NSEAutomation) automáticamente usando la línea del Programa 1"""
    print("\n" + "=" * 60)
    print("INICIANDO PROGRAMA 2 - AUTOMATIZACIÓN NSE")
    print("=" * 60)
    
    # Inicializar automatización con la línea especificada
    nse = NSEAutomation(linea_especifica=linea_especifica)
    nse.is_running = True
    
    # Verificar archivo CSV
    if not os.path.exists(nse.csv_file):
        print(f"❌ ERROR: Archivo CSV no encontrado: {nse.csv_file}")
        return False
    
    print(f"✅ Archivo CSV encontrado: {nse.csv_file}")
    print(f"🎯 Procesando línea: {linea_especifica}")
    
    # Verificar imagen de referencia
    if not os.path.exists(nse.reference_image):
        print(f"⚠️  Advertencia: Imagen de referencia no encontrada: {nse.reference_image}")
        print("   El proceso se detendrá si no puede encontrar la imagen después de 30 intentos")
    else:
        print(f"✅ Imagen de referencia encontrada: {nse.reference_image}")
    
    print()
    print("Configuración:")
    print(f"  - Línea a procesar: {linea_especifica}")
    print(f"  - Archivo CSV: {nse.csv_file}")
    print(f"  - Imagen de referencia: {nse.reference_image}")
    print("  - Usando AHKWriter para escritura")
    print()
    
    try:
        print("Iniciando Programa 2 en 3 segundos...")
        time.sleep(3)
        
        print("🚀 INICIANDO AUTOMATIZACIÓN NSE ...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar script NSE
        resultado = nse.execute_nse_script()
        
        if resultado:
            print("✅ Programa 2 finalizado exitosamente")
        else:
            print("❌ Programa 2 falló")
            
        return resultado
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
        nse.is_running = False
        return False
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        nse.is_running = False
        return False

def ejecutar_programa3(linea_especifica):
    """Ejecuta el tercer programa (NSEServicesAutomation) automáticamente usando la línea del Programa 1"""
    print("\n" + "=" * 60)
    print("INICIANDO PROGRAMA 3 - SERVICIOS NSE")
    print("=" * 60)
    
    # Inicializar automatización con la línea especificada
    nse_services = NSEServicesAutomation(linea_especifica=linea_especifica)
    
    # Verificar archivo CSV
    if not os.path.exists(nse_services.csv_file):
        print(f"❌ ERROR: Archivo CSV no encontrado: {nse_services.csv_file}")
        return False
    
    print(f"✅ Archivo CSV encontrado: {nse_services.csv_file}")
    print(f"🎯 Procesando línea: {linea_especifica}")
    
    # Iniciar servicios AHK
    print("🔄 Iniciando servicios AHK...")
    if not nse_services.iniciar_ahk():
        print("❌ No se pudieron iniciar los servicios AHK")
        return False
    
    try:
        print("Iniciando Programa 3 en 3 segundos...")
        time.sleep(3)
        
        print("🚀 INICIANDO PROCESAMIENTO DE SERVICIOS NSE ...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar procesamiento de línea específica
        nse_services.is_running = True
        resultado = nse_services.procesar_linea_especifica()
        
        if resultado:
            print(f"🎉 LÍNEA {linea_especifica} PROCESADA EXITOSAMENTE!")
        else:
            print(f"❌ HUBO PROBLEMAS PROCESANDO LA LÍNEA {linea_especifica}")
        
        return resultado
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
        return False
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        return False
    finally:
        nse_services.is_running = False
        nse_services.detener_ahk()

def ejecutar_programa4(linea_especifica):
    """Ejecuta el cuarto programa (GEAutomation) automáticamente usando la línea del Programa 1"""
    print("\n" + "=" * 60)
    print("INICIANDO PROGRAMA 4 - AUTOMATIZACIÓN GE")
    print("=" * 60)
    
    # Inicializar automatización con la línea especificada
    ge_auto = GEAutomation(linea_especifica=linea_especifica)
    ge_auto.is_running = True
    
    # Verificar archivo CSV
    if not os.path.exists(ge_auto.csv_file):
        print(f"❌ ERROR: Archivo CSV no encontrado: {ge_auto.csv_file}")
        return False
    
    print(f"✅ Archivo CSV encontrado: {ge_auto.csv_file}")
    print(f"🎯 Procesando línea: {linea_especifica}")
    
    # Verificar que el CSV tiene datos
    try:
        df = pd.read_csv(ge_auto.csv_file)
        if len(df) == 0:
            print("❌ ERROR: El archivo CSV está vacío")
            return False
        print(f"✅ CSV tiene {len(df)} filas de datos")
    except Exception as e:
        print(f"❌ ERROR: No se pudo leer el CSV: {e}")
        return False
    
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
    print(f"  - Línea a procesar: {linea_especifica}")
    print(f"  - Archivo CSV: {ge_auto.csv_file}")
    print(f"  - Imagen de referencia: {ge_auto.reference_image}")
    print(f"  - Imagen ventana archivo: {ge_auto.ventana_archivo_img}")
    print(f"  - Imagen ventana error: {ge_auto.ventana_error_img}")
    print("  - Usando AHK Writer para escritura de texto adicional")
    print("  - Usando AHK Manager para carga de archivos")
    print("  - Usando AHK Enter para presionar Enter")
    print("  - Usando AHK para flechas abajo")
    print()
    
    try:
        print("Iniciando Programa 4 en 3 segundos...")
        time.sleep(3)
        
        print("🚀 INICIANDO AUTOMATIZACIÓN GE ...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar script
        success = ge_auto.perform_actions()
        
        if success:
            print("✅ Programa 4 finalizado exitosamente")
        else:
            print("❌ Programa 4 falló")
            
        return success
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
        ge_auto.is_running = False
        return False
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        ge_auto.is_running = False
        return False

# Función principal combinada
def main():
    """Función principal que ejecuta los cuatro programas secuencialmente sin pausas"""
    global LINEA_A_PROCESAR
    
    print("COMBINACIÓN DE 4 PROGRAMAS - EJECUCIÓN AUTOMÁTICA")
    print("Este script ejecutará los cuatro programas de forma secuencial automáticamente")
    print("El Programa 1 determinará qué línea procesar")
    print("Los Programas 2, 3 y 4 usarán la misma línea automáticamente")
    print("Presiona Ctrl+C para cancelar en cualquier momento")
    
    try:
        # Cuenta regresiva inicial
        print("\nIniciando ejecución automática en:")
        for i in range(5, 0, -1):
            print(f"⏰ {i}...")
            time.sleep(1)
        
        # Ejecutar Programa 1 - Este determina la línea a procesar
        resultado_programa1, linea_procesada = ejecutar_programa1()
        
        # Verificar si el Programa 1 fue exitoso y obtuvo una línea
        if resultado_programa1 and linea_procesada:
            print(f"\n📋 Línea determinada por Programa 1: {linea_procesada}")
            
            # Ejecutar Programas 2, 3 y 4 con la línea del Programa 1
            programas = [
                ("PROGRAMA 2", ejecutar_programa2),
                ("PROGRAMA 3", ejecutar_programa3), 
                ("PROGRAMA 4", ejecutar_programa4)
            ]
            
            resultados = {}
            for nombre_programa, funcion_ejecutar in programas:
                # Pequeña pausa entre programas
                print("\n" + "=" * 60)
                print(f"TRANSICIÓN AL {nombre_programa}")
                print("=" * 60)
                print(f"Iniciando {nombre_programa} en 3 segundos...")
                time.sleep(3)
                
                # Ejecutar programa con la línea del Programa 1
                resultado = funcion_ejecutar(linea_procesada)
                resultados[nombre_programa] = resultado
            
            resultado_programa2 = resultados["PROGRAMA 2"]
            resultado_programa3 = resultados["PROGRAMA 3"] 
            resultado_programa4 = resultados["PROGRAMA 4"]
        else:
            print("❌ Programa 1 falló o no determinó una línea, saltando Programas 2, 3 y 4")
            resultado_programa2 = False
            resultado_programa3 = False
            resultado_programa4 = False
        
        # Resultado final
        print("\n" + "=" * 60)
        print("RESULTADO FINAL DE LA EJECUCIÓN COMBINADA")
        print("=" * 60)
        
        resultados = {
            "Programa 1 (Procesador CSV)": resultado_programa1,
            "Programa 2 (Automatización NSE)": resultado_programa2,
            "Programa 3 (Servicios NSE)": resultado_programa3,
            "Programa 4 (Automatización GE)": resultado_programa4
        }
        
        for programa, resultado in resultados.items():
            estado = "✅ EXITOSO" if resultado else "❌ FALLIDO"
            print(f"  {programa}: {estado}")
        
        if all(resultados.values()):
            print(f"\n🎉 ¡TODOS LOS PROGRAMAS SE EJECUTARON EXITOSAMENTE!")
            print(f"📊 Línea procesada: {linea_procesada}")
        else:
            programas_fallidos = [k for k, v in resultados.items() if not v]
            print(f"\n⚠️  PROGRAMAS FALLIDOS: {', '.join(programas_fallidos)}")
        
        print("=" * 60)
        
    except KeyboardInterrupt:
        print("\n❌ Ejecución cancelada por el usuario")
    except Exception as e:
        print(f"\n❌ Error general en la ejecución combinada: {e}")
    finally:
        # Pausa final breve para que el usuario pueda ver los resultados
        print("\nEjecución finalizada. El programa se cerrará en 10 segundos...")
        time.sleep(10)

if __name__ == "__main__":
    main()