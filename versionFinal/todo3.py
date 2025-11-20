import pandas as pd
import pyautogui
import time
import logging
import os
import cv2
import numpy as np
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
        logging.FileHandler('automation_completa.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class AutomatizacionCompleta:
    def __init__(self, archivo_csv="NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"):
        self.archivo_csv = archivo_csv
        self.df = None
        self.linea_procesada = None
        self.id_obtenido = None
        self.is_running = False
        
        # Inicializar todos los manejadores AHK
        self.ahk_manager = AHKManagerCD()
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        self.ahk_enter = EnterAHKManager()
        self.ahk_manager_ge = AHKManager()
        
        # Configurar pyautogui
        pyautogui.PAUSE = 0.5
        pyautogui.FAILSAFE = True
        
        # Configuraciones específicas para cada parte
        self.configurar_p2()
        self.configurar_p3()
        self.configurar_p4()

    def configurar_p2(self):
        """Configuración para la parte NSE (p2)"""
        self.reference_image_p2 = "img/VentanaAsignar.png"
        self.coords_select_p2 = {
            7: [33, 92], 8: [33, 131], 9: [33, 159], 10: [33, 197],
            11: [33, 231], 12: [398, 92], 13: [398, 131], 14: [398, 159],
            15: [33, 301], 16: [33, 333], 17: [33, 367]
        }
        self.coords_type_p2 = {
            7: [163, 92], 8: [163, 131], 9: [163, 159], 10: [163, 197],
            11: [163, 231], 12: [528, 92], 13: [528, 131], 14: [528, 159],
            15: [163, 301], 16: [163, 333], 17: [163, 367]
        }
        self.coords_asignar_p2 = [446, 281]
        self.coords_cerrar_p2 = [396, 352]

    def configurar_p3(self):
        """Configuración para la parte Servicios (p3)"""
        self.coords_p3 = {
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
        self.coords_relativas_p3 = {}

    def configurar_p4(self):
        """Configuración para la parte GE (p4)"""
        self.reference_image_p4 = "img/textoAdicional.PNG"
        self.ventana_archivo_img_p4 = "img/cargarArchivo.png"
        self.ventana_error_img_p4 = "img/ventanaError.png"
        self.coords_p4 = {
            'agregar_ruta': (327, 381),
            'archivo': (1396, 608),
            'abrir': (1406, 634),
            'seleccionar_mapa': (168, 188),
            'anotar': (1366, 384),
            'agregar_texto_adicional': (1449, 452),
            'limpiar_trazo': (360, 980),
            'lote_again': (70, 266)
        }
        self.coords_texto_relativas_p4 = {
            'campo_texto': (230, 66),
            'agregar_texto': (64, 100),
            'cerrar_ventana_texto': (139, 98)
        }

    def iniciar_todos_ahk(self):
        """Inicia todos los procesos AHK"""
        logger.info("Iniciando todos los procesos AHK...")
        return (self.ahk_manager.start_ahk() and 
                self.ahk_writer.start_ahk() and 
                self.ahk_click_down.start_ahk() and
                self.ahk_enter.start_ahk() and
                self.ahk_manager_ge.start_ahk())

    def detener_todos_ahk(self):
        """Detiene todos los procesos AHK"""
        logger.info("Deteniendo todos los procesos AHK...")
        self.ahk_manager.stop_ahk()
        self.ahk_writer.stop_ahk()
        self.ahk_click_down.stop_ahk()
        self.ahk_enter.stop_ahk()
        self.ahk_manager_ge.stop_ahk()

    def cargar_csv(self):
        """Carga el archivo CSV"""
        try:
            self.df = pd.read_csv(self.archivo_csv)
            logger.info(f"CSV cargado correctamente: {len(self.df)} líneas")
            return True
        except Exception as e:
            logger.error(f"Error cargando CSV: {e}")
            return False

    # ===== MÉTODOS DE p1.py =====
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

    def procesar_p1(self):
        """Ejecuta el flujo completo de p1.py para obtener la línea específica"""
        try:
            if not self.cargar_csv():
                return False

            # Paso 2: Click en (89, 263)
            logger.info("Paso 2: Click en (89, 263)")
            pyautogui.click(89, 263)
            time.sleep(1)
            
            # Paso 3: Usar AHKManager en (1483, 519) para obtener ID
            logger.info("Paso 3: Obteniendo ID con AHKManager en (1483, 519)")
            id_obtenido = self.ahk_manager.ejecutar_acciones_ahk(1483, 519)
            
            if not id_obtenido:
                logger.error("No se pudo obtener el ID")
                return False
                
            self.id_obtenido = int(id_obtenido)
            logger.info(f"ID obtenido: {self.id_obtenido}")
            
            # Paso 4: Buscar el ID en el CSV
            logger.info(f"Paso 4: Buscando ID {self.id_obtenido} en CSV")
            linea_especifica = self.buscar_por_id(self.id_obtenido)
            
            if linea_especifica is None:
                logger.error(f"ID {self.id_obtenido} no encontrado en CSV")
                return False
            
            self.linea_procesada = linea_especifica
            logger.info(f"Línea específica obtenida: índice {linea_especifica.name}")
            
            # Paso 5: Escribir valor de columna 2 en (1483, 519)
            if len(linea_especifica) >= 2:
                valor_columna_2 = str(linea_especifica.iloc[1])
                logger.info(f"Paso 5: Escribiendo valor '{valor_columna_2}' en (1483, 519)")
                
                exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, valor_columna_2)
                if not exito_escritura:
                    logger.error("Error en la escritura")
                    return False
            
            # Paso 6: Revisar si columna 4 es mayor a 0
            if len(linea_especifica) >= 4:
                valor_columna_4 = linea_especifica.iloc[3]
                logger.info(f"Paso 6: Valor columna 4 = {valor_columna_4}")
                
                if pd.notna(valor_columna_4) and float(valor_columna_4) > 0:
                    veces_down = int(float(valor_columna_4))
                    logger.info(f"Paso 7: Ejecutando {veces_down} veces DOWN en (1507, 636)")
                    
                    exito_down = self.ahk_click_down.ejecutar_click_down(1507, 636, veces_down)
                    if not exito_down:
                        logger.error("Error en click + down")
                        return False
            
            # Paso 8: Click en (1290, 349)
            logger.info("Paso 8: Click en (1290, 349)")
            pyautogui.click(1290, 349)
            time.sleep(1)
            
            logger.info("Procesamiento P1 completado exitosamente")
            return True
            
        except Exception as e:
            logger.error(f"Error en P1: {e}")
            return False

    # ===== MÉTODOS DE p2.py =====
    def detect_image_with_cv2(self, image_path, confidence=0.6):
        """Detectar imagen en pantalla usando template matching con OpenCV"""
        try:
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            template = cv2.imread(image_path)
            if template is None:
                print(f"Error: No se pudo cargar la imagen {image_path}")
                return False, None
            
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            if max_val < confidence:
                return False, None
            
            return True, max_loc
        except Exception as e:
            print(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.6):
        """Esperar a que aparezca una imagen con múltiples intentos"""
        print(f"🔍 Buscando imagen: {image_path}")
        
        for attempt in range(1, max_attempts + 1):
            found, location = self.detect_image_with_cv2(image_path, confidence)
            
            if found:
                print(f"✅ Imagen detectada en el intento {attempt}")
                return True, location
            
            print(f"⏳ Intento {attempt}/{max_attempts} - Imagen no encontrada")
            
            if attempt < max_attempts:
                if attempt % 10 == 0:
                    time.sleep(10)
                else:
                    time.sleep(2)
        
        print("❌ Imagen no encontrada después de 30 intentos.")
        return False, None

    def should_skip_process_p2(self, row):
        """Determina si se debe saltar el proceso basado en la columna 6"""
        if pd.notna(row[5]):
            col_value = str(row[5]).strip()
            if col_value and col_value != "" and col_value != "nan":
                return True
        return False

    def procesar_p2(self):
        """Ejecuta el flujo completo de p2.py"""
        try:
            if self.linea_procesada is None:
                logger.error("No hay línea procesada de P1")
                return False

            row = self.linea_procesada
            linea_idx = self.linea_procesada.name
            
            print(f"🔄 Procesando P2 para línea {linea_idx + 1}")
            
            # Verificar si se debe saltar el proceso
            if self.should_skip_process_p2(row):
                print(f"⏭️  Saltando línea {linea_idx + 1} - Columna 6 tiene valor")
                return True

            # Verificar que sea tipo V
            if str(row[4]).strip().upper() != "V":
                print(f"⚠️  Saltando línea {linea_idx + 1} - No es tipo V: {row[4]}")
                return True

            # Clic en botones de P2
            pyautogui.click(169, 189)
            time.sleep(2)
            pyautogui.click(1491, 386)
            time.sleep(2)
            
            # Detección de imagen
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image_p2, max_attempts=30)
            
            if not image_found:
                print("❌ No se puede continuar sin detectar la imagen de referencia.")
                return False
            
            # Procesar tipo V
            self.handle_type_v_p2(row, base_location)
            
            print(f"✅ P2 completado para línea {linea_idx + 1}")
            return True
            
        except Exception as e:
            logger.error(f"Error en P2: {e}")
            return False

    def handle_type_v_p2(self, row, base_location):
        """Manejar tipo V con coordenadas relativas - COLUMNAS 7-17"""
        base_x, base_y = base_location
        
        for col_index in range(7, 18):
            if pd.notna(row[col_index-1]) and row[col_index-1] > 0:
                x_cs_rel, y_cs_rel = self.coords_select_p2[col_index]
                x_ct_rel, y_ct_rel = self.coords_type_p2[col_index]
                
                x_cs_abs = base_x + x_cs_rel
                y_cs_abs = base_y + y_cs_rel
                x_ct_abs = base_x + x_ct_rel
                y_ct_abs = base_y + y_ct_rel
                
                pyautogui.click(x_cs_abs, y_cs_abs)
                time.sleep(2)
                
                texto = str(int(row[col_index-1]))
                self.ahk_writer.ejecutar_escritura_ahk(x_ct_abs, y_ct_abs, texto)
                time.sleep(2)
        
        # Botones finales
        x_asignar_abs = base_x + self.coords_asignar_p2[0]
        y_asignar_abs = base_y + self.coords_asignar_p2[1]
        pyautogui.click(x_asignar_abs, y_asignar_abs)
        time.sleep(2)
        
        x_cerrar_abs = base_x + self.coords_cerrar_p2[0]
        y_cerrar_abs = base_y + self.coords_cerrar_p2[1]
        pyautogui.click(x_cerrar_abs, y_cerrar_abs)
        time.sleep(2)

    # ===== MÉTODOS DE p3.py =====
    def buscar_imagen_p3(self, imagen_path, timeout=30, confidence=0.8):
        """Busca una imagen en la pantalla usando OpenCV"""
        try:
            template = cv2.imread(imagen_path)
            if template is None:
                logger.error(f"❌ No se pudo cargar la imagen: {imagen_path}")
                return None
            
            for intento in range(timeout):
                screenshot = ImageGrab.grab()
                screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confidence:
                    x, y = max_loc
                    logger.info(f"✅ Imagen encontrada en intento {intento + 1}")
                    return (x, y)
                
                time.sleep(1)
            
            return None
        except Exception as e:
            logger.error(f"❌ Error en búsqueda de imagen: {e}")
            return None

    def actualizar_coordenadas_relativas_p3(self, referencia):
        """Actualiza coordenadas relativas para P3"""
        if referencia is None:
            return False
        
        ref_x, ref_y = referencia
        
        self.coords_relativas_p3 = {
            'menu_principal': (ref_x + 81, ref_y + 81),
            'campo_cantidad': (ref_x + 108, ref_y + 350),
            'boton_guardar': (ref_x + 63, ref_y + 390),
            'cierre': (ref_x + 863, ref_y + 16),
            'casilla_servicio': (ref_x + 121, ref_y + 236),
            'casilla_tipo': (ref_x + 121, ref_y + 261),
            'casilla_empresa': (ref_x + 121, ref_y + 290),
            'casilla_producto': (ref_x + 121, ref_y + 322),
        }
        
        self.coords_relativas_p3['boton_error'] = self.coords_p3['boton_error']
        self.coords_relativas_p3['inicio_servicios'] = self.coords_p3['inicio_servicios']
        
        return True

    def procesar_p3(self):
        """Ejecuta el flujo completo de p3.py"""
        try:
            if self.linea_procesada is None:
                logger.error("No hay línea procesada de P1")
                return False

            row = self.linea_procesada
            linea_idx = self.linea_procesada.name
            
            print(f"🎯 PROCESANDO P3 PARA LÍNEA: {linea_idx + 1}")
            
            # Solo procesar servicios si la columna 18 tiene valor > 0
            if pd.notna(row[17]) and row[17] > 0:
                print(f"✅ Línea {linea_idx + 1} tiene servicios para procesar")
                
                pyautogui.click(*self.coords_p3['inicio_servicios'])
                time.sleep(2)
                
                # Buscar imagen y actualizar coordenadas
                referencia = self.buscar_imagen_p3("img/ventanaAdministracion4.PNG", timeout=30)
                
                if referencia is None:
                    print("❌ ERROR: No se pudo encontrar la ventana de servicios")
                    return False
                
                if not self.actualizar_coordenadas_relativas_p3(referencia):
                    print("❌ ERROR: No se pudieron actualizar las coordenadas relativas")
                    return False
                
                # Continuar con el procesamiento
                pyautogui.click(*self.coords_relativas_p3['menu_principal'])
                time.sleep(2)
                    
                # Procesar servicios
                servicios_procesados = self.procesar_servicios_p3(row)
                
                # Cierre
                pyautogui.click(*self.coords_relativas_p3['cierre'])
                time.sleep(5)
                
                print(f"✅ P3 completado: {servicios_procesados} servicios procesados")
                return True
            else:
                print(f"⏭️  Línea {linea_idx + 1} no tiene servicios para procesar")
                return True
            
        except Exception as e:
            logger.error(f"Error en P3: {e}")
            return False

    def procesar_servicios_p3(self, row):
        """Procesa todos los servicios de P3"""
        servicios_procesados = 0
        
        # VOZ COBRE TELMEX (columna 19)
        if pd.notna(row[18]) and row[18] > 0:
            print(f"  └─ Procesando VOZ COBRE TELMEX: {row[18]}")
            self.handle_voz_cobre(row[18])
            servicios_procesados += 1
            
        # Datos s/dom (columna 20)
        if pd.notna(row[19]) and row[19] > 0:
            print(f"  └─ Procesando DATOS S/DOM: {row[19]}")
            self.handle_datos_sdom(row[19])
            servicios_procesados += 1
            
        # Datos-cobre-telmex-inf (columna 21)
        if pd.notna(row[20]) and row[20] > 0:
            print(f"  └─ Procesando DATOS COBRE TELMEX: {row[20]}")
            self.handle_datos_cobre_telmex(row[20])
            servicios_procesados += 1
            
        # Datos-fibra-telmex-inf (columna 22)
        if pd.notna(row[21]) and row[21] > 0:
            print(f"  └─ Procesando DATOS FIBRA TELMEX: {row[21]}")
            self.handle_datos_fibra_telmex(row[21])
            servicios_procesados += 1
            
        # TV cable otros (columna 23)
        if pd.notna(row[22]) and row[22] > 0:
            print(f"  └─ Procesando TV CABLE OTROS: {row[22]}")
            self.handle_tv_cable_otros(row[22])
            servicios_procesados += 1
            
        # Dish (columna 24)
        if pd.notna(row[23]) and row[23] > 0:
            print(f"  └─ Procesando DISH: {row[23]}")
            self.handle_dish(row[23])
            servicios_procesados += 1
            
        # TVS (columna 25)
        if pd.notna(row[24]) and row[24] > 0:
            print(f"  └─ Procesando TVS: {row[24]}")
            self.handle_tvs(row[24])
            servicios_procesados += 1
            
        # SKY (columna 26)
        if pd.notna(row[25]) and row[25] > 0:
            print(f"  └─ Procesando SKY: {row[25]}")
            self.handle_sky(row[25])
            servicios_procesados += 1
            
        # VETV (columna 27)
        if pd.notna(row[26]) and row[26] > 0:
            print(f"  └─ Procesando VETV: {row[26]}")
            self.handle_vetv(row[26])
            servicios_procesados += 1
            
        return servicios_procesados

    def handle_voz_cobre(self, cantidad):
        """Manejar servicio Voz Cobre"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_datos_sdom(self, cantidad):
        """Manejar servicio Datos s/dom"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_servicio'], 2)
        self.ahk_enter.presionar_enter(1)
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_datos_cobre_telmex(self, cantidad):
        """Manejar servicio Datos Cobre Telmex"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_servicio'], 2)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_producto'], 1)
        self.ahk_enter.presionar_enter(1)
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_datos_fibra_telmex(self, cantidad):
        """Manejar servicio Datos Fibra Telmex"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_servicio'], 2)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_tipo'], 1)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_empresa'], 1)
        self.ahk_enter.presionar_enter(1)
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_tv_cable_otros(self, cantidad):
        """Manejar servicio TV Cable Otros"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_servicio'], 3)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_empresa'], 4)
        self.ahk_enter.presionar_enter(1)
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_dish(self, cantidad):
        """Manejar servicio Dish"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_servicio'], 3)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_tipo'], 2)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_empresa'], 1)
        self.ahk_enter.presionar_enter(1)
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_tvs(self, cantidad):
        """Manejar servicio TVS"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_servicio'], 3)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_tipo'], 2)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_empresa'], 2)
        self.ahk_enter.presionar_enter(1)
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_sky(self, cantidad):
        """Manejar servicio Sky"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_servicio'], 3)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_tipo'], 2)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_empresa'], 3)
        self.ahk_enter.presionar_enter(1)
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_vetv(self, cantidad):
        """Manejar servicio VETV"""
        pyautogui.click(*self.coords_relativas_p3['menu_principal'])
        time.sleep(2)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_servicio'], 3)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_tipo'], 2)
        self.ahk_enter.presionar_enter(1)
        self.ahk_click_down.ejecutar_click_down(*self.coords_relativas_p3['casilla_empresa'], 5)
        self.ahk_enter.presionar_enter(1)
        time.sleep(2)
        self.ahk_writer.ejecutar_escritura_ahk(
            self.coords_relativas_p3['campo_cantidad'][0],
            self.coords_relativas_p3['campo_cantidad'][1],
            str(int(cantidad))
        )
        time.sleep(2)
        pyautogui.click(*self.coords_relativas_p3['boton_guardar'])
        time.sleep(2)
        self.handle_error_click_p3()

    def handle_error_click_p3(self):
        """Manejar clics de error en P3"""
        for _ in range(5):
            pyautogui.click(*self.coords_relativas_p3['boton_error'])
            time.sleep(2)

    # ===== MÉTODOS DE p4.py =====
    def encontrar_ventana_archivo_p4(self):
        """Busca la ventana de archivo usando template matching con reintentos inteligentes"""
        intentos = 1
        confianza_minima = 0.6
        tiempo_espera_base = 1
        tiempo_espera_largo = 10
        
        template = cv2.imread(self.ventana_archivo_img_p4)
        if template is None:
            logger.error(f"No se pudo cargar la imagen '{self.ventana_archivo_img_p4}'")
            return None
        
        while self.is_running: 
            try:
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

    def detectar_ventana_error_p4(self):
        """Detecta la ventana de error y presiona Enter para cerrarla"""
        try:
            template = cv2.imread(self.ventana_error_img_p4) 
            if template is None:
                logger.error(f"No se pudo cargar la imagen '{self.ventana_error_img_p4}'")
                return False
            
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            confianza_minima = 0.6
            
            if max_val >= confianza_minima:
                logger.info(f"Ventana de error detectada con confianza: {max_val:.2f}")
                
                if not self.ahk_enter.start_ahk():
                    logger.error("No se pudo iniciar AutoHotkey")
                    return False
                    
                if self.ahk_enter.presionar_enter(1):
                    time.sleep(2.5)
                else:
                    logger.error("Error enviando comando a AHK")
                    return False
                
                logger.info("Ventana de error detectada y cerrada")
                self.ahk_enter.stop_ahk()
                return True
            else:
                return False
                
        except Exception as e:
            logger.error(f"Error al detectar ventana de error: {e}")
            return False

    def handle_archivo_special_behavior_p4(self, nombre_archivo):
        """Maneja el comportamiento especial para cargar archivos usando AHK Manager"""
        coordenadas_ventana = self.encontrar_ventana_archivo_p4()

        if coordenadas_ventana:
            x_ventana, y_ventana = coordenadas_ventana
            logger.info(f"Coordenadas ventana: x={x_ventana}, y={y_ventana}")
            
            x_campo = x_ventana + 294
            y_campo = y_ventana + 500
            logger.info(f"Coordenadas campo texto: x={x_campo}, y={y_campo}")
            
            if not self.ahk_manager_ge.start_ahk():
                logger.error("No se pudo iniciar AutoHotkey")
                return False
            
            if self.ahk_manager_ge.ejecutar_acciones_ahk(x_campo, y_campo, nombre_archivo):
                time.sleep(1.5)
            else:
                logger.error("Error enviando comando a AHK")
                return False
            
            self.ahk_manager_ge.stop_ahk()
            return True
        else:
            logger.error("No se pudo encontrar la ventana de archivo.")
            return False

    def escribir_texto_adicional_ahk_p4(self, x, y, texto):
        """Escribe texto adicional usando AHK Writer"""
        if not texto or pd.isna(texto) or str(texto).strip() == '':
            print("⚠️  Texto adicional vacío, saltando escritura")
            return True
            
        print(f"📝 Intentando escribir texto: '{texto}' en coordenadas ({x}, {y})")
        
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
                pyautogui.click(x, y)
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.5)
                pyautogui.press('delete')
                time.sleep(0.5)
                pyautogui.write(texto, interval=0.05)
                print(f"✅ Texto escrito con pyautogui: '{texto}'")
                success = True
            except Exception as e:
                print(f"❌ También falló pyautogui: {e}")
                
        return success

    def wait_for_image_with_retries_p4(self, image_path, max_attempts=30, confidence=0.7):
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

    def procesar_p4(self):
        """Ejecuta el flujo completo de p4.py"""
        try:
            if self.linea_procesada is None:
                logger.error("No hay línea procesada de P1")
                return False

            row = self.linea_procesada
            linea_idx = self.linea_procesada.name
            
            print(f"🔄 Procesando P4 para línea {linea_idx + 1}")
            
            # Obtener valores del CSV
            try:
                num_txt_type = str(int(row.iloc[28])) if not pd.isna(row.iloc[28]) else None
                texto_adicional = str(row.iloc[29]) if not pd.isna(row.iloc[29]) else ""
            except (ValueError, IndexError) as e:
                print(f"❌ Error obteniendo valores del CSV: {e}")
                return False

            if not num_txt_type:
                print(f"⚠️  num_txt_type vacío, saltando...")
                return True

            nombre_archivo = f"NN {num_txt_type}.kml"
            print(f"📁 Archivo a cargar: {nombre_archivo}")
            print(f"📝 Texto adicional: '{texto_adicional}'")

            # SECUENCIA DE ACCIONES P4
            success = self.ejecutar_secuencia_p4(nombre_archivo, texto_adicional)
            
            if success:
                print(f"✅ P4 completado para línea {linea_idx + 1}")
            else:
                print(f"❌ P4 falló para línea {linea_idx + 1}")
            
            return success
            
        except Exception as e:
            logger.error(f"Error en P4: {e}")
            return False

    def ejecutar_secuencia_p4(self, nombre_archivo, texto_adicional):
        """Ejecuta la secuencia de acciones de P4"""
        try:
            # 1. Seleccionar Agregar ruta de GE
            pyautogui.click(*self.coords_p4['agregar_ruta'])
            time.sleep(2)
            pyautogui.click(*self.coords_p4['archivo'])
            time.sleep(2)
            pyautogui.click(*self.coords_p4['abrir'])
            time.sleep(2)
            
            # 2. Cargar archivo
            success = self.handle_archivo_special_behavior_p4(nombre_archivo)
            
            if not success:
                print("❌ No se pudo cargar el archivo. Regresando a agregar_ruta...")
                pyautogui.click(*self.coords_p4['agregar_ruta'])
                time.sleep(2)
                return False
            
            # 3. Presionar Enter para confirmar
            if not self.ahk_enter.presionar_enter(1):
                print("⚠️  No se pudo presionar Enter con AHK, usando pyautogui")
                pyautogui.press('enter')
            
            time.sleep(3)

            pyautogui.click(*self.coords_p4['agregar_ruta'])
            time.sleep(2)

            pyautogui.click(1406, 675)  # cargar ruta
            time.sleep(2)

            pyautogui.click(70, 266)  # seleccionar lote
            time.sleep(2)
            
            # 4. Seleccionar en el mapa
            pyautogui.click(*self.coords_p4['seleccionar_mapa'])
            time.sleep(2)
            
            # 5. Anotar
            pyautogui.click(*self.coords_p4['anotar'])
            time.sleep(2)
            
            # 6. Agregar texto adicional
            pyautogui.click(*self.coords_p4['agregar_texto_adicional'])
            time.sleep(2)
            
            # 7. DETECCIÓN DE IMAGEN para el campo de texto
            image_found, base_location = self.wait_for_image_with_retries_p4(self.reference_image_p4, max_attempts=10)
            
            if image_found:
                # Calcular coordenadas absolutas del campo de texto
                x_campo = base_location[0] + self.coords_texto_relativas_p4['campo_texto'][0]
                y_campo = base_location[1] + self.coords_texto_relativas_p4['campo_texto'][1]
                x_agregar = base_location[0] + self.coords_texto_relativas_p4['agregar_texto'][0]
                y_agregar = base_location[1] + self.coords_texto_relativas_p4['agregar_texto'][1]
                x_cerrar = base_location[0] + self.coords_texto_relativas_p4['cerrar_ventana_texto'][0]
                y_cerrar = base_location[1] + self.coords_texto_relativas_p4['cerrar_ventana_texto'][1]
                
                # Escribir el texto adicional
                if texto_adicional and texto_adicional.strip():
                    writing_success = self.escribir_texto_adicional_ahk_p4(x_campo, y_campo, texto_adicional)
                    if not writing_success:
                        print("⚠️  Falló la escritura con AHK, intentando con pyautogui...")
                        pyautogui.write(texto_adicional, interval=0.05)
                else:
                    print("ℹ️  Texto adicional vacío, no se escribe nada")
                
                time.sleep(2)

                # 8. Agregar texto adicional
                pyautogui.click(x_agregar, y_agregar)
                time.sleep(3)
            
                # 9. Cerrar ventana de texto adicional
                pyautogui.click(x_cerrar, y_cerrar)
                time.sleep(2)
            else:
                print("❌ No se pudo detectar la imagen del campo de texto")
                return False
            
            # 10. Limpiar trazo
            pyautogui.click(*self.coords_p4['limpiar_trazo'])
            time.sleep(1)
            
            # 11. Seleccionar Lote nuevamente
            pyautogui.click(*self.coords_p4['lote_again'])
            time.sleep(2)
            
            # 12. Presionar flecha abajo
            if not self.ahk_click_down.ejecutar_click_down(*self.coords_p4['lote_again'], 1):
                print("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                pyautogui.press('down')
            else:
                print("✅ Flecha abajo presionada con AHK")
            
            time.sleep(2)
            
            # 13. Detectar ventana de error
            if self.detectar_ventana_error_p4():
                print("✅ Ventana de error detectada y cerrada")
            
            return True
            
        except Exception as e:
            logger.error(f"Error en secuencia P4: {e}")
            # Intentar cerrar ventana de error en caso de excepción
            self.detectar_ventana_error_p4()
            return False

    # ===== MÉTODO PRINCIPAL =====
    def ejecutar_proceso_completo(self):
        """Ejecuta el proceso completo unificado"""
        logger.info("Iniciando proceso completo unificado...")
        self.is_running = True
        
        if not self.iniciar_todos_ahk():
            logger.error("No se pudieron iniciar los servicios AHK")
            return False
        
        try:
            # Ejecutar P1
            logger.info("=== EJECUTANDO P1 ===")
            if not self.procesar_p1():
                logger.error("Fallo en P1, deteniendo proceso")
                return False
            
            # Ejecutar P2
            logger.info("=== EJECUTANDO P2 ===")
            if not self.procesar_p2():
                logger.error("Fallo en P2, deteniendo proceso")
                return False
            
            # Ejecutar P3
            logger.info("=== EJECUTANDO P3 ===")
            if not self.procesar_p3():
                logger.error("Fallo en P3, deteniendo proceso")
                return False
            
            # Ejecutar P4
            logger.info("=== EJECUTANDO P4 ===")
            if not self.procesar_p4():
                logger.error("Fallo en P4")
                return False
            
            logger.info("🎉 PROCESO COMPLETO FINALIZADO EXITOSAMENTE!")
            return True
            
        except Exception as e:
            logger.error(f"Error en proceso completo: {e}")
            return False
        finally:
            self.is_running = False
            self.detener_todos_ahk()

# Función principal
def main():
    print("=" * 60)
    print("     AUTOMATIZACIÓN COMPLETA UNIFICADA")
    print("=" * 60)
    print()
    
    # Configuración
    archivo_csv = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
    
    # Crear automatización
    automatizacion = AutomatizacionCompleta(archivo_csv)
    
    # Verificar archivo CSV
    if not os.path.exists(archivo_csv):
        print(f"❌ ERROR: Archivo CSV no encontrado: {archivo_csv}")
        input("Presiona Enter para salir...")
        return
    
    print(f"✅ Archivo CSV encontrado: {archivo_csv}")
    
    # Verificar imágenes de referencia
    images_to_check = [
        automatizacion.reference_image_p2,
        automatizacion.reference_image_p4,
        automatizacion.ventana_archivo_img_p4,
        automatizacion.ventana_error_img_p4,
        "img/ventanaAdministracion4.PNG"
    ]
    
    for image_path in images_to_check:
        if not os.path.exists(image_path):
            print(f"⚠️  Advertencia: Imagen no encontrada: {image_path}")
        else:
            print(f"✅ Imagen encontrada: {image_path}")
    
    print()
    print("Este proceso ejecutará secuencialmente:")
    print("  1. P1: Obtención de línea específica")
    print("  2. P2: Automatización NSE") 
    print("  3. P3: Servicios NSE")
    print("  4. P4: Automatización Google Earth")
    print()
    print("⚠️  ADVERTENCIA: Asegúrate de que las ventanas objetivo estén activas")
    print("   Presiona Ctrl+C para cancelar en cualquier momento")
    print()
    
    try:
        input("Presiona Enter para INICIAR el proceso completo...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO PROCESO COMPLETO...")
        print()
        
        # Ejecutar proceso completo
        success = automatizacion.ejecutar_proceso_completo()
        
        if success:
            print("🎉 TODAS LAS ETAPAS COMPLETADAS EXITOSAMENTE!")
        else:
            print("❌ El proceso encontró errores")
        
        print()
        input("Presiona Enter para salir...")
        
    except KeyboardInterrupt:
        print()
        print("❌ Proceso cancelado por el usuario")
        automatizacion.is_running = False
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        automatizacion.is_running = False

if __name__ == "__main__":
    main()