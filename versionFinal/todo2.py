import os
import time
import sys
import pandas as pd
import pyautogui
import cv2
import numpy as np
import logging
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
        logging.FileHandler('automation_unified.log'),
        logging.StreamHandler()
    ]
)

class UnifiedAutomation:
    def __init__(self):
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.df = None
        self.current_row = None
        self.current_id = None
        
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
        self.parte2_config = {
            'reference_image': "img/VentanaAsignar.png",
            'coords_select': {
                7: [33, 92], 8: [33, 131], 9: [33, 159], 10: [33, 197],
                11: [33, 231], 12: [398, 92], 13: [398, 131], 14: [398, 159],
                15: [33, 301], 16: [33, 333], 17: [33, 367]
            },
            'coords_type': {
                7: [163, 92], 8: [163, 131], 9: [163, 159], 10: [163, 197],
                11: [163, 231], 12: [528, 92], 13: [528, 131], 14: [528, 159],
                15: [163, 301], 16: [163, 333], 17: [163, 367]
            },
            'coords_asignar': [446, 281],
            'coords_cerrar': [396, 352]
        }
        
        self.parte3_config = {
            'reference_image': "img/ventanaAdministracion4.PNG",
            'coords': {
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
        }
        
        self.parte4_config = {
            'reference_image': "img/textoAdicional.PNG",
            'ventana_archivo_img': "img/cargarArchivo.png",
            'ventana_error_img': "img/ventanaError.png",
            'coords': {
                'agregar_ruta': (327, 381),
                'archivo': (1396, 608),
                'abrir': (1406, 634),
                'seleccionar_mapa': (168, 188),
                'anotar': (1366, 384),
                'agregar_texto_adicional': (1449, 452),
                'limpiar_trazo': (360, 980),
                'lote_again': (70, 266)
            },
            'coords_texto_relativas': {
                'campo_texto': (230, 66),
                'agregar_texto': (64, 100),
                'cerrar_ventana_texto': (139, 98)
            }
        }

    def cargar_csv(self):
        """Carga el archivo CSV"""
        try:
            self.df = pd.read_csv(self.csv_file)
            logging.info(f"CSV cargado correctamente: {len(self.df)} registros")
            return True
        except Exception as e:
            logging.error(f"Error cargando CSV: {e}")
            return False

    def iniciar_ahk(self):
        """Inicia todos los procesos AHK"""
        logging.info("Iniciando procesos AHK...")
        return (self.ahk_manager_cd.start_ahk() and 
                self.ahk_writer.start_ahk() and 
                self.ahk_click_down.start_ahk() and
                self.ahk_enter.start_ahk() and
                self.ahk_manager.start_ahk())

    def detener_ahk(self):
        """Detiene todos los procesos AHK"""
        logging.info("Deteniendo procesos AHK...")
        self.ahk_manager_cd.stop_ahk()
        self.ahk_writer.stop_ahk()
        self.ahk_click_down.stop_ahk()
        self.ahk_enter.stop_ahk()
        self.ahk_manager.stop_ahk()

    def click(self, x, y, duration=0.1):
        """Hacer clic en coordenadas específicas"""
        pyautogui.click(x, y, duration=duration)
        time.sleep(0.5)

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

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

    # ========== PARTE 1: OBTENER ID Y PROCESAR REGISTRO ==========
    def ejecutar_parte1(self):
        """Ejecuta la parte 1: Obtener ID y procesar registro"""
        try:
            # Paso 2: Click en (89, 263)
            logging.info("Paso 2: Click en (89, 263)")
            self.click(89, 263)
            self.sleep(1)
            
            # Paso 3: Usar AHKManager en (1483, 519) para obtener ID
            logging.info("Paso 3: Obteniendo ID con AHKManager en (1483, 519)")
            id_obtenido = self.ahk_manager_cd.ejecutar_acciones_ahk(1483, 519)
            
            if not id_obtenido:
                logging.error("No se pudo obtener el ID")
                return False
                
            self.current_id = int(id_obtenido)
            logging.info(f"ID obtenido: {self.current_id}")
            
            # Paso 4: Buscar el ID en el CSV
            logging.info(f"Paso 4: Buscando ID {self.current_id} en CSV")
            if self.df is None:
                logging.error("CSV no cargado")
                return False
                
            # Buscar en la primera columna (asumimos que es la columna 0)
            resultado = self.df[self.df.iloc[:, 0] == self.current_id]
            
            if len(resultado) == 0:
                logging.warning(f"ID {self.current_id} no encontrado en el CSV")
                return False
            
            self.current_row = resultado.iloc[0]
            logging.info(f"ID {self.current_id} encontrado, datos: {self.current_row.tolist()}")
            
            # Paso 5: Escribir valor de columna 2 en (1483, 519)
            if len(self.current_row) >= 2:  # Verificar que existe columna 2
                valor_columna_2 = str(self.current_row.iloc[1])
                logging.info(f"Paso 5: Escribiendo valor '{valor_columna_2}' en (1483, 519)")
                
                exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, valor_columna_2)
                if not exito_escritura:
                    logging.error("Error en la escritura")
                    return False
            else:
                logging.warning("No hay columna 2 en el registro")
            
            # Paso 6: Revisar si columna 4 es mayor a 0
            if len(self.current_row) >= 4:  # Verificar que existe columna 4
                valor_columna_4 = self.current_row.iloc[3]
                logging.info(f"Paso 6: Valor columna 4 = {valor_columna_4}")
                
                # Paso 7: Si es mayor a 0, usar AHKClickDown
                if pd.notna(valor_columna_4) and float(valor_columna_4) > 0:
                    veces_down = int(float(valor_columna_4))
                    logging.info(f"Paso 7: Ejecutando {veces_down} veces DOWN en (1507, 636)")
                    
                    exito_down = self.ahk_click_down.ejecutar_click_down(1507, 636, veces_down)
                    if not exito_down:
                        logging.error("Error en click + down")
                        return False
                else:
                    logging.info("Paso 7: Saltado (columna 4 <= 0)")
            else:
                logging.warning("No hay columna 4 en el registro")
            
            # Paso 8: Click en (1290, 349)
            logging.info("Paso 8: Click en (1290, 349)")
            self.click(1290, 349)
            self.sleep(1)
            
            logging.info("Parte 1 completada exitosamente")
            return True
            
        except Exception as e:
            logging.error(f"Error en ejecutar_parte1: {e}")
            return False

    # ========== PARTE 2: ASIGNAR NIVELES SOCIOECONÓMICOS ==========
    def ejecutar_parte2(self):
        """Ejecuta la parte 2: Asignar NSE"""
        if self.current_row is None:
            logging.error("No hay fila actual para procesar en parte 2")
            return False

        try:
            # Verificar si se debe saltar el proceso (columna 6 tiene valor)
            if self.should_skip_process():
                logging.info(f"⏭️  Saltando parte 2 - Columna 6 tiene valor: {self.current_row[5]}")
                return True
            
            # Verificar que sea tipo V
            if str(self.current_row[4]).strip().upper() != "V":
                logging.info(f"⚠️  Saltando parte 2 - No es tipo V: {self.current_row[4]}")
                return True
            
            # Click en el boton seleccionar lote 
            self.click(169, 189)
            self.sleep(2)
            
            # Click en el boton asignar nse
            self.click(1491, 386)
            self.sleep(2)
            
            # Esperar imagen de referencia
            image_found, base_location = self.wait_for_image_with_retries(
                self.parte2_config['reference_image'], 
                max_attempts=30
            )
            
            if not image_found:
                logging.error("❌ No se puede continuar sin detectar la imagen de referencia.")
                return False
            
            # Procesar tipo V
            logging.info("🎯 Imagen detectada, procediendo con tipo V")
            self.handle_type_v(base_location)
            
            logging.info("✅ Parte 2 completada exitosamente")
            return True
            
        except Exception as e:
            logging.error(f"Error en ejecutar_parte2: {e}")
            return False

    def should_skip_process(self):
        """Determina si se debe saltar el proceso basado en la columna 6"""
        if pd.notna(self.current_row[5]):
            col_value = str(self.current_row[5]).strip()
            if col_value and col_value != "" and col_value != "nan":
                return True
        return False

    def handle_type_v(self, base_location):
        """Manejar tipo V con coordenadas relativas - COLUMNAS 7-17"""
        base_x, base_y = base_location
        
        # Lógica V para columnas 7-17 con coordenadas relativas
        for col_index in range(7, 18):  # 7 a 17 inclusive
            if pd.notna(self.current_row[col_index-1]) and self.current_row[col_index-1] > 0:
                # Usar coordenadas relativas de la tabla verde, sumando a la base
                x_cs_rel, y_cs_rel = self.parte2_config['coords_select'][col_index]
                x_ct_rel, y_ct_rel = self.parte2_config['coords_type'][col_index]
                
                # Calcular coordenadas absolutas
                x_cs_abs = base_x + x_cs_rel
                y_cs_abs = base_y + y_cs_rel
                x_ct_abs = base_x + x_ct_rel
                y_ct_abs = base_y + y_ct_rel
                
                self.click(x_cs_abs, y_cs_abs)
                self.sleep(2)
                
                # Usar AHKWriter para escribir
                texto = str(int(self.current_row[col_index-1]))
                success = self.ahk_writer.ejecutar_escritura_ahk(x_ct_abs, y_ct_abs, texto)
                if not success:
                    logging.error(f"❌ Error al escribir con AHK en ({x_ct_abs}, {y_ct_abs}): {texto}")
                    return False
                self.sleep(2)
        
        # Botón ASIGNAR antes de cerrar
        x_asignar_rel, y_asignar_rel = self.parte2_config['coords_asignar']
        x_asignar_abs = base_x + x_asignar_rel
        y_asignar_abs = base_y + y_asignar_rel
        self.click(x_asignar_abs, y_asignar_abs)
        self.sleep(2)
        
        # Botón CERRAR
        x_cerrar_rel, y_cerrar_rel = self.parte2_config['coords_cerrar']
        x_cerrar_abs = base_x + x_cerrar_rel
        y_cerrar_abs = base_y + y_cerrar_rel
        self.click(x_cerrar_abs, y_cerrar_abs)
        self.sleep(2)
        
        return True

    # ========== PARTE 3: ASIGNAR SERVICIOS ==========
    def ejecutar_parte3(self):
        """Ejecuta la parte 3: Asignar servicios"""
        if self.current_row is None:
            logging.error("No hay fila actual para procesar en parte 3")
            return False

        try:
            # Solo procesar servicios si la columna 18 tiene valor > 0
            if pd.notna(self.current_row[17]) and self.current_row[17] > 0:
                logging.info(f"✅ Línea tiene servicios para procesar: {self.current_row[17]}")
                
                self.click(*self.parte3_config['coords']['inicio_servicios'])
                self.sleep(2)
                
                # Buscar ventana de servicios
                logging.info("🔍 Buscando ventana de servicios...")
                referencia = self.buscar_imagen(
                    self.parte3_config['reference_image'], 
                    timeout=30
                )
                
                if referencia is None:
                    logging.error("❌ ERROR: No se pudo encontrar la ventana de servicios")
                    return False
                
                # Actualizar coordenadas relativas
                if not self.actualizar_coordenadas_relativas(referencia):
                    logging.error("❌ ERROR: No se pudieron actualizar las coordenadas relativas")
                    return False
                
                # Continuar con el procesamiento normal usando coordenadas relativas
                self.click(*self.coords_relativas['menu_principal'])
                self.sleep(2)
                    
                # Procesar servicios
                servicios_procesados = self.procesar_servicios()
                
                # Usar coordenadas relativas para el cierre
                self.click(*self.coords_relativas['cierre'])
                self.sleep(5)
                
                logging.info(f"✅ Parte 3 completada: {servicios_procesados} servicios procesados")
                return True
            else:
                logging.info("⏭️  No tiene servicios para procesar")
                return True
            
        except Exception as e:
            logging.error(f"Error en ejecutar_parte3: {e}")
            return False

    def actualizar_coordenadas_relativas(self, referencia):
        """Actualiza coordenadas para que sean relativas al punto de referencia"""
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
        self.coords_relativas['boton_error'] = self.parte3_config['coords']['boton_error']
        self.coords_relativas['inicio_servicios'] = self.parte3_config['coords']['inicio_servicios']
        
        logging.info("✅ Coordenadas actualizadas a relativas")
        return True

    def procesar_servicios(self):
        """Procesa todos los servicios de la fila actual"""
        servicios_procesados = 0
        
        if pd.notna(self.current_row[18]) and self.current_row[18] > 0:  # VOZ COBRE TELMEX
            logging.info(f"  └─ Procesando VOZ COBRE TELMEX: {self.current_row[18]}")
            self.handle_voz_cobre(self.current_row[18])
            servicios_procesados += 1
            
        if pd.notna(self.current_row[19]) and self.current_row[19] > 0:  # Datos s/dom
            logging.info(f"  └─ Procesando DATOS S/DOM: {self.current_row[19]}")
            self.handle_datos_sdom(self.current_row[19])
            servicios_procesados += 1
            
        if pd.notna(self.current_row[20]) and self.current_row[20] > 0:  # Datos-cobre-telmex-inf
            logging.info(f"  └─ Procesando DATOS COBRE TELMEX: {self.current_row[20]}")
            self.handle_datos_cobre_telmex(self.current_row[20])
            servicios_procesados += 1
            
        if pd.notna(self.current_row[21]) and self.current_row[21] > 0:  # Datos-fibra-telmex-inf
            logging.info(f"  └─ Procesando DATOS FIBRA TELMEX: {self.current_row[21]}")
            self.handle_datos_fibra_telmex(self.current_row[21])
            servicios_procesados += 1
            
        if pd.notna(self.current_row[22]) and self.current_row[22] > 0:  # TV cable otros
            logging.info(f"  └─ Procesando TV CABLE OTROS: {self.current_row[22]}")
            self.handle_tv_cable_otros(self.current_row[22])
            servicios_procesados += 1
            
        if pd.notna(self.current_row[23]) and self.current_row[23] > 0:  # Dish
            logging.info(f"  └─ Procesando DISH: {self.current_row[23]}")
            self.handle_dish(self.current_row[23])
            servicios_procesados += 1
            
        if pd.notna(self.current_row[24]) and self.current_row[24] > 0:  # TVS
            logging.info(f"  └─ Procesando TVS: {self.current_row[24]}")
            self.handle_tvs(self.current_row[24])
            servicios_procesados += 1
            
        if pd.notna(self.current_row[25]) and self.current_row[25] > 0:  # SKY
            logging.info(f"  └─ Procesando SKY: {self.current_row[25]}")
            self.handle_sky(self.current_row[25])
            servicios_procesados += 1
            
        if pd.notna(self.current_row[26]) and self.current_row[26] > 0:  # VETV
            logging.info(f"  └─ Procesando VETV: {self.current_row[26]}")
            self.handle_vetv(self.current_row[26])
            servicios_procesados += 1
        
        return servicios_procesados

    # Métodos de servicios (similares a los de p3.py)
    def handle_voz_cobre(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_sdom(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down_servicio(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter_servicio()
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_cobre_telmex(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down_servicio(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_producto'], 1)
        self.press_enter_servicio()
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_fibra_telmex(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down_servicio(*self.coords_relativas['casilla_servicio'], 2)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_tipo'], 1)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_empresa'], 1)
        self.press_enter_servicio()
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_tv_cable_otros(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down_servicio(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_empresa'], 4)
        self.press_enter_servicio()
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_dish(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down_servicio(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_empresa'], 1)
        self.press_enter_servicio()
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_tvs(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down_servicio(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_empresa'], 2)
        self.press_enter_servicio()
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_sky(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down_servicio(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_empresa'], 3)
        self.press_enter_servicio()
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_vetv(self, cantidad):
        self.click(*self.coords_relativas['menu_principal'])
        self.sleep(2)
        self.press_down_servicio(*self.coords_relativas['casilla_servicio'], 3)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_tipo'], 2)
        self.press_enter_servicio()
        self.press_down_servicio(*self.coords_relativas['casilla_empresa'], 5)
        self.press_enter_servicio()
        self.sleep(2)
        self.write_servicio(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords_relativas['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def write_servicio(self, text):
        """Escribir texto en servicio usando AHK Writer"""
        campo_coords = self.coords_relativas['campo_cantidad']
        self.click(*campo_coords)
        return self.ahk_writer.ejecutar_escritura_ahk(campo_coords[0], campo_coords[1], text)

    def press_down_servicio(self, x, y, times=1):
        """Presionar flecha down en servicio usando AHK"""
        return self.ahk_click_down.ejecutar_click_down(x, y, times)

    def press_enter_servicio(self):
        """Presionar enter en servicio usando AHK"""
        return self.ahk_enter.presionar_enter(1)

    def handle_error_click(self):
        """Manejar clics de error"""
        for _ in range(5):
            self.click(*self.coords_relativas['boton_error'])
            self.sleep(2)

    # ========== PARTE 4: ESCRIBIR NOMBRE DEL NEGOCIO ==========
    def ejecutar_parte4(self):
        """Ejecuta la parte 4: Escribir nombre del negocio"""
        if self.current_row is None:
            logging.error("No hay fila actual para procesar en parte 4")
            return False

        try:
            # Obtener valores del CSV
            num_txt_type = str(int(self.current_row[28])) if not pd.isna(self.current_row[28]) else None
            texto_adicional = str(self.current_row[29]) if not pd.isna(self.current_row[29]) else ""

            if not num_txt_type:
                logging.info("⚠️  num_txt_type vacío, saltando parte 4...")
                return True

            nombre_archivo = f"NN {num_txt_type}.kml"
            logging.info(f"📁 Archivo a cargar: {nombre_archivo}")
            logging.info(f"📝 Texto adicional: '{texto_adicional}'")

            # SECUENCIA DE ACCIONES
            # 1. Seleccionar Agregar ruta de GE
            self.click(*self.parte4_config['coords']['agregar_ruta'])
            self.sleep(2)
            self.click(*self.parte4_config['coords']['archivo'])
            self.sleep(2)
            self.click(*self.parte4_config['coords']['abrir'])
            self.sleep(2) 
            
            # 2. Usar detección de ventana de archivo para cargar el archivo
            success = self.handle_archivo_special_behavior(nombre_archivo)
            
            if not success:
                logging.error("❌ No se pudo cargar el archivo. Regresando a agregar_ruta...")
                self.click(*self.parte4_config['coords']['agregar_ruta'])
                self.sleep(2)
                return False
            
            # 3. Presionar Enter para confirmar la carga del archivo
            if not self.presionar_enter_ahk(1):
                logging.warning("⚠️  No se pudo presionar Enter con AHK, usando pyautogui")
                pyautogui.press('enter')
            
            self.sleep(3)

            self.click(*self.parte4_config['coords']['agregar_ruta'])
            self.sleep(2)

            self.click(1406, 675)  # cargar ruta
            self.sleep(2)

            self.click(70, 266)  # seleccionar lote
            self.sleep(2)
            
            # 4. Seleccionar en el mapa
            self.click(*self.parte4_config['coords']['seleccionar_mapa'])
            self.sleep(2)
            
            # 5. Anotar
            self.click(*self.parte4_config['coords']['anotar'])
            self.sleep(2)
            
            # 6. Agregar texto adicional
            self.click(*self.parte4_config['coords']['agregar_texto_adicional'])
            self.sleep(2)
            
            # 7. DETECCIÓN DE IMAGEN para el campo de texto
            image_found, base_location = self.wait_for_image_with_retries(
                self.parte4_config['reference_image'], 
                max_attempts=10
            )
            
            if image_found:
                # Calcular coordenadas absolutas del campo de texto
                x_campo = base_location[0] + self.parte4_config['coords_texto_relativas']['campo_texto'][0]
                y_campo = base_location[1] + self.parte4_config['coords_texto_relativas']['campo_texto'][1]
                x_agregar = base_location[0] + self.parte4_config['coords_texto_relativas']['agregar_texto'][0]
                y_agregar = base_location[1] + self.parte4_config['coords_texto_relativas']['agregar_texto'][1]
                x_cerrar = base_location[0] + self.parte4_config['coords_texto_relativas']['cerrar_ventana_texto'][0]
                y_cerrar = base_location[1] + self.parte4_config['coords_texto_relativas']['cerrar_ventana_texto'][1]
                
                # Escribir el texto adicional si hay texto
                if texto_adicional and texto_adicional.strip():
                    writing_success = self.escribir_texto_adicional_ahk(x_campo, y_campo, texto_adicional)
                    if not writing_success:
                        logging.warning("⚠️  Falló la escritura con AHK, intentando con pyautogui...")
                        pyautogui.write(texto_adicional, interval=0.05)
                else:
                    logging.info("ℹ️  Texto adicional vacío, no se escribe nada")
                
                self.sleep(2)

                # 8. Agregar de texto adicional
                self.click(x_agregar, y_agregar)
                self.sleep(3)
            
                # 9. Cerrar ventana de texto adicional
                self.click(x_cerrar, y_cerrar)
                self.sleep(2)
            else:
                logging.error("❌ No se pudo detectar la imagen del campo de texto")
                return False
            
            # 10. Limpiar trazo
            self.click(*self.parte4_config['coords']['limpiar_trazo'])
            self.sleep(1)
            
            # 11. Seleccionar Lote nuevamente
            self.click(*self.parte4_config['coords']['lote_again'])
            self.sleep(2)
            
            # 12. Presionar flecha abajo con AHK
            if not self.presionar_flecha_abajo_ahk(*self.parte4_config['coords']['lote_again'], 1):
                logging.warning("⚠️  No se pudo presionar flecha abajo con AHK, usando pyautogui")
                pyautogui.press('down')
            else:
                logging.info("✅ Flecha abajo presionada con AHK")
            
            self.sleep(2)
            
            # 13. Detectar ventana de error
            if self.detectar_ventana_error():
                logging.info("✅ Ventana de error detectada y cerrada")
            
            logging.info("✅ Parte 4 completada exitosamente")
            return True
            
        except Exception as e:
            logging.error(f"Error en ejecutar_parte4: {e}")
            # Intentar cerrar ventana de error en caso de excepción
            self.detectar_ventana_error()
            return False

    # Métodos auxiliares para Parte 4
    def encontrar_ventana_archivo(self):
        """Busca la ventana de archivo usando template matching con reintentos inteligentes"""
        intentos = 1
        confianza_minima = 0.6
        tiempo_espera_base = 1
        tiempo_espera_largo = 10
        
        # Cargar template una sola vez fuera del bucle
        template = cv2.imread(self.parte4_config['ventana_archivo_img'])
        if template is None:
            logging.error(f"No se pudo cargar la imagen '{self.parte4_config['ventana_archivo_img']}'")
            return None
        
        while True: 
            try:
                # Capturar pantalla completa
                screenshot = pyautogui.screenshot()
                pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                # Realizar template matching
                result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)
                
                if max_val >= confianza_minima:
                    logging.info(f"Ventana encontrada con confianza: {max_val:.2f}")
                    return max_loc
                else:
                    # Estrategia de espera progresiva
                    if intentos % 10 == 0 and intentos > 0:
                        logging.info(f"Intento {intentos}: Mejor coincidencia: {max_val:.2f}")
                        logging.info("Esperando 10 segundos...")
                        time.sleep(tiempo_espera_largo)
                    else:
                        time.sleep(tiempo_espera_base)
                    intentos += 1
                    
            except Exception as e:
                logging.error(f"Error durante la búsqueda: {e}")
                time.sleep(tiempo_espera_base)
                intentos += 1

    def handle_archivo_special_behavior(self, nombre_archivo):
        """Maneja el comportamiento especial para cargar archivos usando AHK Manager"""
        # Buscar la ventana de archivo
        coordenadas_ventana = self.encontrar_ventana_archivo()

        if coordenadas_ventana:
            x_ventana, y_ventana = coordenadas_ventana
            logging.info(f"Coordenadas ventana: x={x_ventana}, y={y_ventana}")
            
            # Calcular coordenadas del campo de texto
            x_campo = x_ventana + 294
            y_campo = y_ventana + 500
            logging.info(f"Coordenadas campo texto: x={x_campo}, y={y_campo}")
            
            # Iniciar AHK Manager para escribir el nombre del archivo
            if not self.ahk_manager.start_ahk():
                logging.error("No se pudo iniciar AutoHotkey")
                return False
            
            # Enviar comandos a AHK Manager para escribir el nombre del archivo
            if self.ahk_manager.ejecutar_acciones_ahk(x_campo, y_campo, nombre_archivo):
                time.sleep(1.5)  # Esperar a que AHK termine
            else:
                logging.error("Error enviando comando a AHK")
                return False
            
            self.ahk_manager.stop_ahk()
            return True
        else:
            logging.error("No se pudo encontrar la ventana de archivo.")
            return False

    def escribir_texto_adicional_ahk(self, x, y, texto):
        """Escribe texto adicional usando AHK Writer"""
        if not texto or pd.isna(texto) or str(texto).strip() == '':
            logging.info("⚠️  Texto adicional vacío, saltando escritura")
            return True
            
        logging.info(f"📝 Intentando escribir texto: '{texto}' en coordenadas ({x}, {y})")
        
        # Verificar que las coordenadas son válidas
        if x <= 0 or y <= 0:
            logging.error(f"❌ Coordenadas inválidas: ({x}, {y})")
            return False
        
        if not self.ahk_writer.start_ahk():
            logging.error("No se pudo iniciar AHK Writer")
            return False
        
        logging.info("🔄 AHK Writer iniciado, enviando comando...")
        success = self.ahk_writer.ejecutar_escritura_ahk(x, y, texto)
        self.ahk_writer.stop_ahk()
        
        if success:
            logging.info(f"✅ Texto escrito exitosamente: '{texto}'")
        else:
            logging.error(f"❌ Error al escribir texto: '{texto}'")
            # Método de fallback con pyautogui
            try:
                self.click(x, y)
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.5)
                pyautogui.press('delete')
                time.sleep(0.5)
                pyautogui.write(texto, interval=0.05)
                logging.info(f"✅ Texto escrito con pyautogui: '{texto}'")
                success = True
            except Exception as e:
                logging.error(f"❌ También falló pyautogui: {e}")
                
        return success

    def presionar_flecha_abajo_ahk(self, x, y, veces=1):
        """Presiona flecha abajo usando AHK"""
        if not self.ahk_click_down.start_ahk():
            logging.error("No se pudo iniciar AutoHotkey para flecha abajo")
            return False
        
        try:
            self.ahk_click_down.ejecutar_click_down(x, y, veces)
            return True
        except Exception as e:
            logging.error(f"Error presionando flecha abajo: {e}")
            return False
        finally:
            self.ahk_click_down.stop_ahk()

    def presionar_enter_ahk(self, veces=1):
        """Presiona Enter usando AHK"""
        if not self.ahk_enter.start_ahk():
            logging.error("No se pudo iniciar AutoHotkey para Enter")
            return False
        
        success = self.ahk_enter.presionar_enter(veces)
        self.ahk_enter.stop_ahk()
        return success

    def detectar_ventana_error(self):
        """Detecta la ventana de error y presiona Enter para cerrarla"""
        try:
            # Cargar template de la ventana de error
            template = cv2.imread(self.parte4_config['ventana_error_img']) 
            if template is None:
                logging.error(f"No se pudo cargar la imagen '{self.parte4_config['ventana_error_img']}'")
                return False
            
            # Capturar pantalla completa
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # Realizar template matching
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            confianza_minima = 0.6
            
            if max_val >= confianza_minima:
                logging.info(f"Ventana de error detectada con confianza: {max_val:.2f}")
                
                # Presionar Enter para cerrar la ventana de error usando AHK
                if not self.ahk_enter.start_ahk():
                    logging.error("No se pudo iniciar AutoHotkey")
                    return False
                    
                # Enviar comandos a AHK
                if self.ahk_enter.presionar_enter(1):
                    time.sleep(2.5)
                else:
                    logging.error("Error enviando comando a AHK")
                    return False
                
                logging.info("Ventana de error detectada y cerrada")
                self.ahk_enter.stop_ahk()
                return True
            else:
                return False
                
        except Exception as e:
            logging.error(f"Error al detectar ventana de error: {e}")
            return False

    def detect_image_with_cv2(self, image_path, confidence=0.7):
        """Detectar imagen en pantalla usando template matching con OpenCV"""
        try:
            # Capturar pantalla completa
            screenshot = pyautogui.screenshot()
            pantalla = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # Cargar template de la imagen de referencia
            template = cv2.imread(image_path)
            if template is None:
                logging.error(f"Error: No se pudo cargar la imagen {image_path}")
                return False, None
            
            # Realizar template matching
            result = cv2.matchTemplate(pantalla, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            # Umbral de confianza
            if max_val < confidence:
                logging.info(f"Imagen no encontrada. Mejor coincidencia: {max_val:.2f}")
                return False, None
            
            logging.info(f"✅ Imagen encontrada con confianza: {max_val:.2f}")
            return True, max_loc
        except Exception as e:
            logging.error(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image_with_retries(self, image_path, max_attempts=30, confidence=0.7):
        """Esperar a que aparezca una imagen con múltiples intentos usando OpenCV"""
        logging.info(f"🔍 Buscando imagen: {image_path}")
        
        for attempt in range(1, max_attempts + 1):
            found, location = self.detect_image_with_cv2(image_path, confidence)
            
            if found:
                logging.info(f"✅ Imagen detectada en el intento {attempt} en coordenadas: {location}")
                return True, location
            
            logging.info(f"⏳ Intento {attempt}/{max_attempts} - Imagen no encontrada")
            
            if attempt < max_attempts:
                if attempt % 10 == 0:
                    logging.info("⏰ Espera prolongada de 10 segundos...")
                    time.sleep(10)
                else:
                    time.sleep(2)
        
        logging.error("❌ Imagen no encontrada después de 30 intentos. Terminando proceso.")
        return False, None

    # ========== MÉTODO PRINCIPAL ==========
    def ejecutar_proceso_completo(self):
        """Ejecuta el proceso completo unificado"""
        if not self.cargar_csv():
            return False
            
        if not self.iniciar_ahk():
            return False
        
        try:
            # Ejecutar las partes en secuencia
            logging.info("🔄 Iniciando Parte 1: Obtener ID y procesar registro...")
            if not self.ejecutar_parte1():
                logging.error("❌ Falló la Parte 1")
                return False
            
            logging.info("🔄 Iniciando Parte 2: Asignar niveles socioeconómicos...")
            if not self.ejecutar_parte2():
                logging.error("❌ Falló la Parte 2")
                return False
            
            logging.info("🔄 Iniciando Parte 3: Asignar servicios...")
            if not self.ejecutar_parte3():
                logging.error("❌ Falló la Parte 3")
                return False
            
            logging.info("🔄 Iniciando Parte 4: Escribir nombre del negocio...")
            if not self.ejecutar_parte4():
                logging.error("❌ Falló la Parte 4")
                return False
            
            logging.info("🎉 PROCESO COMPLETADO EXITOSAMENTE!")
            return True
            
        finally:
            # Siempre detener AHK al finalizar
            self.detener_ahk()


# Función principal
def main():
    # Configurar pyautogui
    pyautogui.PAUSE = 0.5
    pyautogui.FAILSAFE = True
    
    # Crear automatización unificada
    automation = UnifiedAutomation()
    
    # Ejecutar procesamiento
    print("Iniciando proceso unificado...")
    print("Asegúrate de que la ventana objetivo esté activa")
    print("Presiona Ctrl+C para cancelar")
    
    try:
        input("Presiona Enter para comenzar...")
        time.sleep(3)  # Tiempo para cambiar a la ventana correcta
        automation.ejecutar_proceso_completo()
    except KeyboardInterrupt:
        print("\nProceso cancelado por el usuario")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("Proceso finalizado")

if __name__ == "__main__":
    main()