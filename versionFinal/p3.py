import os
import time
import sys
import pandas as pd
import logging
import pyautogui
import cv2
import numpy as np
from PIL import ImageGrab
from utils.ahk_writer import AHKWriter
from utils.ahk_click_down import AHKClickDown
from utils.ahk_enter import EnterAHKManager

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('nse_automation.log'),
        logging.StreamHandler()
    ]
)

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
            if not self.ahk_enter.start_ahk():  # ¡FALTABA ESTA LÍNEA!
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
            self.ahk_enter.stop_ahk()  # ¡FALTABA ESTA LÍNEA!
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
                return False
            
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

# Las funciones clear_screen(), print_header() y main() permanecen igual...
def clear_screen():
    """Limpia la pantalla de la consola"""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_header():
    """Imprime el encabezado del programa"""
    print("=" * 60)
    print("     CONTROLADOR NSE - LÍNEA ESPECÍFICA (PYTHON + AHK)")
    print("=" * 60)
    print()

def main():
    clear_screen()
    print_header()
    
    # Solicitar línea específica al usuario
    try:
        linea_input = input("🔢 Ingresa el número de línea a procesar (ej: 5): ").strip()
        if not linea_input:
            print("❌ No se ingresó número de línea")
            return
            
        linea_especifica = int(linea_input)
        if linea_especifica < 1:
            print("❌ El número de línea debe ser mayor a 0")
            return
    except ValueError:
        print("❌ Por favor ingresa un número válido")
        return
    
    # Inicializar automatización
    nse = NSEServicesAutomation(linea_especifica=linea_especifica)
    
    # Verificar dependencias
    print("Verificando dependencias...")
    try:
        import pandas as pd
        import cv2
        import numpy as np
        from PIL import ImageGrab
        print("✅ Dependencias verificadas")
    except ImportError as e:
        print(f"❌ Error: {e}")
        print("Instala las dependencias con: pip install pandas opencv-python pillow numpy")
        input("Presiona Enter para salir...")
        return
    
    # Verificar archivo CSV
    if not os.path.exists(nse.csv_file):
        print(f"❌ ERROR: Archivo CSV no encontrado: {nse.csv_file}")
        input("Presiona Enter para salir...")
        return
    
    print(f"✅ Archivo CSV encontrado: {nse.csv_file}")
    
    print()
    print("Configuración:")
    print(f"  - Línea a procesar: {linea_especifica}")
    print(f"  - Archivo CSV: {nse.csv_file}")
    print()
    
    # Iniciar servicios AHK
    print("🔄 Iniciando servicios AHK...")
    if not nse.iniciar_ahk():
        print("❌ No se pudieron iniciar los servicios AHK")
        input("Presiona Enter para salir...")
        return
    
    # Confirmación final antes de ejecutar
    print("⚠️  ADVERTENCIA: El script ejecutará SOLO la línea especificada")
    print("   Asegúrate de que ya se ejecutó el programa principal hasta CERRAR")
    print("   El script comenzará en 3 segundos")
    print("   Presiona Ctrl+C para cancelar")
    print()
    
    try:
        input("Presiona Enter para INICIAR procesamiento...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print(f"🚀 INICIANDO PROCESAMIENTO DE LÍNEA {linea_especifica}...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar procesamiento de línea específica
        nse.is_running = True
        resultado = nse.procesar_linea_especifica()
        
        if resultado:
            print(f"🎉 LÍNEA {linea_especifica} PROCESADA EXITOSAMENTE!")
        else:
            print(f"❌ HUBO PROBLEMAS PROCESANDO LA LÍNEA {linea_especifica}")
        
        print()
        input("Presiona Enter para salir...")
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
    finally:
        nse.is_running = False
        nse.detener_ahk()

if __name__ == "__main__":
    main()