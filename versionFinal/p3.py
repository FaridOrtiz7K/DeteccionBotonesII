import os
import time
import sys
import pandas as pd
import logging
import pyautogui
from utils.ahk_writer import AHKWriter
from utils.ahk_click_down import AHKClickDown

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
        
        # Inicializar controladores AHK
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        
        # Configurar coordenadas centralizadas para fácil mantenimiento
        self.coords = {
            'menu_principal': (100, 114),
            'campo_cantidad': (127, 383),
            'boton_guardar': (82, 423),
            'boton_error': (704, 384),
            'cierre': (882, 49),
            'inicio_servicios': (1563, 385)
        }

    def iniciar_ahk(self):
        """Iniciar ambos servicios AHK"""
        try:
            if not self.ahk_writer.start_ahk():
                logging.error("No se pudo iniciar AHK Writer")
                return False
            if not self.ahk_click_down.start_ahk():
                logging.error("No se pudo iniciar AHK Click Down")
                return False
            logging.info("✅ Ambos servicios AHK iniciados correctamente")
            return True
        except Exception as e:
            logging.error(f"Error iniciando servicios AHK: {e}")
            return False

    def detener_ahk(self):
        """Detener ambos servicios AHK"""
        try:
            self.ahk_writer.stop_ahk()
            self.ahk_click_down.stop_ahk()
            logging.info("✅ Servicios AHK detenidos correctamente")
        except Exception as e:
            logging.error(f"Error deteniendo servicios AHK: {e}")

    def click(self, x, y, duration=0.1):
        """Hacer clic en coordenadas específicas"""
        pyautogui.click(x, y, duration=duration)
        time.sleep(0.5)

    def write(self, text):
        """Escribir texto usando AHK Writer"""
        try:
            # Primero hacer clic en el campo de cantidad, luego escribir
            if self.click(*self.coords['campo_cantidad']):
                return self.ahk_writer.ejecutar_escritura_ahk(
                    self.coords['campo_cantidad'][0],
                    self.coords['campo_cantidad'][1],
                    str(text)
                )
            return False
        except Exception as e:
            logging.error(f"Error escribiendo texto '{text}': {e}")
            return False

    def press_down(self, times=1):
        """Presionar flecha down usando AHK"""
        try:
            # Usamos AHK Click Down con las veces especificadas
            # Hacemos clic en una posición neutral primero
            return self.ahk_click_down.ejecutar_click_down(100, 100, times)
        except Exception as e:
            logging.error(f"Error presionando DOWN {times} veces: {e}")
            return False

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

    def handle_error_click(self):
        """Manejar clics de error"""
        for _ in range(5):
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
                self.click(*self.coords['menu_principal'])
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
                    
                self.click(*self.coords['cierre'])
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

    # Funciones de servicios (mantenemos las originales pero ahora usan AHK)
    def handle_voz_cobre(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_sdom(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(2)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_cobre_telmex(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(2)
        self.click(159, 355)
        self.sleep(2)
        self.press_down(1)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_fibra_telmex(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(2)
        self.click(152, 294)
        self.sleep(2)
        self.press_down(1)
        self.click(150, 323)
        self.sleep(2)
        self.press_down(1)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_tv_cable_otros(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(3)
        self.click(150, 323)
        self.sleep(2)
        self.press_down(4)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_dish(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(3)
        self.click(152, 294)
        self.sleep(2)
        self.press_down(2)
        self.click(150, 323)
        self.sleep(2)
        self.press_down(1)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_tvs(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(3)
        self.click(152, 294)
        self.sleep(2)
        self.press_down(2)
        self.click(150, 323)
        self.sleep(2)
        self.press_down(2)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_sky(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(3)
        self.click(152, 294)
        self.sleep(2)
        self.press_down(2)
        self.click(150, 323)
        self.sleep(2)
        self.press_down(3)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

    def handle_vetv(self, cantidad):
        self.click(*self.coords['menu_principal'])
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(3)
        self.click(152, 294)
        self.sleep(2)
        self.press_down(2)
        self.click(150, 323)
        self.sleep(2)
        self.press_down(5)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(*self.coords['boton_guardar'])
        self.sleep(2)
        self.handle_error_click()

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
        print("✅ Dependencias verificadas")
    except ImportError as e:
        print(f"❌ Error: {e}")
        print("Instala las dependencias con: pip install pandas")
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