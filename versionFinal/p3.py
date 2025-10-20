import os
import time
import sys
import pyautogui
import pandas as pd

class NSEServicesAutomation:
    def __init__(self):
        self.start_count = 1
        self.loop_count = 589
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.current_line = 0
        self.is_running = False
        
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5

    def click(self, x, y, duration=0.1):
        """Hacer clic en coordenadas específicas"""
        pyautogui.click(x, y, duration=duration)
        time.sleep(0.5)

    def write(self, text):
        """Escribir texto"""
        pyautogui.write(str(text))
        time.sleep(0.5)

    def press_enter(self):
        """Presionar Enter"""
        pyautogui.press('enter')
        time.sleep(0.5)

    def press_down(self, times=1):
        """Presionar flecha abajo"""
        for _ in range(times):
            pyautogui.press('down')
            time.sleep(0.5)

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

    def execute_services_script(self):
        """Función principal para procesar servicios"""
        try:
            # Leer CSV
            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            for i in range(self.start_count - 1, min(self.start_count + self.loop_count - 1, total_lines)):
                if not self.is_running:
                    break
                    
                self.current_line = i + 1
                print(f"🔄 Procesando servicios línea {self.current_line}/{total_lines}")
                
                row = df.iloc[i]
                
                # Solo procesar servicios si la columna 18 tiene valor > 0
                if pd.notna(row[17]) and row[17] > 0:
                    self.click(1563, 385)
                    self.sleep(2)
                    self.click(100, 114)
                    self.sleep(2)
                    
                    # Llamar a funciones de servicios
                    if pd.notna(row[18]) and row[18] > 0:  # VOZ COBRE TELMEX
                        self.handle_voz_cobre(row[18])
                    
                    if pd.notna(row[19]) and row[19] > 0:  # Datos s/dom
                        self.handle_datos_sdom(row[19])
                    
                    if pd.notna(row[20]) and row[20] > 0:  # Datos-cobre-telmex-inf
                        self.handle_datos_cobre_telmex(row[20])
                    
                    if pd.notna(row[21]) and row[21] > 0:  # Datos-fibra-telmex-inf
                        self.handle_datos_fibra_telmex(row[21])
                    
                    if pd.notna(row[22]) and row[22] > 0:  # TV cable otros
                        self.handle_tv_cable_otros(row[22])
                    
                    if pd.notna(row[23]) and row[23] > 0:  # Dish
                        self.handle_dish(row[23])
                    
                    if pd.notna(row[24]) and row[24] > 0:  # TVS
                        self.handle_tvs(row[24])
                    
                    if pd.notna(row[25]) and row[25] > 0:  # SKY
                        self.handle_sky(row[25])
                    
                    if pd.notna(row[26]) and row[26] > 0:  # VETV
                        self.handle_vetv(row[26])
                    
                    self.click(882, 49)
                    self.sleep(5)
                
                print(f"✅ Servicios línea {self.current_line} completados")
            
            print("🎉 PROCESAMIENTO DE SERVICIOS COMPLETADO EXITOSAMENTE!")
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            raise

    # Funciones de servicios (las mismas que antes)
    def handle_voz_cobre(self, cantidad):
        self.click(100, 114)
        self.sleep(2)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_sdom(self, cantidad):
        self.click(100, 114)
        self.sleep(2)
        self.click(138, 269)
        self.sleep(2)
        self.press_down(2)
        self.click(127, 383)
        self.sleep(2)
        self.write(str(int(cantidad)))
        self.sleep(2)
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_cobre_telmex(self, cantidad):
        self.click(100, 114)
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
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_datos_fibra_telmex(self, cantidad):
        self.click(100, 114)
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
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_tv_cable_otros(self, cantidad):
        self.click(100, 114)
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
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_dish(self, cantidad):
        self.click(100, 114)
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
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_tvs(self, cantidad):
        self.click(100, 114)
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
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_sky(self, cantidad):
        self.click(100, 114)
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
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_vetv(self, cantidad):
        self.click(100, 114)
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
        self.click(82, 423)
        self.sleep(2)
        self.handle_error_click()

    def handle_error_click(self):
        """Manejar clics de error"""
        for _ in range(5):
            self.click(704, 384)
            self.sleep(2)

def clear_screen():
    """Limpia la pantalla de la consola"""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_header():
    """Imprime el encabezado del programa"""
    print("=" * 60)
    print("     CONTROLADOR NSE - SOLO SERVICIOS (PYTHON)")
    print("=" * 60)
    print()

def main():
    # Inicializar automatización
    nse = NSEServicesAutomation()
    nse.is_running = True
    
    clear_screen()
    print_header()
    
    # Verificar dependencias
    print("Verificando dependencias...")
    try:
        import pyautogui
        import pandas as pd
        print("✅ Dependencias verificadas")
    except ImportError as e:
        print(f"❌ Error: {e}")
        print("Instala las dependencias con: pip install pyautogui pandas")
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
    print(f"  - Inicia en: {nse.start_count}")
    print(f"  - Número de lotes: {nse.loop_count}")
    print(f"  - Archivo CSV: {nse.csv_file}")
    print()
    
    # Confirmación final antes de ejecutar
    print("⚠️  ADVERTENCIA: El script SOLO ejecutará la parte de SERVICIOS")
    print("   Asegúrate de que ya se ejecutó el programa principal hasta CERRAR")
    print("   El script comenzará en 3 segundos")
    print("   Presiona Ctrl+C para cancelar")
    print()
    
    try:
        input("Presiona Enter para INICIAR procesamiento de SERVICIOS...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO PROCESAMIENTO DE SERVICIOS...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar script de servicios
        nse.execute_services_script()
        
        print()
        input("Presiona Enter para salir...")
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
        nse.is_running = False
        input("Presiona Enter para salir...")
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        nse.is_running = False
        input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()