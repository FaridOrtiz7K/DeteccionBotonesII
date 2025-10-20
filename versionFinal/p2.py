import os
import time
import sys
import pyautogui
import pandas as pd

class NSEAutomation:
    def __init__(self):
        self.start_count = 1
        self.loop_count = 589
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.reference_image = "asignar_nse_reference.png"
        self.current_line = 0
        self.is_running = False
        
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # COORDENADAS RELATIVAS (de la tabla verde)
        self.coords_select = {
            6: [33, 92], 7: [33, 131], 8: [33, 159], 9: [33, 197],
            10: [33, 231], 11: [398, 92], 12: [398, 131], 13: [398, 159],
            14: [33, 301], 15: [33, 333], 16: [33, 367]
        }
        
        self.coords_type = {
            6: [163, 92], 7: [163, 131], 8: [163, 159], 9: [163, 197],
            10: [163, 231], 11: [528, 92], 12: [528, 131], 13: [528, 159],
            14: [163, 301], 15: [163, 333], 16: [163, 367]
        }
        
        # Coordenadas para botones (relativas)
        self.coords_asignar = [446, 281]
        self.coords_cerrar = [396, 352]
        
        # Coordenadas absolutas para tipo U
        self.coords_select_u = {
            6: [1268, 637], 7: [1268, 661], 8: [1268, 685], 9: [1268, 709],
            10: [1268, 733], 11: [1268, 757], 12: [1268, 781], 13: [1268, 825],
            14: [1268, 856], 15: [1268, 881], 16: [1268, 908]
        }

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

    def press_delete(self):
        """Presionar Delete"""
        pyautogui.press('delete')
        time.sleep(0.5)

    def sleep(self, seconds):
        """Esperar segundos"""
        time.sleep(seconds)

    def detect_image(self, image_path, confidence=0.7):
        """Detectar imagen en pantalla"""
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if location:
                return True, location
            return False, None
        except Exception as e:
            print(f"Error en detección de imagen: {e}")
            return False, None

    def wait_for_image(self, image_path, timeout=10, confidence=0.7):
        """Esperar a que aparezca una imagen"""
        print(f"🔍 Buscando imagen: {image_path}")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            found, location = self.detect_image(image_path, confidence)
            if found:
                print("✅ Imagen detectada")
                return True, location
            time.sleep(1)
        
        print("❌ Imagen no encontrada")
        return False, None

    def execute_nse_script(self):
        """Función principal de ejecución NSE - Solo hasta CERRAR"""
        try:
            # Leer CSV
            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            for i in range(self.start_count - 1, min(self.start_count + self.loop_count - 1, total_lines)):
                if not self.is_running:
                    break
                    
                self.current_line = i + 1
                print(f"🔄 Procesando línea {self.current_line}/{total_lines}")
                
                row = df.iloc[i]
                
                # Proceso principal
                self.click(89, 263)
                self.sleep(1.5)
                self.click(1483, 519)
                self.sleep(1.5)
                self.press_delete()
                self.sleep(1)
                self.write(str(row[1]))  # Columna 2
                self.sleep(1.5)
                
                # Manejar columna D
                if pd.notna(row[3]) and row[3] > 0:
                    self.click(1507, 650)
                    self.sleep(2.5)
                    self.press_down(int(row[3]))
                
                self.click(1290, 349)
                self.sleep(1.5)
                
                # Paso 2 - Abrir botón para asignar NSE
                self.click(1648, 752)  # Clic en ASIGNAR
                self.sleep(3)
                
                # ESPACIO PARA DETECCIÓN DE IMAGEN
                image_found, location = self.wait_for_image(self.reference_image, timeout=10)
                
                if image_found:
                    print("🎯 Usando coordenadas relativas")
                    # Manejar tipo U o V con coordenadas relativas
                    if str(row[4]).strip().upper() == "U":
                        self.handle_type_u(row)
                    elif str(row[4]).strip().upper() == "V":
                        self.handle_type_v(row)
                else:
                    print("⚠️ Usando coordenadas absolutas por fallo en detección")
                    # Manejar tipo U o V con coordenadas absolutas
                    if str(row[4]).strip().upper() == "U":
                        self.handle_type_u_absolute(row)
                    elif str(row[4]).strip().upper() == "V":
                        self.handle_type_v_absolute(row)
                
                # FINALIZAR ITERACIÓN (solo hasta CERRAR)
                self.click(89, 263)
                self.sleep(3)
                self.press_down()
                self.sleep(3)
                
                print(f"✅ Línea {self.current_line} completada (hasta CERRAR)")
            
            # Proceso posterior al bucle
            self.click(39, 55)
            print("🎉 AUTOMATIZACIÓN HASTA CERRAR COMPLETADA EXITOSAMENTE!")
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            raise

    def handle_type_u(self, row):
        """Manejar tipo U con coordenadas relativas"""
        self.click(169, 189)
        self.sleep(2)
        self.click(1463, 382)
        self.sleep(2)
        self.click(1266, 590)
        self.sleep(2)
        
        # Lógica U para columnas 6-16
        for col_index in range(6, 17):
            if pd.notna(row[col_index-1]) and row[col_index-1] > 0:
                x, y = self.coords_select_u[col_index]
                self.click(x, y)
                self.sleep(3)
        
        self.click(1306, 639)
        self.sleep(2)

    def handle_type_v(self, row):
        """Manejar tipo V con coordenadas relativas"""
        self.click(169, 189)
        self.sleep(3)
        self.click(1491, 386)
        self.sleep(3)
        
        # Lógica V para columnas 6-16 con coordenadas relativas
        for col_index in range(6, 17):
            if pd.notna(row[col_index-1]) and row[col_index-1] > 0:
                # Usar coordenadas relativas de la tabla verde
                x_cs, y_cs = self.coords_select[col_index]
                x_ct, y_ct = self.coords_type[col_index]
                
                self.click(x_cs, y_cs)
                self.sleep(2)
                self.click(x_ct, y_ct)
                self.sleep(2)
                self.write(str(int(row[col_index-1])))
                self.sleep(2)
        
        # Usar coordenadas relativas para CERRAR
        x_cerrar, y_cerrar = self.coords_cerrar
        self.click(x_cerrar, y_cerrar)
        self.sleep(2)

    def handle_type_u_absolute(self, row):
        """Manejar tipo U con coordenadas absolutas (fallback)"""
        self.handle_type_u(row)  # Ya usa coordenadas absolutas

    def handle_type_v_absolute(self, row):
        """Manejar tipo V con coordenadas absolutas (fallback)"""
        self.click(169, 189)
        self.sleep(3)
        self.click(1491, 386)
        self.sleep(3)
        
        # Coordenadas absolutas originales para tipo V
        coords_select_abs = {
            6: [1235, 563], 7: [1235, 602], 8: [1235, 630], 9: [1235, 668],
            10: [1235, 702], 11: [1600, 563], 12: [1600, 602], 13: [1600, 630],
            14: [1235, 772], 15: [1235, 804], 16: [1235, 838]
        }
        
        coords_type_abs = {
            6: [1365, 563], 7: [1365, 602], 8: [1365, 630], 9: [1365, 668],
            10: [1365, 702], 11: [1730, 563], 12: [1730, 602], 13: [1730, 630],
            14: [1365, 772], 15: [1365, 804], 16: [1365, 838]
        }
        
        for col_index in range(6, 17):
            if pd.notna(row[col_index-1]) and row[col_index-1] > 0:
                x_cs, y_cs = coords_select_abs[col_index]
                x_ct, y_ct = coords_type_abs[col_index]
                
                self.click(x_cs, y_cs)
                self.sleep(2)
                self.click(x_ct, y_ct)
                self.sleep(2)
                self.write(str(int(row[col_index-1])))
                self.sleep(2)
        
        self.click(1598, 823)
        self.sleep(2)

def clear_screen():
    """Limpia la pantalla de la consola"""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_header():
    """Imprime el encabezado del programa"""
    print("=" * 60)
    print("     CONTROLADOR NSE - SOLO HASTA CERRAR (PYTHON)")
    print("=" * 60)
    print()

def main():
    # Inicializar automatización
    nse = NSEAutomation()
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
        print("Instala las dependencias con: pip install pyautogui pandas opencv-python pillow")
        input("Presiona Enter para salir...")
        return
    
    # Verificar archivo CSV
    if not os.path.exists(nse.csv_file):
        print(f"❌ ERROR: Archivo CSV no encontrado: {nse.csv_file}")
        input("Presiona Enter para salir...")
        return
    
    print(f"✅ Archivo CSV encontrado: {nse.csv_file}")
    
    # Verificar imagen de referencia
    if not os.path.exists(nse.reference_image):
        print(f"⚠️  Advertencia: Imagen de referencia no encontrada: {nse.reference_image}")
        print("   El programa usará coordenadas absolutas como fallback")
    else:
        print(f"✅ Imagen de referencia encontrada: {nse.reference_image}")
    
    print()
    print("Configuración:")
    print(f"  - Inicia en: {nse.start_count}")
    print(f"  - Número de lotes: {nse.loop_count}")
    print(f"  - Archivo CSV: {nse.csv_file}")
    print(f"  - Imagen de referencia: {nse.reference_image}")
    print()
    
    # Confirmación final antes de ejecutar
    print("⚠️  ADVERTENCIA: El script SOLO ejecutará hasta CERRAR")
    print("   No procesará la sección de servicios")
    print("   El script comenzará en 3 segundos")
    print("   Asegúrate de que la ventana de destino esté activa")
    print("   Presiona Ctrl+C para cancelar")
    print()
    
    try:
        input("Presiona Enter para INICIAR la automatización NSE...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO AUTOMATIZACIÓN NSE (SOLO HASTA CERRAR)...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar script NSE
        nse.execute_nse_script()
        
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