import os
import time
import sys
import pyautogui
import pandas as pd
import cv2
import numpy as np
from utils.ahk_writer import AHKWriter

class NSEAutomation:
    def __init__(self):
        self.start_count = 1
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.reference_image = "asignar_nse_reference.png"
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
            return
            
        try:
            # Leer CSV
            df = pd.read_csv(self.csv_file)
            total_lines = len(df)
            
            print(f"📊 Total de líneas en CSV: {total_lines}")
            
            # Procesar solo la línea específica (start_count - 1)
            if self.start_count - 1 >= total_lines:
                print(f"❌ Error: No existe la línea {self.start_count} en el CSV")
                return
                
            row = df.iloc[self.start_count - 1]
            print(f"🔄 Procesando línea {self.start_count}/{total_lines}")
            
            # Verificar si se debe saltar el proceso (columna 6 tiene valor)
            if self.should_skip_process(row):
                print(f"⏭️  Saltando línea {self.start_count} - Columna 6 tiene valor: {row[5]}")
                return
            
            # Verificar que sea tipo V
            if str(row[4]).strip().upper() != "V":
                print(f"⚠️  Saltando línea {self.start_count} - No es tipo V: {row[4]}")
                return
            
            # Paso directo - Abrir botón para asignar NSE
            self.click(1648, 752)  # Clic en ASIGNAR
            self.sleep(3)
            
            # ESPACIO PARA DETECCIÓN DE IMAGEN CON REINTENTOS
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=30)
            
            if not image_found:
                print("❌ No se puede continuar sin detectar la imagen de referencia.")
                return
            
            # Si se encontró la imagen, continuar con el proceso usando las coordenadas base
            print("🎯 Imagen detectada, procediendo con tipo V")
            self.handle_type_v(row, base_location)
            
            print(f"✅ Línea {self.start_count} completada (hasta CERRAR)")
            print("🎉 AUTOMATIZACIÓN COMPLETADA EXITOSAMENTE!")
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            raise
        finally:
            # Detener AHKWriter
            self.ahk_writer.stop_ahk()

    def handle_type_v(self, row, base_location):
        """Manejar tipo V con coordenadas relativas - COLUMNAS 7-17"""
        # Calcular coordenadas absolutas sumando las relativas a la posición base
        base_x, base_y = base_location
        
        # Coordenadas absolutas para los primeros clics
        click_1_x = base_x + 169
        click_1_y = base_y + 189
        click_2_x = base_x + 1491
        click_2_y = base_y + 386
        
        self.click(click_1_x, click_1_y)
        self.sleep(3)
        self.click(click_2_x, click_2_y)
        self.sleep(3)
        
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

def main():
    # Inicializar automatización
    nse = NSEAutomation()
    nse.is_running = True
    
    # Verificar archivo CSV
    if not os.path.exists(nse.csv_file):
        print(f"❌ ERROR: Archivo CSV no encontrado: {nse.csv_file}")
        input("Presiona Enter para salir...")
        return
    
    print(f"✅ Archivo CSV encontrado: {nse.csv_file}")
    
    # Verificar imagen de referencia
    if not os.path.exists(nse.reference_image):
        print(f"⚠️  Advertencia: Imagen de referencia no encontrada: {nse.reference_image}")
        print("   El proceso se detendrá si no puede encontrar la imagen después de 30 intentos")
    else:
        print(f"✅ Imagen de referencia encontrada: {nse.reference_image}")
    
    print()
    print("Configuración:")
    print(f"  - Línea a procesar: {nse.start_count}")
    print(f"  - Archivo CSV: {nse.csv_file}")
    print(f"  - Imagen de referencia: {nse.reference_image}")
    print("  - Usando AHKWriter para escritura")
    print()
    
    try:
        input("Presiona Enter para INICIAR la automatización NSE...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO AUTOMATIZACIÓN NSE (PROCESO ÚNICO)...")
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