import os
import time
import sys
import pyautogui
import pandas as pd
import cv2
import numpy as np
from utils.ahk_writer import AHKWriter

class GEAutomation:
    def __init__(self):
        self.csv_file = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        self.reference_image = "img/textoAdicional.PNG"  # Imagen para detectar ventana de texto adicional
        self.is_running = False
        
        # Inicializar AHKWriter
        self.ahk_writer = AHKWriter()
        
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # COORDENADAS ABSOLUTAS (ajustar según tu resolución)
        self.coords = {
            'agregar_ruta': (327, 381),
            'archivo': (1396, 608),
            'abrir': (1406, 634),
            'documents': (1120, 666),
            'cargar_ruta': (1406, 675),
            'lote': (70, 266),
            'seleccionar_mapa': (168, 188),
            'anotar': (1366, 384),
            'agregar_texto_adicional': (1449, 452),
            'campo_texto': (1421, 526),  # Cambia con respecto a la deteccion de la imagen
            'agregar_texto': (1263, 572), # Cambia con respecto a la deteccion de la imagen
            'cerrar_ventana_texto': (1338, 570),#Cambia con respecto a la deteccion de la imagen
            'limpiar_trazo': (360, 980),
            'lote_again': (70, 266)
        }
        
        # Coordenadas relativas para detección de imagen (campo de texto)
        self.coords_texto_relativas = {
            'campo_texto': (222 , 54),  # Se ajustará según la detección
            'agregar_texto': (64 , 100),
            'cerrar_ventana_texto': (139 , 98)
        }

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
            return True, max_loc  # Devuelve las coordenadas (x, y) de la esquina superior izquierda
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

    def perform_actions(self):
        """Función principal que realiza todas las acciones"""
        # Iniciar AHKWriter
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

            # Realizar el bucle 9 veces (como en el código original)
            for iteration in range(1, 10):
                if iteration > total_lines:
                    print(f"⚠️  No hay más líneas en el CSV. Iteración {iteration} saltada.")
                    continue
                    
                print(f"🔄 Procesando iteración {iteration}/9")
                self.process_single_iteration(df, iteration, total_lines)
                
                # Guardar cada 10 iteraciones (en este caso, solo al final del bucle)
                if iteration % 10 == 0:
                    self.save_progress()
                    
            # Guardar al final
            self.save_progress()
            print("✅ Script completado exitosamente!")
            return True
            
        except Exception as e:
            print(f"❌ Error durante la ejecución: {e}")
            return False
        finally:
            # Detener AHKWriter
            self.ahk_writer.stop_ahk()

    def process_single_iteration(self, df, iteration, total_lines):
        """Procesar una sola iteración del bucle"""
        # Obtener la fila correspondiente (0-indexed)
        row_index = iteration - 1
        row = df.iloc[row_index]
        
        # Obtener valores del CSV
        if len(row) >= 2:
            num_txt_type = str(row[0])  # Primera columna para el nombre del archivo
            texto_adicional = str(row[1])  # Segunda columna para el texto adicional
        else:
            print(f"⚠️  Fila {iteration} no tiene suficientes columnas")
            return

        # SECUENCIA DE ACCIONES
        try:
            # 1. Seleccionar Agregar ruta de GE
            self.click(*self.coords['agregar_ruta'])
            self.sleep(2)
            
            # 2. Clic en Archivo
            self.click(*self.coords['archivo'])
            self.sleep(3)
            
            # 3. Clic en Abrir
            self.click(*self.coords['abrir'])
            self.sleep(3)
            
            # 4. Clic en Documents
            self.click(*self.coords['documents'])
            self.sleep(3)
            
            
            # 6. Escribir nombre del archivo KML con AHKWriter
            nombre_archivo = f"RA {num_txt_type}.kml"
            self.write_with_ahk(self.coords['documents'][0], self.coords['documents'][1], nombre_archivo)
            self.sleep(1)
            
            # 7. Presionar Enter
            pyautogui.press('enter')
            self.sleep(3)
            
            # 8. Seleccionar Agregar ruta de GE nuevamente
            self.click(*self.coords['agregar_ruta'])
            self.sleep(2)
            
            # 9. Cargar ruta
            self.click(*self.coords['cargar_ruta'])
            self.sleep(2)
            
            # 10. Seleccionar Lote
            self.click(*self.coords['lote'])
            self.sleep(2)
            
            # 11. Seleccionar en el mapa
            self.click(*self.coords['seleccionar_mapa'])
            self.sleep(2)
            
            # 12. Anotar
            self.click(*self.coords['anotar'])
            self.sleep(2)
            
            # 13. Agregar texto adicional
            self.click(*self.coords['agregar_texto_adicional'])
            self.sleep(2)
            
            # 14. DETECCIÓN DE IMAGEN para el campo de texto
            image_found, base_location = self.wait_for_image_with_retries(self.reference_image, max_attempts=10)
            
            if image_found:
                # Calcular coordenadas absolutas del campo de texto basado en la detección
                x_campo = base_location[0] + self.coords_texto_relativas['campo_texto'][0]
                y_campo = base_location[1] + self.coords_texto_relativas['campo_texto'][1]
                x_agregar = base_location[0] + self.coords_texto_relativas['agregar_texto'][0]
                y_agregar = base_location[1] + self.coords_texto_relativas['agregar_texto'][1]
                x_cerrar = base_location[0] + self.coords_texto_relativas['cerrar_ventana_texto'][0]
                y_cerrar = base_location[1] + self.coords_texto_relativas['cerrar_ventana_texto'][1]
                # Hacer clic en el campo de texto detectado
                self.click(x_campo, y_campo)
                self.sleep(2)
                
                # Escribir texto adicional con AHKWriter
                self.write_with_ahk(x_campo, y_campo, texto_adicional)
                self.sleep(1)
                # 15. Agregar texto
                self.click(x_agregar, y_agregar)
                self.sleep(3)
            
                # 16. Cerrar ventana de texto
                self.click(x_cerrar, y_cerrar)
                self.sleep(2)
            else:
                # Fallback a coordenadas fijas si no se detecta la imagen
                print("⚠️  Usando coordenadas fijas para campo de texto")
                self.click(*self.coords['campo_texto'])
                self.sleep(2)
                pyautogui.press('backspace')
                self.sleep(1)
                self.write_with_ahk(*self.coords['campo_texto'], texto_adicional)
                self.sleep(1)
                # 15. Agregar texto
                self.click(*self.coords['agregar_texto'])
                self.sleep(3)
            
                # 16. Cerrar ventana de texto
                self.click(*self.coords['cerrar_ventana_texto'])
                self.sleep(2)
            
            
            # 17. Limpiar trazo
            self.click(*self.coords['limpiar_trazo'])
            self.sleep(1)
            
            # 18. Seleccionar Lote nuevamente
            self.click(*self.coords['lote_again'])
            self.sleep(2)
            
            # 19. Presionar flecha abajo
            pyautogui.press('down')
            self.sleep(2)
            
            print(f"✅ Iteración {iteration} completada")
            
        except Exception as e:
            print(f"❌ Error en iteración {iteration}: {e}")

    def save_progress(self):
        """Guardar progreso con Ctrl + S"""
        print("💾 Guardando progreso...")
        pyautogui.hotkey('ctrl', 's')
        self.sleep(6)

def main():
    # Inicializar automatización
    ge_auto = GEAutomation()
    ge_auto.is_running = True
    
    # Verificar archivo CSV
    if not os.path.exists(ge_auto.csv_file):
        print(f"❌ ERROR: Archivo CSV no encontrado: {ge_auto.csv_file}")
        input("Presiona Enter para salir...")
        return
    
    print(f"✅ Archivo CSV encontrado: {ge_auto.csv_file}")
    
    # Verificar imagen de referencia
    if not os.path.exists(ge_auto.reference_image):
        print(f"⚠️  Advertencia: Imagen de referencia no encontrada: {ge_auto.reference_image}")
        print("   El proceso usará coordenadas fijas como fallback")
    else:
        print(f"✅ Imagen de referencia encontrada: {ge_auto.reference_image}")
    
    print()
    print("Configuración:")
    print(f"  - Archivo CSV: {ge_auto.csv_file}")
    print(f"  - Imagen de referencia: {ge_auto.reference_image}")
    print("  - Usando AHKWriter para escritura")
    print("  - 9 iteraciones programadas")
    print()
    
    try:
        input("Presiona Enter para INICIAR la automatización...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO AUTOMATIZACIÓN ...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar script
        success = ge_auto.perform_actions()
        
        if success:
            print("🎉 AUTOMATIZACIÓN COMPLETADA EXITOSAMENTE!")
        else:
            print("❌ La automatización encontró errores")
        
        print()
        input("Presiona Enter para salir...")
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
        ge_auto.is_running = False
        input("Presiona Enter para salir...")
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        ge_auto.is_running = False
        input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()