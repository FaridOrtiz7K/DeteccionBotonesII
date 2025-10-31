import os
import time
import sys
import pandas as pd
import pyautogui
import cv2
import numpy as np
import logging
from utils.ahk_managerCopyDelete import AHKManagerCD
from utils.ahk_writer import AHKWriter
from utils.ahk_click_down import AHKClickDown

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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
                return False
                
            id_obtenido = int(id_obtenido)
            logger.info(f"ID obtenido: {id_obtenido}")
            
            # Paso 4: Buscar el ID en el CSV
            logger.info(f"Paso 4: Buscando ID {id_obtenido} en CSV")
            registro = self.buscar_por_id(id_obtenido)
            
            if registro is None:
                logger.error(f"ID {id_obtenido} no encontrado en CSV")
                return False
            
            # Paso 5: Escribir valor de columna 2 en (1483, 519)
            if len(registro) >= 2:  # Verificar que existe columna 2
                valor_columna_2 = str(registro.iloc[1])
                logger.info(f"Paso 5: Escribiendo valor '{valor_columna_2}' en (1483, 519)")
                
                exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, valor_columna_2)
                if not exito_escritura:
                    logger.error("Error en la escritura")
                    return False
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
                        return False
                else:
                    logger.info("Paso 7: Saltado (columna 4 <= 0)")
            else:
                logger.warning("No hay columna 4 en el registro")
            
            # Paso 8: Click en (1290, 349)
            logger.info("Paso 8: Click en (1290, 349)")
            pyautogui.click(1290, 349)
            time.sleep(1)
            
            logger.info("Procesamiento completado exitosamente")
            return True
            
        except Exception as e:
            logger.error(f"Error en procesar_registro: {e}")
            return False
    
    def procesar_todo(self, pausa_entre_registros=2):
        """Procesa múltiples registros (si es necesario)"""
        if not self.cargar_csv():
            return False
            
        if not self.iniciar_ahk():
            return False
        
        try:
            # Este método procesa un registro por ejecución
            # Si necesitas procesar múltiples registros automáticamente,
            # podemos modificar esta parte
            logger.info("Iniciando procesamiento de registro...")
            exito = self.procesar_registro()
            
            if exito:
                logger.info("Procesamiento completado")
            else:
                logger.error("Procesamiento falló")
                
            return exito
            
        finally:
            # Siempre detener AHK al finalizar
            self.detener_ahk()

class NSEAutomation:
    def __init__(self):
        self.start_count = 4 # Línea específica a procesar (1-indexed)
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

def ejecutar_programa1():
    """Ejecuta el primer programa (ProcesadorCSV)"""
    print("=" * 60)
    print("INICIANDO PROGRAMA 1 - PROCESADOR CSV")
    print("=" * 60)
    
    # Configurar pyautogui
    pyautogui.PAUSE = 0.5
    pyautogui.FAILSAFE = True
    
    # Nombre del archivo CSV - cambia esto por tu archivo real
    archivo_csv = "NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"  # Cambia por el nombre real de tu CSV
    
    # Crear procesador
    procesador = ProcesadorCSV(archivo_csv)
    
    # Ejecutar procesamiento
    print("Iniciando procesamiento...")
    print("Asegúrate de que la ventana objetivo esté activa")
    print("Presiona Ctrl+C para cancelar")
    
    try:
        input("Presiona Enter para comenzar el Programa 1...")
        time.sleep(3)  # Tiempo para cambiar a la ventana correcta
        procesador.procesar_todo()
    except KeyboardInterrupt:
        print("\nProceso cancelado por el usuario")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("Programa 1 finalizado")

def ejecutar_programa2():
    """Ejecuta el segundo programa (NSEAutomation)"""
    print("\n" + "=" * 60)
    print("INICIANDO PROGRAMA 2 - AUTOMATIZACIÓN NSE")
    print("=" * 60)
    
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
        input("Presiona Enter para INICIAR el Programa 2 (Automatización NSE)...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO AUTOMATIZACIÓN NSE ...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar script NSE
        nse.execute_nse_script()
        
        print()
        print("Programa 2 finalizado exitosamente")
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
        nse.is_running = False
    except Exception as e:
        print()
        print(f"❌ Error durante la ejecución: {e}")
        nse.is_running = False

# Función principal combinada
def main():
    """Función principal que ejecuta ambos programas secuencialmente"""
    print("COMBINACIÓN DE PROGRAMAS")
    print("Este script ejecutará ambos programas de forma secuencial")
    
    try:
        # Ejecutar Programa 1
        ejecutar_programa1()
        
        # Pequeña pausa entre programas
        print("\n" + "=" * 60)
        print("TRANSICIÓN ENTRE PROGRAMAS")
        print("=" * 60)
        time.sleep(2)
        
        # Ejecutar Programa 2
        ejecutar_programa2()
        
        print("\n" + "=" * 60)
        print("EJECUCIÓN COMPLETADA - AMBOS PROGRAMAS FINALIZADOS")
        print("=" * 60)
        
    except Exception as e:
        print(f"Error general en la ejecución combinada: {e}")
    finally:
        input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()