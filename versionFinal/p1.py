import pandas as pd
import pyautogui
import time
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

# Función principal
def main():
    # Configurar pyautogui
    pyautogui.PAUSE = 0.5
    pyautogui.FAILSAFE = True
    
    # Nombre del archivo CSV - cambia esto por tu archivo real
    archivo_csv = "p1.csv"  # Cambia por el nombre real de tu CSV
    
    # Crear procesador
    procesador = ProcesadorCSV(archivo_csv)
    
    # Ejecutar procesamiento
    print("Iniciando procesamiento...")
    print("Asegúrate de que la ventana objetivo esté activa")
    print("Presiona Ctrl+C para cancelar")
    
    try:
        input("Presiona Enter para comenzar...")
        time.sleep(3)  # Tiempo para cambiar a la ventana correcta
        procesador.procesar_todo()
    except KeyboardInterrupt:
        print("\nProceso cancelado por el usuario")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("Proceso finalizado")

if __name__ == "__main__":
    main()