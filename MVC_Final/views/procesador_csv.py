import pandas as pd
import pyautogui
import time
import logging
from utils.ahk_managerCopyDelete import AHKManagerCD
from utils.ahk_writer import AHKWriter
from utils.ahk_click_down import AHKClickDown
from utils.ahk_enter import EnterAHKManager
from ..models.datos_globales import DatosGlobales

logger = logging.getLogger(__name__)

class ProcesadorCSV:
    def __init__(self, archivo_csv, estado_global):
        self.archivo_csv = archivo_csv
        self.df = None
        self.ahk_manager = AHKManagerCD()
        self.ahk_writer = AHKWriter()
        self.ahk_click_down = AHKClickDown()
        self.ahk_enter = EnterAHKManager()
        self.estado_global = estado_global
        
    def cargar_csv(self):
        try:
            self.df = pd.read_csv(self.archivo_csv)
            if len(self.df) == 0:
                logger.info("CSV vacío detectado")
                return True
            logger.info(f"CSV cargado correctamente: {len(self.df)} registros")
            return True
        except Exception as e:
            logger.error(f"Error cargando CSV: {e}")
            return False
    
    def iniciar_ahk(self):
        logger.info("Iniciando procesos AHK...")
        time.sleep(1.5)
        return (self.ahk_manager.start_ahk() and 
                self.ahk_writer.start_ahk() and 
                self.ahk_click_down.start_ahk() and
                self.ahk_enter.start_ahk())
    
    def detener_ahk(self):
        logger.info("Deteniendo procesos AHK...")
        self.ahk_manager.stop_ahk()
        self.ahk_writer.stop_ahk()
        self.ahk_click_down.stop_ahk()
        self.ahk_enter.stop_ahk()
        time.sleep(1.5)
    
    def buscar_por_id(self, id_buscar, max_intentos=2):
        if self.df is None or len(self.df) == 0:
            logger.warning("CSV no cargado o vacío")
            return None
            
        for intento in range(1, max_intentos + 1):
            if self.estado_global.esperar_si_pausado():
                return None
                
            resultado = self.df[self.df.iloc[:, 0] == id_buscar]
            
            if len(resultado) > 0:
                logger.info(f"ID {id_buscar} encontrado en intento {intento}")
                return resultado.iloc[0]
            else:
                logger.warning(f"Intento {intento}: ID {id_buscar} no encontrado en el CSV")
                if intento < max_intentos:
                    logger.info(f"Esperando 2 segundos antes de reintentar...")
                    for _ in range(2):
                        if self.estado_global.esperar_si_pausado():
                            return None
                        time.sleep(1)
        
        logger.error(f"ID {id_buscar} no encontrado después de {max_intentos} intentos")
        return None
    
    def procesar_registro(self):
        try:
            logger.info("Paso 1: Click en (83, 266)")
            pyautogui.click(83, 266)
            if self.estado_global.esperar_si_pausado():
                return False, None
            time.sleep(1)
            
            logger.info("Paso 2: Presionando ENTER")
            self.ahk_enter.presionar_enter(1)
            
            logger.info("Paso 2.1: Click en (168, 188)")
            pyautogui.click(168, 188)
            if self.estado_global.esperar_si_pausado():
                return False, None
            time.sleep(0.5)
            
            logger.info("Paso 3: Obteniendo ID con AHKManager en (1483, 519)")
            id_obtenido = self.ahk_manager.ejecutar_acciones_ahk(1483, 519)
            
            if not id_obtenido:
                logger.error("No se pudo obtener el ID")
                return False, None
            
            id_obtenido = int(id_obtenido)
            logger.info(f"ID obtenido: {id_obtenido}")

            logger.info(f"Paso 4: Buscando ID {id_obtenido} en CSV (2 intentos máx)")
            registro = self.buscar_por_id(id_obtenido, max_intentos=2)
            
            if registro is None:
                logger.error(f"ID {id_obtenido} no encontrado en CSV después de 2 intentos. Saltando...")
                return True, None
            
            linea_procesada = None
            for idx in range(len(self.df)):
                if self.df.iloc[idx, 0] == id_obtenido:
                    linea_procesada = idx + 1
                    break
            
            if len(registro) >= 2:
                valor_columna_2 = registro.iloc[1]
                if pd.isna(valor_columna_2):
                    valor_columna_2 = ""
                else:
                    valor_columna_2 = str(valor_columna_2)
                
                logger.info(f"Paso 5: Escribiendo valor '{valor_columna_2}' en (1483, 519)")
                
                if valor_columna_2:
                    exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, valor_columna_2)
                    if not exito_escritura:
                        logger.error("Error en la escritura")
                        return False, linea_procesada
                else:
                    exito_escritura = self.ahk_writer.ejecutar_escritura_ahk(1483, 519, "")
                    if not exito_escritura:
                        logger.error("Error en la escritura")
                        return False, linea_procesada
                    logger.info("Columna 2 vacía, no se escribe nada")
                
                for _ in range(2):
                    if self.estado_global.esperar_si_pausado():
                        return False, linea_procesada
                    time.sleep(1)
            else:
                logger.warning("No hay columna 2 en el registro")
            
            if len(registro) >= 4:
                valor_columna_4 = registro.iloc[3]
                if pd.isna(valor_columna_4):
                    logger.info("Paso 6: Valor columna 4 = ")
                else:
                    logger.info(f"Paso 6: Valor columna 4 = {valor_columna_4}")
                
                if pd.notna(valor_columna_4) and float(valor_columna_4) > 0:
                    veces_down = int(float(valor_columna_4))
                    logger.info(f"Paso 7: Ejecutando {veces_down} veces DOWN en (1507, 636)")
                    
                    exito_down = self.ahk_click_down.ejecutar_click_down(1526, 646, veces_down)
                    if not exito_down:
                        logger.error("Error en click + down")
                        return False, linea_procesada
                    
                    for _ in range(2):
                        if self.estado_global.esperar_si_pausado():
                            return False, linea_procesada
                        time.sleep(1)
                else:
                    logger.info("Paso 7: Saltado (columna 4 <= 0 o vacía)")
            else:
                logger.warning("No hay columna 4 en el registro")
            
            logger.info("Paso 8: Click en (1290, 349)")
            pyautogui.click(1290, 349)
            
            for _ in range(2):
                if self.estado_global.esperar_si_pausado():
                    return False, linea_procesada
                time.sleep(1)
            
            logger.info("Procesamiento completado exitosamente")
            return True, linea_procesada
            
        except Exception as e:
            logger.error(f"Error en procesar_registro: {e}")
            return False, None
    
    def procesar_todo(self):
        if not self.cargar_csv():
            return False, None
            
        if self.df is None or len(self.df) == 0:
            logger.info("CSV vacío. No hay registros para procesar.")
            return True, None
            
        if not self.iniciar_ahk():
            return False, None
        
        try:
            logger.info("Iniciando procesamiento de registro...")
            exito, linea_procesada = self.procesar_registro()
            
            if exito:
                logger.info(f"Procesamiento completado. Línea procesada: {linea_procesada}")
            else:
                logger.error("Procesamiento falló")
                
            return exito, linea_procesada
            
        finally:
            self.detener_ahk()