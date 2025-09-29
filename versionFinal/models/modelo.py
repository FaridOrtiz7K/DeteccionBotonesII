import pandas as pd
from datetime import datetime
import json
import os
import logging
from typing import Optional, Dict, Any, List

logger = logging.getLogger(__name__)

class ScriptModel:
    def __init__(self, config_file: str = "config/config.json"):
        self.csv_data: Optional[pd.DataFrame] = None
        self.current_line: int = 0
        self.total_lines: int = 0
        self.is_running: bool = False
        self.is_paused: bool = False
        self.start_count: int = 1
        self.loop_count: int = 589
        self.csv_file_path: str = ""
        self.script_log: List[str] = []
        
        # Configuración
        self.config_file = config_file
        self.coords = self._load_config()
        self.credentials_file = "config/credentials.json"
        
        # Estadísticas
        self.stats = {
            'processed_lines': 0,
            'errors': 0,
            'start_time': None,
            'end_time': None
        }
        
        self._ensure_config_directory()
        
    def _ensure_config_directory(self) -> None:
        """Asegura que exista el directorio de configuración"""
        os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
        os.makedirs(os.path.dirname(self.credentials_file), exist_ok=True)
        
    def _load_config(self) -> Dict[str, Any]:
        """Cargar configuración desde archivo JSON"""
        default_coords = {
            'select_list': (89, 263),
            'case_numero': (1483, 519),
            'actualizar': (1290, 349),
            'seleccionar_mapa': (169, 189),
            'asignar_un_nse': (1463, 382),
            'casilla_un_nse': (1266, 590),
            'asignar_varios_nse': (1491, 386),
            'confirmar_nse_u': (1306, 639),
            'confirmar_nse_v': (1648, 752),
            'cerrar_asignar_nse': (1598, 823),
            'administrar_servicios': (1563, 385),
            'detalle_nse': (100, 114),
            'cerrar_administrar_servicios': (882, 49),
            'logo_ge': (39, 55)
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    logger.info(f"Configuración cargada desde: {self.config_file}")
                    return loaded_config
            else:
                # Crear archivo de configuración por defecto
                self._save_config(default_coords)
                logger.info(f"Archivo de configuración creado: {self.config_file}")
                return default_coords
                
        except Exception as e:
            logger.error(f"Error cargando configuración: {str(e)}")
            return default_coords
    
    def _save_config(self, config_data: Dict[str, Any]) -> bool:
        """Guardar configuración en archivo JSON"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=4, ensure_ascii=False)
            logger.info(f"Configuración guardada en: {self.config_file}")
            return True
        except Exception as e:
            logger.error(f"Error guardando configuración: {str(e)}")
            return False
    
    def update_coordinate(self, key: str, x: int, y: int) -> bool:
        """Actualizar una coordenada específica"""
        if key in self.coords:
            self.coords[key] = (x, y)
            return self._save_config(self.coords)
        else:
            logger.warning(f"Intento de actualizar coordenada inexistente: {key}")
            return False
    
    def get_coordinate(self, key: str) -> Optional[tuple]:
        """Obtener una coordenada específica"""
        return self.coords.get(key)
    
    def load_csv(self, file_path: str) -> bool:
        """Cargar archivo CSV con validaciones"""
        try:
            # Verificar que el archivo existe
            if not os.path.exists(file_path):
                self.add_log(f"Error: El archivo {file_path} no existe")
                return False
            
            # Verificar extensión
            if not file_path.lower().endswith('.csv'):
                self.add_log("Error: El archivo debe ser un CSV")
                return False
            
            # Cargar CSV
            self.csv_data = pd.read_csv(file_path)
            self.total_lines = len(self.csv_data)
            self.csv_file_path = file_path
            
            # Validar estructura básica del CSV
            if self.total_lines == 0:
                self.add_log("Error: El archivo CSV está vacío")
                self.csv_data = None
                return False
            
            # Verificar columnas mínimas requeridas (ajustar según necesidades)
            required_columns = ['B', 'D']  # Columnas mencionadas en process_row
            actual_columns = self.csv_data.columns.tolist()
            
            self.add_log(f"CSV cargado exitosamente: {file_path}")
            self.add_log(f"Total de líneas: {self.total_lines}")
            self.add_log(f"Columnas detectadas: {', '.join(actual_columns)}")
            
            return True
            
        except pd.errors.EmptyDataError:
            self.add_log("Error: El archivo CSV está vacío")
            return False
        except pd.errors.ParserError as e:
            self.add_log(f"Error de formato en CSV: {str(e)}")
            return False
        except Exception as e:
            self.add_log(f"Error inesperado cargando CSV: {str(e)}")
            return False

    def reset(self) -> None:
        """Reiniciar el estado del modelo"""
        self.current_line = 0
        self.is_running = False
        self.is_paused = False
        self.stats = {
            'processed_lines': 0,
            'errors': 0,
            'start_time': None,
            'end_time': None
        }
        self.add_log("Script reiniciado")

    def get_current_row(self) -> Optional[pd.Series]:
        """Obtener la fila actual del CSV"""
        if self.csv_data is not None and self.current_line < self.total_lines:
            return self.csv_data.iloc[self.current_line]
        return None

    def set_start_parameters(self, start_count: int, loop_count: int) -> bool:
        """Establecer parámetros de inicio con validación"""
        try:
            # Validar parámetros
            if start_count < 1:
                self.add_log("Error: El inicio debe ser al menos 1")
                return False
            
            if loop_count < 1:
                self.add_log("Error: El número de lotes debe ser al menos 1")
                return False
            
            # Ajustar si excede el total de líneas
            actual_start = min(start_count, self.total_lines)
            actual_loop = min(loop_count, self.total_lines - actual_start + 1)
            
            self.start_count = actual_start
            self.loop_count = actual_loop
            self.current_line = max(0, actual_start - 1)
            
            self.add_log(f"Parámetros establecidos: inicio={actual_start}, lotes={actual_loop}")
            
            if actual_start != start_count or actual_loop != loop_count:
                self.add_log("Nota: Los parámetros fueron ajustados para no exceder el total de líneas")
            
            return True
            
        except Exception as e:
            self.add_log(f"Error estableciendo parámetros: {str(e)}")
            return False

    def add_log(self, message: str) -> None:
        """Agregar mensaje al log con timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.script_log.append(log_entry)
        
        # Mantener un límite de logs en memoria
        if len(self.script_log) > 1000:
            self.script_log = self.script_log[-500:]  # Mantener últimos 500
        
        # También loggear con el sistema de logging
        logger.info(message)

    def get_recent_logs(self, count: int = 10) -> List[str]:
        """Obtener logs recientes"""
        return self.script_log[-count:] if self.script_log else []

    def clear_logs(self) -> None:
        """Limpiar todos los logs"""
        self.script_log.clear()
        self.add_log("Logs limpiados")

    def validate_credentials(self, username: str, password: str) -> bool:
        """Validar credenciales de usuario"""
        try:
            # Primero intentar cargar desde archivo
            if os.path.exists(self.credentials_file):
                with open(self.credentials_file, 'r', encoding='utf-8') as f:
                    credentials_data = json.load(f)
                    stored_users = credentials_data.get('users', {})
                    
                    if username in stored_users and stored_users[username] == password:
                        self.add_log(f"Usuario autenticado: {username}")
                        return True
            
            # Fallback a usuarios por defecto
            valid_users = {
                "admin": "123",
                "user1": "password1",
                "user2": "password2"
            }
            
            if username in valid_users and valid_users[username] == password:
                self.add_log(f"Usuario autenticado (default): {username}")
                return True
                
            self.add_log(f"Intento de autenticación fallido para usuario: {username}")
            return False
            
        except Exception as e:
            logger.error(f"Error validando credenciales: {str(e)}")
            return False

    def update_credentials(self, username: str, password: str) -> bool:
        """Actualizar o agregar credenciales de usuario"""
        try:
            credentials_data = {'users': {}}
            
            # Cargar credenciales existentes si el archivo existe
            if os.path.exists(self.credentials_file):
                with open(self.credentials_file, 'r', encoding='utf-8') as f:
                    credentials_data = json.load(f)
            
            # Actualizar usuario
            credentials_data['users'][username] = password
            
            # Guardar
            with open(self.credentials_file, 'w', encoding='utf-8') as f:
                json.dump(credentials_data, f, indent=4, ensure_ascii=False)
            
            self.add_log(f"Credenciales actualizadas para usuario: {username}")
            return True
            
        except Exception as e:
            logger.error(f"Error actualizando credenciales: {str(e)}")
            return False

    def start_processing(self) -> None:
        """Iniciar el procesamiento (actualizar estadísticas)"""
        self.stats = {
            'processed_lines': 0,
            'errors': 0,
            'start_time': datetime.now(),
            'end_time': None
        }
        self.add_log("Procesamiento iniciado")

    def finish_processing(self) -> None:
        """Finalizar el procesamiento (actualizar estadísticas)"""
        self.stats['end_time'] = datetime.now()
        duration = self.stats['end_time'] - self.stats['start_time']
        self.add_log(f"Procesamiento finalizado. Duración: {duration}")

    def increment_processed(self) -> None:
        """Incrementar contador de líneas procesadas"""
        self.stats['processed_lines'] += 1

    def increment_errors(self) -> None:
        """Incrementar contador de errores"""
        self.stats['errors'] += 1

    def get_progress(self) -> Dict[str, Any]:
        """Obtener progreso actual del procesamiento"""
        if self.total_lines == 0:
            return {'percentage': 0, 'current': 0, 'total': 0}
        
        percentage = (self.current_line / self.total_lines) * 100
        return {
            'percentage': round(percentage, 2),
            'current': self.current_line,
            'total': self.total_lines,
            'processed': self.stats['processed_lines'],
            'errors': self.stats['errors']
        }

    def get_processing_time(self) -> Optional[str]:
        """Obtener tiempo transcurrido de procesamiento"""
        if self.stats['start_time'] and not self.stats['end_time']:
            duration = datetime.now() - self.stats['start_time']
            return str(duration).split('.')[0]  # Remover microsegundos
        elif self.stats['start_time'] and self.stats['end_time']:
            duration = self.stats['end_time'] - self.stats['start_time']
            return str(duration).split('.')[0]
        return None

    def export_logs(self, file_path: str) -> bool:
        """Exportar logs a un archivo de texto"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("Log de Ejecución - Script Automatización SIGP\n")
                f.write("=" * 50 + "\n\n")
                for log_entry in self.script_log:
                    f.write(log_entry + "\n")
            
            self.add_log(f"Logs exportados a: {file_path}")
            return True
        except Exception as e:
            self.add_log(f"Error exportando logs: {str(e)}")
            return False

    def validate_csv_structure(self) -> Dict[str, Any]:
        """Validar la estructura del CSV cargado"""
        if self.csv_data is None:
            return {'valid': False, 'message': 'No hay CSV cargado'}
        
        try:
            issues = []
            
            # Verificar valores nulos en columnas críticas
            critical_columns = [1, 3]  # Columnas B y D (índices 1 y 3)
            for col_idx in critical_columns:
                if col_idx < len(self.csv_data.columns):
                    null_count = self.csv_data.iloc[:, col_idx].isnull().sum()
                    if null_count > 0:
                        issues.append(f"Columna {self.csv_data.columns[col_idx]}: {null_count} valores nulos")
            
            # Verificar tipos de datos
            for col_idx in [3]:  # Columna D debería ser numérica
                if col_idx < len(self.csv_data.columns):
                    try:
                        pd.to_numeric(self.csv_data.iloc[:, col_idx], errors='raise')
                    except:
                        issues.append(f"Columna {self.csv_data.columns[col_idx]}: valores no numéricos")
            
            return {
                'valid': len(issues) == 0,
                'message': 'Estructura válida' if len(issues) == 0 else 'Problemas encontrados',
                'issues': issues,
                'summary': {
                    'total_rows': len(self.csv_data),
                    'total_columns': len(self.csv_data.columns),
                    'columns': self.csv_data.columns.tolist()
                }
            }
            
        except Exception as e:
            return {'valid': False, 'message': f'Error en validación: {str(e)}'}

    def __str__(self) -> str:
        """Representación en string del estado del modelo"""
        progress = self.get_progress()
        return (f"ScriptModel: "
                f"Running={self.is_running}, "
                f"Paused={self.is_paused}, "
                f"Progress={progress['percentage']}%, "
                f"Line={self.current_line + 1}/{self.total_lines}")