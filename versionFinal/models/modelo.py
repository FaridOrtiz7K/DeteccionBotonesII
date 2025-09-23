import pandas as pd
import time
from datetime import datetime

class ScriptModel:
    def __init__(self):
        self.csv_data = None
        self.current_line = 0
        self.total_lines = 0
        self.is_running = False
        self.is_paused = False
        self.start_count = 1
        self.loop_count = 589
        self.csv_file_path = ""
        self.script_log = []
        
        # Configuración de coordenadas (estas deberían ajustarse según la resolución)
        self.coords = {
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
        
        self.coords_select_u = {
            6: (1268, 637), 7: (1268, 661), 8: (1268, 685), 9: (1268, 709),
            10: (1268, 733), 11: (1268, 757), 12: (1268, 781), 13: (1268, 825),
            14: (1268, 856), 15: (1268, 881), 16: (1268, 908)
        }
        
        self.coords_select_v = {
            6: (1235, 563), 7: (1235, 602), 8: (1235, 630), 9: (1235, 668),
            10: (1235, 702), 11: (1600, 563), 12: (1600, 602), 13: (1600, 630),
            14: (1235, 772), 15: (1235, 804), 16: (1235, 838)
        }
        
        self.coords_type_v = {
            6: (1365, 563), 7: (1365, 602), 8: (1365, 630), 9: (1365, 668),
            10: (1365, 702), 11: (1730, 563), 12: (1730, 602), 13: (1730, 630),
            14: (1365, 772), 15: (1365, 804), 16: (1365, 838)
        }

    def load_csv(self, file_path):
        try:
            self.csv_data = pd.read_csv(file_path)
            self.total_lines = len(self.csv_data)
            self.csv_file_path = file_path
            self.add_log(f"CSV cargado: {file_path} - {self.total_lines} líneas")
            return True
        except Exception as e:
            self.add_log(f"Error cargando CSV: {str(e)}")
            return False

    def reset(self):
        self.current_line = 0
        self.is_running = False
        self.is_paused = False
        self.add_log("Script reiniciado")

    def get_current_row(self):
        if self.csv_data is not None and self.current_line < self.total_lines:
            return self.csv_data.iloc[self.current_line]
        return None

    def set_start_parameters(self, start_count, loop_count):
        self.start_count = start_count
        self.loop_count = loop_count
        self.current_line = max(0, start_count - 1)
        self.add_log(f"Parámetros establecidos: inicio={start_count}, lotes={loop_count}")

    def add_log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.script_log.append(log_entry)
        # Mantener solo los últimos 100 logs
        if len(self.script_log) > 100:
            self.script_log.pop(0)

    def get_recent_logs(self, count=10):
        return self.script_log[-count:] if self.script_log else []

    def validate_credentials(self, username, password):
        # Aquí puedes implementar una validación más segura
        valid_users = {
            "admin": "admin",
            "user": "password"
        }
        return username in valid_users and valid_users[username] == password