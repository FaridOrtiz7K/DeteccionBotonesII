import pandas as pd
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
        
        # Configuración de coordenadas
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
        if len(self.script_log) > 100:
            self.script_log.pop(0)

    def get_recent_logs(self, count=10):
        return self.script_log[-count:] if self.script_log else []

    def validate_credentials(self, username, password):
        valid_users = {
            "admin": "123",
            "user1": "password1",
            "user2": "password2"
        }
        return username in valid_users and valid_users[username] == password