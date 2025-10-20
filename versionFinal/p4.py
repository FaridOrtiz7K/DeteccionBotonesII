import os
import time
import sys

class AHKController:
    def __init__(self):
        self.command_file = "ahk_command.txt"
        self.done_file = "ahk_done.txt"
        self.response_file = "ahk_response.txt"
        self.progress_file = "ahk_progress.txt"
        
    def send_command(self, command, params=None):
        """Envía un comando a AHK y espera respuesta"""
        if params is None:
            params = []
            
        # Limpiar archivos de respuesta anteriores
        for file in [self.done_file, self.response_file, self.progress_file]:
            if os.path.exists(file):
                os.remove(file)
        
        # Escribir comando
        command_str = command
        if params:
            command_str += "|" + "|".join(str(p) for p in params)
            
        try:
            with open(self.command_file, 'w', encoding='utf-8') as f:
                f.write(command_str)
            
            # Esperar respuesta
            timeout = 30  # 30 segundos máximo de espera
            start_time = time.time()
            
            while not os.path.exists(self.done_file):
                if time.time() - start_time > timeout:
                    return False, "Timeout esperando respuesta de AHK"
                time.sleep(0.1)
            
            # Leer respuesta si existe
            response = "OK"
            if os.path.exists(self.response_file):
                with open(self.response_file, 'r', encoding='utf-8') as f:
                    response = f.read().strip()
                os.remove(self.response_file)
            
            # Limpiar archivo done
            if os.path.exists(self.done_file):
                os.remove(self.done_file)
                
            return True, response
            
        except Exception as e:
            return False, f"Error comunicándose con AHK: {str(e)}"
    
    def set_config(self, csv_file=None):
        """Configura los parámetros del script"""
        params = []
        if csv_file:
            params.append(csv_file)
        return self.send_command("SET_CONFIG", params)
    
    def read_csv(self):
        """Solicita leer el CSV"""
        return self.send_command("READ_CSV")
    
    def execute_ge_route(self):
        """Ejecuta el script GE Route completo"""
        return self.send_command("EXECUTE_GE_ROUTE")
    
    def stop_script(self):
        """Detiene el script en ejecución"""
        return self.send_command("STOP_SCRIPT")
    
    def get_status(self):
        """Obtiene el estado actual"""
        return self.send_command("GET_STATUS")
    
    def get_progress(self):
        """Obtiene el progreso actual"""
        if os.path.exists(self.progress_file):
            try:
                with open(self.progress_file, 'r', encoding='utf-8') as f:
                    progress_data = f.read().strip()
                os.remove(self.progress_file)
                return True, progress_data
            except:
                return False, "Error leyendo progreso"
        return False, "No hay datos de progreso"

def clear_screen():
    """Limpia la pantalla de la consola"""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_header():
    """Imprime el encabezado del programa"""
    print("=" * 60)
    print("       CONTROLADOR GE ROUTE - AUTOMATIZACIÓN")
    print("=" * 60)
    print()

def main():
    # Configuración inicial
    csv_file = r"NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
    
    # Inicializar controlador AHK
    ahk = AHKController()
    
    clear_screen()
    print_header()
    
    # Verificar conexión con AHK
    print("Verificando conexión con AutoHotkey...")
    success, response = ahk.get_status()
    
    if not success:
        print(f"❌ ERROR: No se pudo conectar con AutoHotkey")
        print(f"   Mensaje: {response}")
        print()
        print("Asegúrate de que el script GE_Route_Automation.ahk esté en ejecución")
        input("Presiona Enter para salir...")
        return
    
    print("✅ Conexión con AutoHotkey establecida")
    print()
    
    # Configurar parámetros
    print("Configurando parámetros...")
    print(f"  - Archivo CSV: {csv_file}")
    print()
    
    success, response = ahk.set_config(csv_file)
    if not success:
        print(f"❌ Error en configuración: {response}")
        input("Presiona Enter para salir...")
        return
    
    print("✅ Configuración aplicada correctamente")
    print()
    
    # Verificar CSV
    print("Verificando archivo CSV...")
    success, response = ahk.read_csv()
    if not success:
        print(f"❌ Error con el archivo CSV: {response}")
        input("Presiona Enter para salir...")
        return
    
    print("✅ Archivo CSV verificado correctamente")
    print()
    
    # Confirmación final antes de ejecutar
    print("⚠️  ADVERTENCIA: El script comenzará en 3 segundos")
    print("   Asegúrate de que la ventana de GE esté activa")
    print("   Presiona Ctrl+C para cancelar")
    print()
    
    try:
        input("Presiona Enter para INICIAR la automatización GE Route...")
        
        # Cuenta regresiva
        for i in range(3, 0, -1):
            print(f"▶️  Iniciando en {i}...")
            time.sleep(1)
        
        print()
        print("🚀 INICIANDO AUTOMATIZACIÓN GE ROUTE...")
        print("   Presiona Ctrl+C en cualquier momento para detener")
        print()
        
        # Ejecutar script GE Route
        success, response = ahk.execute_ge_route()
        if not success:
            print(f"❌ Error al iniciar ejecución: {response}")
            input("Presiona Enter para salir...")
            return
        
        print("✅ Ejecución GE Route iniciada")
        print()
        
        # Monitorear progreso
        print("Monitoreando progreso...")
        last_progress = 0
        
        try:
            while True:
                # Verificar estado
                success, status = ahk.get_status()
                if success:
                    if "RUNNING" in status:
                        # Obtener progreso
                        success, progress = ahk.get_progress()
                        if success and "PROGRESS" in progress:
                            parts = progress.split("|")
                            if len(parts) >= 2:
                                current_iteration = parts[1]
                                if current_iteration != last_progress:
                                    print(f"📊 Procesando iteración: {current_iteration}/9")
                                    last_progress = current_iteration
                        else:
                            print(".", end="", flush=True)
                    elif "COMPLETED" in status:
                        print()
                        print("🎉 AUTOMATIZACIÓN COMPLETADA EXITOSAMENTE!")
                        break
                    elif "IDLE" in status:
                        print()
                        print("⏸️  Script en espera")
                        break
                    elif "STOPPED" in status:
                        print()
                        print("⏹️  Script detenido")
                        break
                else:
                    print(f"❌ Error obteniendo estado: {status}")
                    break
                
                time.sleep(2)
                
        except KeyboardInterrupt:
            print()
            print("🛑 Deteniendo script...")
            ahk.stop_script()
            print("Script detenido por el usuario")
        
        print()
        input("Presiona Enter para salir...")
        
    except KeyboardInterrupt:
        print()
        print("❌ Ejecución cancelada por el usuario")
        input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()