import subprocess
import time
import os
import logging

logger = logging.getLogger(__name__)

class AHKWriter:
    def __init__(self):
        self.ahk_process = None
        self.script_path = "ahk_writer.ahk"
        
    def crear_script_ahk(self):
        """Crea automáticamente el script de AutoHotkey para escribir texto"""
        ahk_script = """
#Persistent
#SingleInstance force

; Script de AutoHotkey para escribir texto en coordenadas específicas
Loop {
    ; Esperar comandos de Python
    FileRead, comando, ahk_writer_command.txt
    if (ErrorLevel = 0) {
        FileDelete, ahk_writer_command.txt
        
        ; Parsear comando: x,y,texto
        Array := StrSplit(comando, "|")
        x_campo := Array[1]
        y_campo := Array[2]
        texto := Array[3]
        
        ; Hacer click en las coordenadas especificadas
        Click, %x_campo% %y_campo%
        Sleep, 300
        
        ; Seleccionar y limpiar el campo (opcional)
        Send, ^a
        Sleep, 100
        Send, {Delete}
        Sleep, 100
        
        ; Escribir el texto
        SendInput, %texto%
        Sleep, 300
        
        ; Confirmación para Python
        FileAppend, done, ahk_writer_done.txt
    }
    Sleep, 500  ; Revisar cada medio segundo
}
"""
        try:
            with open(self.script_path, "w", encoding="utf-8") as f:
                f.write(ahk_script)
            logger.info("Script de AutoHotkey (writer) creado automáticamente")
            return True
        except Exception as e:
            logger.error(f"Error creando script AHK writer: {e}")
            return False
            
    def start_ahk(self):
        """Inicia AutoHotkey de manera más robusta"""
        if self.ahk_process and self.ahk_process.poll() is None:
            return True  # Ya está en ejecución
            
        try:
            if not os.path.exists(self.script_path):
                if not self.crear_script_ahk():
                    return False
                    
            self.ahk_process = subprocess.Popen(['AutoHotkey_1.1.37.02/AutoHotkeyU64.exe', self.script_path])
            time.sleep(2)
            is_running = self.ahk_process.poll() is None
            if is_running:
                logger.info("AutoHotkey (writer) iniciado correctamente")
            else:
                logger.error("AutoHotkey (writer) no se pudo iniciar")
            return is_running
        except Exception as e:
            logger.error(f"Error iniciando AutoHotkey (writer): {e}")
            return False
            
    def stop_ahk(self):
        """Detiene AutoHotkey correctamente"""
        if self.ahk_process:
            try:
                self.ahk_process.terminate() 
                self.ahk_process.wait(timeout=5) 
                logger.info("AutoHotkey (writer) detenido correctamente")
            except subprocess.TimeoutExpired:
                self.ahk_process.kill() 
                logger.warning("AutoHotkey (writer) fue forzado a detenerse")
            except Exception as e:
                logger.error(f"Error deteniendo AutoHotkey (writer): {e}")
                
    def ejecutar_escritura_ahk(self, x_campo, y_campo, texto):
        """Envía comandos a AutoHotkey para escribir texto en coordenadas específicas"""
        # Limpiar archivos temporales previos
        for temp_file in ["ahk_writer_command.txt", "ahk_writer_done.txt"]:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass
        
        # Crear archivo de comandos para AHK
        # Usamos "|" como separador para evitar conflictos con comas en el texto
        comando = f"{x_campo}|{y_campo}|{texto}"
        
        try:
            with open("ahk_writer_command.txt", "w", encoding="utf-8") as f:
                f.write(comando)
            
            logger.info(f"Comando de escritura enviado a AHK: {comando}")
            
            # Esperar a que AHK complete la acción
            timeout = 10  # 10 segundos de timeout
            start_time = time.time()
            
            while not os.path.exists("ahk_writer_done.txt"):
                if time.time() - start_time > timeout:
                    logger.error("Timeout esperando respuesta de AHK (writer)")
                    return False
                time.sleep(0.1)
            
            # Limpiar archivo de confirmación
            if os.path.exists("ahk_writer_done.txt"):
                os.remove("ahk_writer_done.txt")
            
            logger.info("Escritura completada correctamente")
            return True
            
        except Exception as e:
            logger.error(f"Error en ejecutar_escritura_ahk: {e}")
            # Limpiar archivos temporales en caso de error
            for temp_file in ["ahk_writer_command.txt", "ahk_writer_done.txt"]:
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass
            return False