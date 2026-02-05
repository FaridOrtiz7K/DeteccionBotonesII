import subprocess
import time
import os
import logging

logger = logging.getLogger(__name__)

class AHKManagerCD:
    def __init__(self):
        self.ahk_process = None
        self.script_path = "ahk_copiar_eliminar.ahk"
        
    def crear_script_ahk(self):
        """Crea automáticamente el script de AutoHotkey para copiar y eliminar valores"""
        ahk_script = """
#Persistent
#SingleInstance force

; Script de AutoHotkey para copiar y eliminar valores
Loop {
    ; Esperar comandos de Python
    FileRead, comando, ahk_command.txt
    if (ErrorLevel = 0) {
        FileDelete, ahk_command.txt
        
        ; Parsear comando: x,y
        Array := StrSplit(comando, ",")
        x_campo := Array[1]
        y_campo := Array[2]
        
        ; Ejecutar acciones: copiar y eliminar
        Click, %x_campo% %y_campo%
        Sleep, 200
        
        ; Seleccionar todo el texto
        Send, ^a
        Sleep, 100
        
        ; Copiar el valor seleccionado al portapapeles
        Send, ^c
        Sleep, 200
        
        ; Guardar el valor del portapapeles en un archivo
        ClipboardTemp := Clipboard
        FileAppend, %ClipboardTemp%, ahk_copied_value.txt
        Sleep, 100
        
        ; Confirmación para Python
        FileAppend, done, ahk_done.txt
    }
    Sleep, 500  ; Revisar cada medio segundo
}
"""
        try:
            with open(self.script_path, "w", encoding="utf-8") as f:
                f.write(ahk_script)
            logger.info("Script de AutoHotkey creado automáticamente")
            return True
        except Exception as e:
            logger.error(f"Error creando script AHK: {e}")
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
                logger.info("AutoHotkey iniciado correctamente")
            else:
                logger.error("AutoHotkey no se pudo iniciar")
            return is_running
        except Exception as e:
            logger.error(f"Error iniciando AutoHotkey: {e}")
            return False
            
    def stop_ahk(self):
        """Detiene AutoHotkey correctamente"""
        if self.ahk_process:
            try:
                self.ahk_process.terminate() 
                self.ahk_process.wait(timeout=5) 
                logger.info("AutoHotkey detenido correctamente")
            except subprocess.TimeoutExpired:
                self.ahk_process.kill() 
                logger.warning("AutoHotkey fue forzado a detenerse")
            except Exception as e:
                logger.error(f"Error deteniendo AutoHotkey: {e}")
                
    def ejecutar_acciones_ahk(self, x_campo, y_campo):
        """Envía comandos a AutoHotkey y retorna el valor copiado"""
        # Limpiar archivos temporales previos
        for temp_file in ["ahk_command.txt", "ahk_done.txt", "ahk_copied_value.txt"]:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass
        
        # Crear archivo de comandos para AHK
        comando = f"{x_campo},{y_campo}"
        
        try:
            with open("ahk_command.txt", "w", encoding="utf-8") as f:
                f.write(comando)
            
            logger.info(f"Comando enviado a AHK: {comando}")
            
            # Esperar a que AHK complete la acción
            timeout = 10  # 10 segundos de timeout
            start_time = time.time()
            
            while not os.path.exists("ahk_done.txt"):
                if time.time() - start_time > timeout:
                    logger.error("Timeout esperando respuesta de AHK")
                    return None
                time.sleep(0.1)
            
            # Leer el valor copiado
            valor_copiado = None
            if os.path.exists("ahk_copied_value.txt"):
                with open("ahk_copied_value.txt", "r", encoding="utf-8") as f:
                    valor_copiado = f.read().strip()
                
                # Limpiar archivo temporal
                os.remove("ahk_copied_value.txt")
            
            # Limpiar archivo de confirmación
            if os.path.exists("ahk_done.txt"):
                os.remove("ahk_done.txt")
            
            logger.info(f"Valor copiado: {valor_copiado}")
            return valor_copiado
            
        except Exception as e:
            logger.error(f"Error en ejecutar_acciones_ahk: {e}")
            # Limpiar archivos temporales en caso de error
            for temp_file in ["ahk_command.txt", "ahk_done.txt", "ahk_copied_value.txt"]:
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass
            return None