import subprocess
import time
import os
import logging

logger = logging.getLogger(__name__)

class AHKClickDown:
    def __init__(self):
        self.ahk_process = None
        self.script_path = "ahk_click_down.ahk"
        
    def crear_script_ahk(self):
        """Crea automáticamente el script de AutoHotkey para clics y flechas down"""
        ahk_script = """
#Persistent
#SingleInstance force

; Script de AutoHotkey para clics y flechas down
Loop {
    ; Esperar comandos de Python
    FileRead, comando, ahk_click_down_command.txt
    if (ErrorLevel = 0) {
        FileDelete, ahk_click_down_command.txt
        
        ; Parsear comando: x,y,veces_down
        Array := StrSplit(comando, "|")
        x_campo := Array[1]
        y_campo := Array[2]
        veces_down := Array[3]
        
        ; Hacer click en las coordenadas especificadas
        Click, %x_campo% %y_campo%
        Sleep, 300
        
        ; Presionar flecha down las veces especificadas
        Loop, %veces_down% {
            Send, {Down}
            Sleep, 150
        }
        
        ; Confirmación para Python
        FileAppend, done, ahk_click_down_done.txt
    }
    Sleep, 500  ; Revisar cada medio segundo
}
"""
        try:
            with open(self.script_path, "w", encoding="utf-8") as f:
                f.write(ahk_script)
            logger.info("Script de AutoHotkey (click down) creado automáticamente")
            return True
        except Exception as e:
            logger.error(f"Error creando script AHK click down: {e}")
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
                logger.info("AutoHotkey (click down) iniciado correctamente")
            else:
                logger.error("AutoHotkey (click down) no se pudo iniciar")
            return is_running
        except Exception as e:
            logger.error(f"Error iniciando AutoHotkey (click down): {e}")
            return False
            
    def stop_ahk(self):
        """Detiene AutoHotkey correctamente"""
        if self.ahk_process:
            try:
                self.ahk_process.terminate() 
                self.ahk_process.wait(timeout=5) 
                logger.info("AutoHotkey (click down) detenido correctamente")
            except subprocess.TimeoutExpired:
                self.ahk_process.kill() 
                logger.warning("AutoHotkey (click down) fue forzado a detenerse")
            except Exception as e:
                logger.error(f"Error deteniendo AutoHotkey (click down): {e}")
                
    def ejecutar_click_down(self, x_campo, y_campo, veces_down):
        """Envía comandos a AutoHotkey para hacer clic y presionar flecha down"""
        # Limpiar archivos temporales previos
        for temp_file in ["ahk_click_down_command.txt", "ahk_click_down_done.txt"]:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass
        
        # Crear archivo de comandos para AHK
        comando = f"{x_campo}|{y_campo}|{veces_down}"
        
        try:
            with open("ahk_click_down_command.txt", "w", encoding="utf-8") as f:
                f.write(comando)
            
            logger.info(f"Comando click+down enviado a AHK: {comando}")
            
            # Esperar a que AHK complete la acción
            timeout = 10  # 10 segundos de timeout
            start_time = time.time()
            
            while not os.path.exists("ahk_click_down_done.txt"):
                if time.time() - start_time > timeout:
                    logger.error("Timeout esperando respuesta de AHK (click down)")
                    return False
                time.sleep(0.1)
            
            # Limpiar archivo de confirmación
            if os.path.exists("ahk_click_down_done.txt"):
                os.remove("ahk_click_down_done.txt")
            
            logger.info("Click + Down ejecutado correctamente")
            return True
            
        except Exception as e:
            logger.error(f"Error en ejecutar_click_down: {e}")
            # Limpiar archivos temporales en caso de error
            for temp_file in ["ahk_click_down_command.txt", "ahk_click_down_done.txt"]:
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                    except:
                        pass
            return False