import pyautogui
import keyboard
import time
import csv
import os
from tkinter import messagebox

class GERouteAutomation:
    def __init__(self):
        self.csv_file = r"C:\Users\cmf05\Documents\AutoHotkey\Nombres negocios SIGP.csv"
        
    def check_files(self):
        """Verifica que el archivo CSV exista"""
        if not os.path.exists(self.csv_file):
            messagebox.showerror("Error", f"El archivo CSV no existe en: {self.csv_file}")
            return False
        return True
    
    def read_csv_data(self):
        """Lee los datos del CSV"""
        try:
            with open(self.csv_file, 'r', encoding='utf-8') as file:
                return list(csv.reader(file))
        except Exception as e:
            messagebox.showerror("Error", f"Error leyendo CSV: {str(e)}")
            return None
    
    def safe_click(self, x, y, wait_time=1):
        """Clic seguro con verificación de ESC"""
        if keyboard.is_pressed('esc'):
            return False
        pyautogui.click(x, y)
        time.sleep(wait_time)
        return True
    
    def safe_type(self, text, wait_time=1):
        """Escribir texto de forma segura"""
        if keyboard.is_pressed('esc'):
            return False
        pyautogui.write(text)
        time.sleep(wait_time)
        return True
    
    def safe_press(self, key, wait_time=1):
        """Presionar tecla de forma segura"""
        if keyboard.is_pressed('esc'):
            return False
        pyautogui.press(key)
        time.sleep(wait_time)
        return True
    
    def execute_automation(self):
        """Ejecuta la automatización principal"""
        if not self.check_files():
            return False
            
        lines = self.read_csv_data()
        if not lines or len(lines) < 1:
            messagebox.showerror("Error", "CSV vacío o con formato incorrecto")
            return False
        
        messagebox.showinfo("Preparado", "Script listo. Posicione el cursor y espere 5 segundos.")
        time.sleep(5)
        
        num_ra = 0  # Índice para las líneas del CSV
        
        for iteration in range(9):
            if keyboard.is_pressed('esc'):
                messagebox.showinfo("Información", "Script detenido por el usuario")
                return True
                
            print(f"Procesando iteración {iteration + 1}/9")
            
            # Secuencia de clics para agregar ruta
            steps = [
                (327, 381, 2),   # Agregar ruta de GE
                (1396, 608, 3),  # Archivo
                (1406, 634, 3),  # Abrir
                (1120, 666, 3),  # Documents
            ]
            
            for x, y, wait in steps:
                if not self.safe_click(x, y, wait):
                    return True
            
            # Alt + n
            if keyboard.is_pressed('esc'): return True
            pyautogui.hotkey('alt', 'n')
            time.sleep(3)
            
            # Escribir nombre del archivo KML
            if num_ra < len(lines):
                line_data = lines[num_ra]
                if len(line_data) >= 1:
                    kml_name = f"RA {line_data[0]}.kml"
                    if not self.safe_type(kml_name, 1):
                        return True
            
            time.sleep(3)
            if not self.safe_press('enter', 3):
                return True
            
            # Segunda secuencia de clics
            steps2 = [
                (327, 381, 2),   # Agregar ruta de GE
                (1406, 675, 2),  # Cargar ruta
                (70, 266, 2),    # Select Lote
                (168, 188, 2),   # Seleccionar en el mapa
                (1366, 384, 2),  # Anotar
                (1449, 452, 2),  # Agregar texto adicional
                (1421, 526, 2),  # Escriba el texto adicional
            ]
            
            for x, y, wait in steps2:
                if not self.safe_click(x, y, wait):
                    return True
            
            # Limpiar campo y escribir texto
            if not self.safe_press('backspace', 1):
                return True
            
            if num_ra < len(lines):
                line_data = lines[num_ra]
                if len(line_data) >= 2:
                    if not self.safe_type(line_data[1], 1):
                        return True
                    
                    # Finalizar texto
                    steps3 = [
                        (1263, 572, 3),  # Agregar texto
                        (1338, 570, 2),  # Cerrar
                    ]
                    
                    for x, y, wait in steps3:
                        if not self.safe_click(x, y, wait):
                            return True
            
            # Guardar cada 10 iteraciones (aunque solo son 9)
            if (iteration + 1) % 10 == 0:
                if keyboard.is_pressed('esc'): return True
                pyautogui.hotkey('ctrl', 's')
                time.sleep(6)
            
            # Limpiar y pasar al siguiente
            steps4 = [
                (360, 980, 1),  # Limpiar trazo
                (70, 266, 2),   # Select Lote
            ]
            
            for x, y, wait in steps4:
                if not self.safe_click(x, y, wait):
                    return True
            
            if not self.safe_press('down', 2):
                return True
            
            num_ra += 1
        
        # Guardar final
        if keyboard.is_pressed('esc'): return True
        pyautogui.hotkey('ctrl', 's')
        time.sleep(6)
        
        messagebox.showinfo("Completado", "El script ha finalizado exitosamente!")
        return True

def main():
    """Función principal"""
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1
    
    print("=== Automatización de Rutas GE ===")
    print("Instrucciones:")
    print("1. Asegúrese de tener el archivo CSV en la ubicación correcta")
    print("2. Abra la aplicación donde ejecutará el script")
    print("3. El script comenzará automáticamente después de 5 segundos")
    print("4. Presione ESC en cualquier momento para detener")
    
    automation = GERouteAutomation()
    automation.execute_automation()

if __name__ == "__main__":
    main()