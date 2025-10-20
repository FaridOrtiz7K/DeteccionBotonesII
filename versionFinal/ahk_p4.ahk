import subprocess
import os
import time
from tkinter import messagebox

class AHKExecutor:
    def __init__(self):
        self.ahk_script = r"C:\Users\cmf05\Documents\AutoHotkey\GE_Route_Script.ahk"
        self.ahk_exe = r"C:\Program Files\AutoHotkey\AutoHotkey.exe"  # Ruta típica de AHK
        
    def create_ahk_script(self):
        """Crea el script de AutoHotkey si no existe"""
        ahk_content = '''#Persistent
SetTitleMatchMode, 2

; Define the hotkey to exit the script
Esc::ExitApp

; Define a hotkey to start the actions
F2::PerformActions()

; Function to perform the clicks and actions
PerformActions() {
    ; Check if the CSV file exists
    if (!FileExist("C:\\Users\\cmf05\\Documents\\AutoHotkey\\Nombres negocios SIGP.csv")) {
        MsgBox, The specified CSV file does not exist!
        return
    }

    ; Read the CSV file
    FileRead, csvContent, C:\\Users\\cmf05\\Documents\\AutoHotkey\\Nombres negocios SIGP.csv
    lines := StrSplit(csvContent, "`n")  ; Split into lines (correct delimiter for newlines)
    
    ; Check if we have enough lines
    if (lines.MaxIndex() < 1)  ; Ensure at least 1 line exists
    {
        MsgBox, Not enough data in the CSV file!
        return
    }
	

    numRa := 1 ; Initialize the variable
	numTxtType := 0
    
    Loop, 9 { ; Loop 9 times
        Click, 327, 381 ;Select Agregar ruta de GE
        Sleep, 2000 ; Wait 1 second
        Click, 1396, 608 ;Archivo
        Sleep, 3000
        Click, 1406, 634 ;Abrir 
        Sleep, 3000 ; Wait 3 seconds
        Click, 1120, 666 ;Documents
        Sleep, 3000 ; Wait 3 seconds
        Send, !n ; Press Alt + n
        Sleep, 3000
		
        ; Check if we are at a valid iteration
        if (numRa <= lines.MaxIndex())  ; Ensure we don't go beyond available lines
        {
            ; Split the current line into columns
            values1 := StrSplit(lines[numRa], ",")  ; Access the appropriate line based on numRa

            ; Ensure the second column exists
            if (values1.MaxIndex() >= 2)  
            {

                ; Write the value from the second column (B1 for the first iteration)
				numTxtType := values1[1]
               Send, RA %numTxtType%.kml ; Use the variable
			   
                Sleep, 1000  ; Wait 1 second after sending the value to ensure the program processes it

            }
        }		
		
        
        Sleep, 3000
        Send, {Enter}    
        Sleep, 3000 ; Wait 4 seconds
        Click, 327, 381 ;Select Agregar ruta de GE
        Sleep, 2000 ; Wait 2 seconds
		Click, 1406, 675 ; Cargar ruta
        Sleep, 2000
		Click, 70, 266 ; Select Lote
		Sleep, 2000
		Click, 168, 188 ; Seleccionar en el mapa
		Sleep, 2000
		Click, 1366, 384 ; Anotar
		Sleep, 2000
		Click, 1449, 452 ; Agregar texto adicional
		Sleep, 2000
		Click, 1421, 526 ; Escriba el texto adicional
		Sleep, 2000
        ; Press delete key
        Send, {Backspace}
        Sleep, 1000  ; Small wait before writing the value

       
        ; Check if we are at a valid iteration
        if (numRa <= lines.MaxIndex())  ; Ensure we don't go beyond available lines
        {
            ; Split the current line into columns
            values := StrSplit(lines[numRa], ",")  ; Access the appropriate line based on numRa

            ; Ensure the second column exists
            if (values.MaxIndex() >= 2)  
            {

                ; Write the value from the second column (B1 for the first iteration)
                Send, % values[2]  ; Send the value from the second column
                Sleep, 1000  ; Wait 1 second after sending the value to ensure the program processes it

                ; Left click at (1260, 197) Agregar texto
                Click, 1263, 572
                Sleep, 3000  ; Wait 3 seconds
				Click, 1338, 570 ; cerrar
				Sleep, 2000
            }
        }
        
        ; Every 10 steps, press Ctrl + S
        if (Mod(A_Index, 10) = 0) 
		{ ; Using Mod for clarity
            Send, ^s ; Press Ctrl + S
            Sleep, 6000 ; Wait 6 seconds
        }
		Click, 360, 980 ; Limpiar trazo
		Sleep, 1000		
		Click, 70, 266 ; Select Lote
		Sleep, 2000
		Send, {Down}
		Sleep, 2000	
		

        numRa++ ; Increment the variable
    }

    Send, ^s ; Press Ctrl + S at the end
    MsgBox, The script has finished! ; Notify the user
    return
}
'''
        
        try:
            with open(self.ahk_script, 'w', encoding='utf-8') as f:
                f.write(ahk_content)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear el script AHK: {str(e)}")
            return False
    
    def check_ahk_installation(self):
        """Verifica si AutoHotkey está instalado"""
        if os.path.exists(self.ahk_exe):
            return True
        
        # Buscar en otras ubicaciones comunes
        common_paths = [
            r"C:\Program Files\AutoHotkey\AutoHotkey.exe",
            r"C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe",
            os.path.expanduser(r"~\Desktop\AutoHotkey\AutoHotkey.exe"),
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                self.ahk_exe = path
                return True
        
        return False
    
    def run_ahk_script(self):
        """Ejecuta el script de AutoHotkey"""
        if not self.check_ahk_installation():
            messagebox.showerror("Error", 
                               "AutoHotkey no está instalado o no se encuentra.\n"
                               "Por favor, instale AutoHotkey para usar esta opción.")
            return False
        
        if not os.path.exists(self.ahk_script):
            if not self.create_ahk_script():
                return False
        
        try:
            messagebox.showinfo("Iniciando", 
                              "Se ejecutará el script de AutoHotkey.\n"
                              "Presione F2 en la ventana de AutoHotkey para iniciar.\n"
                              "Presione ESC para detener.")
            
            # Ejecutar AutoHotkey
            process = subprocess.Popen([self.ahk_script], shell=True)
            
            print("Script de AutoHotkey ejecutado. Presione F2 para iniciar la automatización.")
            print("El proceso se está ejecutando en segundo plano.")
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo ejecutar AutoHotkey: {str(e)}")
            return False

def main_ahk_executor():
    """Función principal para ejecutar AutoHotkey"""
    executor = AHKExecutor()
    
    print("=== Ejecutor de Script AutoHotkey ===")
    print("Este programa ejecutará el script original de AutoHotkey.")
    
    if executor.run_ahk_script():
        print("Script AutoHotkey iniciado correctamente.")
    else:
        print("Error al iniciar el script AutoHotkey.")

# Menú principal para elegir entre ambas opciones
def main_menu():
    """Menú principal para elegir entre Python puro o AutoHotkey"""
    print("=== AUTOMATIZACIÓN DE RUTAS GE ===")
    print("1. Ejecutar versión Python (Recomendado)")
    print("2. Ejecutar script AutoHotkey desde Python")
    print("3. Salir")
    
    while True:
        choice = input("\nSeleccione una opción (1-3): ").strip()
        
        if choice == '1':
            # Ejecutar versión Python pura
            from GERouteAutomation import main
            main()
            break
        elif choice == '2':
            # Ejecutar AutoHotkey desde Python
            main_ahk_executor()
            break
        elif choice == '3':
            print("Saliendo...")
            break
        else:
            print("Opción no válida. Por favor, seleccione 1, 2 o 3.")

if __name__ == "__main__":
    main_menu()