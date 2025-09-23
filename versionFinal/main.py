import sys
import os

# Asegurar que Python encuentre los módulos en el mismo directorio
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

try:
    from controller import ScriptController
    import pyautogui
    
    def main():
        # Configurar pyautogui
        pyautogui.FAILSAFE = True
        
        # Crear y ejecutar la aplicación
        app = ScriptController()
        app.run()
        
        # Iniciar el loop principal de Tkinter
        import tkinter as tk
        tk.mainloop()

    if __name__ == "__main__":
        main()

except ImportError as e:
    print(f"Error de importación: {e}")
    print("Asegúrate de que todos los archivos estén en el mismo directorio:")
    print("- model.py")
    print("- login_view.py")
    print("- main_view.py")
    print("- pause_view.py")
    print("- controller.py")
    input("Presiona Enter para salir...")