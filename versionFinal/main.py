# main.py
import tkinter as tk
from controllers.controller import ControladorAutomation
from views import VistaAutomation

def main():
    """Función principal para ejecutar la aplicación MVC"""
    root = tk.Tk()
    
    # Crear controlador y vista
    controlador = ControladorAutomation()
    vista = VistaAutomation(root, controlador)
    controlador.set_vista(vista)
    
    # Iniciar aplicación
    root.mainloop()

if __name__ == "__main__":
    main()