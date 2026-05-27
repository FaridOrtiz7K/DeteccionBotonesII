import os
import sys
import logging
import tkinter as tk
from views.main_window import InterfazAutomation
from controllers.automation_controller import AutomationController

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('nse_automation.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

def main():
    root = tk.Tk()

    # Establecer icono
    icon_path = os.path.join(os.path.dirname(__file__), 'img', 'icon.ico')
    if os.path.exists(icon_path):
        root.iconbitmap(icon_path)
    else:
        logging.warning(f"Icono no encontrado en {icon_path}")

    # Crear controlador y vista
    controller = AutomationController(view=None)  # Temporal
    view = InterfazAutomation(root, controller)
    controller.view = view  # Inyectar vista

    root.mainloop()

if __name__ == "__main__":
    main()