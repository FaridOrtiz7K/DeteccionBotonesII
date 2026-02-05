import os
import sys
import logging

# Agregar directorio actual al path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from models.estado import EstadoEjecucion
from models.datos_globales import DatosGlobales
from views.interfaz_principal import InterfazPrincipal
from controllers.controlador_principal import ControladorPrincipal

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('nse_automation.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def main():
    root = tk.Tk()
    
    # Crear instancias del modelo
    estado_global = EstadoEjecucion()
    datos_globales = DatosGlobales()
    
    # Crear vista
    vista = InterfazPrincipal(root)
    
    # Crear controlador
    controlador = ControladorPrincipal(vista, estado_global, datos_globales)
    
    root.mainloop()

if __name__ == "__main__":
    main()