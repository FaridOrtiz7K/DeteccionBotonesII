import tkinter as tk
from tkinter import messagebox, simpledialog
import time
import csv
import pyautogui
import keyboard
import sys

# Configuración de PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

class NSEScript:
    def __init__(self):
        self.start_count = 1
        self.loop_count = 589
        self.csv_file = r"NCO0004FO_ID Num Uso NSE Serv Nom Neg.csv"
        
    def create_gui(self):
        """Crea la interfaz gráfica para configuración"""
        root = tk.Tk()
        root.title("Configuración del Script")
        root.geometry("300x150")
        
        # Inicia en
        tk.Label(root, text="Inicia en:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        start_entry = tk.Entry(root, width=10)
        start_entry.grid(row=0, column=1, padx=5, pady=5)
        start_entry.insert(0, "1")
        
        # Número de lotes
        tk.Label(root, text="Número de lotes:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        loop_entry = tk.Entry(root, width=10)
        loop_entry.grid(row=1, column=1, padx=5, pady=5)
        loop_entry.insert(0, "589")
        
        def start_script():
            try:
                self.start_count = int(start_entry.get())
                self.loop_count = int(loop_entry.get())
                
                confirm = messagebox.askyesno(
                    "Confirmar", 
                    f"¿Iniciar script desde {self.start_count} con {self.loop_count} lotes?"
                )
                
                if confirm:
                    root.destroy()
                    self.execute_script()
                else:
                    return
                    
            except ValueError:
                messagebox.showerror("Error", "Por favor ingrese valores numéricos válidos")
        
        tk.Button(root, text="Iniciar", command=start_script, 
                 bg="lightblue", width=15).grid(row=2, column=0, columnspan=2, pady=10)
        
        root.mainloop()
    
    def read_csv(self):
        """Lee el archivo CSV y retorna las líneas"""
        try:
            with open(self.csv_file, 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                lines = list(reader)
            return lines
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo CSV: {str(e)}")
            return None
    
    def click_and_wait(self, x, y, wait_time=1.5):
        """Realiza un clic y espera"""
        if keyboard.is_pressed('esc'):
            raise KeyboardInterrupt("Script detenido por el usuario")
        
        pyautogui.click(x, y)
        time.sleep(wait_time)
    
    def send_keys_and_wait(self, text, wait_time=1.5):
        """Envía texto y espera"""
        if keyboard.is_pressed('esc'):
            raise KeyboardInterrupt("Script detenido por el usuario")
        
        pyautogui.write(text)
        time.sleep(wait_time)
    
    def handle_u_logic(self, columns):
        """Maneja la lógica para tipo 'U'"""
        # Coordenadas para tipo U
        coords_select_u = {
            6: (1268, 637), 7: (1268, 661), 8: (1268, 685), 9: (1268, 709),
            10: (1268, 733), 11: (1268, 757), 12: (1268, 781), 13: (1268, 825),
            14: (1268, 856), 15: (1268, 881), 16: (1268, 908)
        }
        
        for col_index in range(6, 17):
            if len(columns) > col_index and columns[col_index].strip() and float(columns[col_index]) > 0:
                x, y = coords_select_u[col_index]
                self.click_and_wait(x, y, 3)
        
        self.click_and_wait(1306, 639, 2)  # Asignar NSE (Confirm)
    
    def handle_v_logic(self, columns):
        """Maneja la lógica para tipo 'V'"""
        # Coordenadas para tipo V
        coords_select = {
            6: (1235, 563), 7: (1235, 602), 8: (1235, 630), 9: (1235, 668),
            10: (1235, 702), 11: (1600, 563), 12: (1600, 602), 13: (1600, 630),
            14: (1235, 772), 15: (1235, 804), 16: (1235, 838)
        }
        
        coords_type = {
            6: (1365, 563), 7: (1365, 602), 8: (1365, 630), 9: (1365, 668),
            10: (1365, 702), 11: (1730, 563), 12: (1730, 602), 13: (1730, 630),
            14: (1365, 772), 15: (1365, 804), 16: (1365, 838)
        }
        
        for col_index in range(6, 17):
            if len(columns) > col_index and columns[col_index].strip() and float(columns[col_index]) > 0:
                # Click en coordenada select
                x_cs, y_cs = coords_select[col_index]
                self.click_and_wait(x_cs, y_cs, 2)
                
                # Click en coordenada type
                x_ct, y_ct = coords_type[col_index]
                self.click_and_wait(x_ct, y_ct, 2)
                
                # Escribir valor
                self.send_keys_and_wait(columns[col_index], 2)
        
        self.click_and_wait(1648, 752, 2)  # Asignar NSE (Confirm)
        self.click_and_wait(1598, 823, 2)  # Cerrar Asignar NSE
    
    def handle_services(self, columns):
        """Maneja la lógica de administración de servicios"""
        if len(columns) > 17 and columns[17].strip() and float(columns[17]) > 0:
            self.click_and_wait(1563, 385, 2)  # Botón Administrar Servicios
            self.click_and_wait(100, 114, 2)   # Selecciona detalle con NSE
            
            # VOZ COBRE TELMEX LINEAS DE COBRE
            if len(columns) > 18 and columns[18].strip() and float(columns[18]) > 0:
                self.handle_voz_cobre(columns[18])
            
            # Datos s/dom
            if len(columns) > 19 and columns[19].strip() and float(columns[19]) > 0:
                self.handle_datos_sdom(columns[19])
            
            # Datos-cobre-telmex-inf
            if len(columns) > 20 and columns[20].strip() and float(columns[20]) > 0:
                self.handle_datos_cobre_telmex(columns[20])
            
            # Datos-fibra-telmex-inf
            if len(columns) > 21 and columns[21].strip() and float(columns[21]) > 0:
                self.handle_datos_fibra_telmex(columns[21])
            
            # TV cable otros
            if len(columns) > 22 and columns[22].strip() and float(columns[22]) > 0:
                self.handle_tv_cable_otros(columns[22])
            
            # Dish
            if len(columns) > 23 and columns[23].strip() and float(columns[23]) > 0:
                self.handle_dish(columns[23])
            
            # TVS
            if len(columns) > 24 and columns[24].strip() and float(columns[24]) > 0:
                self.handle_tvs(columns[24])
            
            # SKY
            if len(columns) > 25 and columns[25].strip() and float(columns[25]) > 0:
                self.handle_sky(columns[25])
            
            # VETV
            if len(columns) > 26 and columns[26].strip() and float(columns[26]) > 0:
                self.handle_vetv(columns[26])
            
            self.click_and_wait(882, 49, 5)  # Cerrar Botón Administrar Servicios
    
    def handle_voz_cobre(self, cantidad):
        """Maneja VOZ COBRE TELMEX LINEAS DE COBRE"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def handle_datos_sdom(self, cantidad):
        """Maneja Datos s/dom"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(138, 269, 2)  # Servicio
        self.navigate_down(2)  # Navegar a Datos
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def handle_datos_cobre_telmex(self, cantidad):
        """Maneja Datos-cobre-telmex-inf"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(138, 269, 2)  # Servicio
        self.navigate_down(2)  # Navegar a Datos
        self.click_and_wait(159, 355, 2)  # Producto
        self.navigate_down(1)  # Navegar hacia abajo
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def handle_datos_fibra_telmex(self, cantidad):
        """Maneja Datos-fibra-telmex-inf"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(138, 269, 2)  # Servicio
        self.navigate_down(2)  # Navegar a Datos
        self.click_and_wait(152, 294, 2)  # Tipo
        self.navigate_down(1)  # Navegar hacia abajo
        self.click_and_wait(150, 323, 2)  # Empresa
        self.navigate_down(1)  # Navegar hacia abajo
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def handle_tv_cable_otros(self, cantidad):
        """Maneja TV cable otros"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(138, 269, 2)  # Servicio
        self.navigate_down(3)  # Navegar a TV
        self.click_and_wait(150, 323, 2)  # Empresa
        self.navigate_down(4)  # Navegar a otros
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def handle_dish(self, cantidad):
        """Maneja Dish"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(138, 269, 2)  # Servicio
        self.navigate_down(3)  # Navegar a TV
        self.click_and_wait(152, 294, 2)  # Tipo
        self.navigate_down(2)  # Navegar a Satelital
        self.click_and_wait(150, 323, 2)  # Empresa
        self.navigate_down(1)  # Navegar a Dish
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def handle_tvs(self, cantidad):
        """Maneja TVS"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(138, 269, 2)  # Servicio
        self.navigate_down(3)  # Navegar a TV
        self.click_and_wait(152, 294, 2)  # Tipo
        self.navigate_down(2)  # Navegar a Satelital
        self.click_and_wait(150, 323, 2)  # Empresa
        self.navigate_down(2)  # Navegar a otro
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def handle_sky(self, cantidad):
        """Maneja SKY"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(138, 269, 2)  # Servicio
        self.navigate_down(3)  # Navegar a TV
        self.click_and_wait(152, 294, 2)  # Tipo
        self.navigate_down(2)  # Navegar a Satelital
        self.click_and_wait(150, 323, 2)  # Empresa
        self.navigate_down(3)  # Navegar a SKY
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def handle_vetv(self, cantidad):
        """Maneja VETV"""
        self.click_and_wait(100, 114, 2)  # Selecciona detalle con NSE
        self.click_and_wait(138, 269, 2)  # Servicio
        self.navigate_down(3)  # Navegar a TV
        self.click_and_wait(152, 294, 2)  # Tipo
        self.navigate_down(2)  # Navegar a Satelital
        self.click_and_wait(150, 323, 2)  # Empresa
        self.navigate_down(5)  # Navegar a VETV
        self.click_and_wait(127, 383, 2)  # Cantidad
        self.send_keys_and_wait(cantidad, 2)  # Cantidad de servicios
        self.click_and_wait(82, 423, 2)   # Guardar
        self.handle_error_click()
    
    def navigate_down(self, count):
        """Navega hacia abajo un número específico de veces"""
        for _ in range(count):
            pyautogui.press('down')
            time.sleep(2)
    
    def handle_error_click(self):
        """Maneja los clics de error"""
        for _ in range(5):
            self.click_and_wait(704, 384, 2)
    
    def execute_script(self):
        """Ejecuta el script principal"""
        try:
            # Leer archivo CSV
            lines = self.read_csv()
            if not lines or len(lines) < 2:
                messagebox.showerror("Error", "CSV vacío o con formato incorrecto")
                return
            
            time.sleep(5)  # Espera inicial de 5 segundos
            
            cont = self.start_count
            
            for i in range(self.loop_count):
                if keyboard.is_pressed('esc'):
                    messagebox.showinfo("Información", "Script detenido por el usuario")
                    return
                
                # Verificar que hay suficientes líneas
                if cont + 1 >= len(lines):
                    messagebox.showinfo("Información", "No hay más líneas en el CSV")
                    break
                
                current_line = lines[cont + 1]  # +1 para saltar la cabecera
                
                if len(current_line) < 5:
                    print(f"Línea {cont + 1} no tiene suficientes columnas")
                    cont += 1
                    continue
                
                # Proceso principal
                self.click_and_wait(89, 263)    # Select in the list
                self.click_and_wait(1483, 519)  # Case numero
                
                pyautogui.press('delete')
                time.sleep(1)
                
                self.send_keys_and_wait(current_line[1])  # Columna B
                
                # Manejar columna D
                if len(current_line) > 4 and current_line[3].strip() and float(current_line[3]) > 0:
                    self.click_and_wait(1507, 650)  # Punto USO
                    
                    # Navegar hacia abajo según columna D
                    for _ in range(int(float(current_line[3]))):
                        pyautogui.press('down')
                        time.sleep(2)
                
                self.click_and_wait(1290, 349)  # Actualizar
                
                # Manejar tipo U o V
                if len(current_line) > 5:
                    if current_line[4] == "U":
                        self.click_and_wait(169, 189)    # Seleccionar en mapa
                        self.click_and_wait(1463, 382)   # Asignar un solo NSE
                        self.click_and_wait(1266, 590)   # Casilla un solo NSE
                        self.handle_u_logic(current_line)
                    
                    elif current_line[4] == "V":
                        self.click_and_wait(169, 189)    # Seleccionar en mapa
                        self.click_and_wait(1491, 386, 3)  # Asignar varios NSE
                        self.handle_v_logic(current_line)
                
                # Manejar servicios
                self.handle_services(current_line)
                
                # Finalizar iteración
                self.click_and_wait(89, 263)  # Click again at (89, 263)
                pyautogui.press('down')
                time.sleep(3)
                
                cont += 1
                print(f"Procesado lote {cont} de {self.start_count + self.loop_count - 1}")
            
            # Proceso posterior al bucle
            self.click_and_wait(39, 55)  # Logo GE
            
            messagebox.showinfo("Información", 
                               "Script completado. Presionará F5 cada 3 minutos (presione Esc para salir)")
            
            while not keyboard.is_pressed('esc'):
                pyautogui.press('f5')
                time.sleep(180)  # 3 minutos = 180 segundos
            
            messagebox.showinfo("Información", "Script finalizado")
            
        except KeyboardInterrupt:
            messagebox.showinfo("Información", "Script detenido por el usuario")
        except Exception as e:
            messagebox.showerror("Error", f"Error durante la ejecución: {str(e)}")

def main():
    script = NSEScript()
    script.create_gui()

if __name__ == "__main__":
    main()