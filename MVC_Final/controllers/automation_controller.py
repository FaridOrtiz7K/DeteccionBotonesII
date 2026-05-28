# controllers/automation_controller.py
import threading
import time
import os
import pandas as pd
import pyautogui
import logging
import tkinter.messagebox as messagebox
import keyboard
import tkinter as tk

from models.estado import estado_global
from models.procesador_csv import ProcesadorCSV
from models.nse_automation import NSEAutomation
from models.nse_services import NSEServicesAutomation
from models.ge_automation import GEAutomation
from utils.ahk_writer import AHKWriter
from utils.ahk_manager_save import AHKSaveManager

logger = logging.getLogger(__name__)

class AutomationController:
    def __init__(self, view):
        """
        Inicializa el controlador con la referencia a la vista.
        """
        self.view = view
        self.save_manager = AHKSaveManager()
        self.hilo_ejecucion = None

        # Variables de ejecución (antes globales)
        self.csv_file = ""
        self.kml_filename = "NN"
        self.linea_actual = 0
        self.linea_maxima = 0
        self.lote_inicio = 1
        self.lote_fin = 1
        self.info_lote_actual = {}

        # Variable para controlar la cuenta regresiva (aunque usamos estado_global)
        self.en_cuenta_regresiva = False

        # Configurar hotkeys (se llamará después de que la vista esté lista)
        self._configurar_hotkeys()

    # ----------------------------------------------------------------------
    # Configuración de hotkeys
    # ----------------------------------------------------------------------
    def _configurar_hotkeys(self):
        """Configura las teclas de acceso rápido."""
        keyboard.add_hotkey('esc', self.pausar_proceso)
        keyboard.add_hotkey('f2', self.pausar_proceso)
        keyboard.add_hotkey('f3', self.reanudar_proceso)
        keyboard.add_hotkey('f4', self.detener_proceso)

    # ----------------------------------------------------------------------
    # Métodos llamados desde la vista (eventos de usuario)
    # ----------------------------------------------------------------------
    def seleccionar_csv(self, archivo):
        """
        Valida y guarda la ruta del archivo CSV seleccionado.
        """
        if not archivo:
            return
        if not os.path.exists(archivo):
            self.view.root.after(0, messagebox.showerror, "Error", f"El archivo no existe: {archivo}")
            return
        self.csv_file = archivo
        self.view.log(f"CSV seleccionado: {archivo}")

        try:
            df = pd.read_csv(archivo)
            self.linea_maxima = len(df)
            self.view.root.after(0, self.view.linea_maxima.set, self.linea_maxima)
            self.view.root.after(0, self.view.lote_fin.set, self.linea_maxima)
            self._actualizar_estado_lineas()
            self.view.log(f"CSV cargado: {self.linea_maxima} registros encontrados")
        except Exception as e:
            self.view.log(f"Error al leer CSV: {e}")
            self.view.root.after(0, messagebox.showerror, "Error", f"No se pudo leer el CSV: {e}")

    def consultar_id(self, id_buscar):
        """
        Consulta un ID en el CSV y muestra la información.
        """
        if not self.csv_file:
            self.view.root.after(0, messagebox.showerror, "Error", "Primero seleccione un archivo CSV")
            return
        if not id_buscar:
            self.view.root.after(0, messagebox.showwarning, "Advertencia", "Ingrese un ID para consultar")
            return

        try:
            df = pd.read_csv(self.csv_file)
            resultado = df[df.iloc[:, 0].astype(str) == str(id_buscar)]

            if len(resultado) == 0:
                self.view.root.after(0, messagebox.showinfo, "Resultado", f"ID {id_buscar} no encontrado en el CSV")
            else:
                info = f"ID {id_buscar} encontrado:\n"
                for idx, col in enumerate(df.columns):
                    valor = resultado.iloc[0, idx]
                    if pd.isna(valor):
                        valor = ""
                    info += f"{col}: {valor}\n"

                if len(df.columns) > 0:
                    ultima_col = df.columns[-1]
                    valor_ultimo = resultado.iloc[0, -1]
                    if pd.isna(valor_ultimo):
                        valor_ultimo = ""
                    info += f"\n--- Resumen del Lote (última columna) ---\n"
                    info += f"{ultima_col}: {valor_ultimo}"

                self.view.root.after(0, messagebox.showinfo, "Información del ID", info)
                self.view.log(f"ID {id_buscar} consultado: {len(resultado)} coincidencias")
        except Exception as e:
            self.view.root.after(0, messagebox.showerror, "Error", f"Error al consultar ID: {e}")

    def iniciar_proceso(self, csv_file, lote_inicio, lote_fin):
        """
        Inicia el proceso de automatización en un hilo separado.
        """
        # Validaciones
        if not csv_file:
            self.view.root.after(0, messagebox.showerror, "Error", "Seleccione un archivo CSV primero")
            return
        if not os.path.exists(csv_file):
            self.view.root.after(0, messagebox.showerror, "Error", f"El archivo CSV no existe: {csv_file}")
            return
        if lote_inicio > lote_fin:
            self.view.root.after(0, messagebox.showerror, "Error", "El lote inicio no puede ser mayor que el lote fin")
            return

        self.csv_file = csv_file
        self.lote_inicio = lote_inicio
        self.lote_fin = lote_fin

        try:
            df = pd.read_csv(csv_file)
            if len(df) == 0:
                self.view.log("⚠️ CSV vacío. Continuando sin procesar datos...")
                self.linea_maxima = 0
                self.linea_actual = 0
            else:
                self.linea_maxima = min(lote_fin, len(df))
                self.linea_actual = max(1, lote_inicio)
                if self.linea_actual > self.linea_maxima:
                    self.view.root.after(0, messagebox.showerror, "Error", "El lote inicio está fuera del rango disponible")
                    return
        except Exception as e:
            self.view.log(f"❌ Error al leer CSV: {e}")
            return

        # Iniciar manager de guardado AHK
        if not self.save_manager.start_ahk():
            self.view.log("⚠️ No se pudo iniciar AutoHotkey para guardar. Continuando sin funcionalidad de guardado...")
        else:
            self.view.log("✅ AutoHotkey para guardar iniciado correctamente")

        # Configurar estado global
        estado_global.set_ejecutando(True)
        estado_global.set_pausado(False)
        estado_global.set_detener_inmediato(False)
        estado_global.set_linea_en_proceso(False)

        # Actualizar vista
        self.view.root.after(0, self.view.estado.set, "Ejecutando...")
        self.view.root.after(0, self.view.estado_label.configure, {"foreground": "green"})
        self.view.root.after(0, self._actualizar_estado_botones)

        # Iniciar hilo de ejecución
        self.hilo_ejecucion = threading.Thread(target=self._ejecutar_procesos)
        self.hilo_ejecucion.daemon = True
        self.hilo_ejecucion.start()

    def pausar_proceso(self):
        """
        Pausa el proceso en ejecución.
        """
        if estado_global.ejecutando and not estado_global.pausado:
            if estado_global.en_cuenta_regresiva:
                estado_global.set_en_cuenta_regresiva(False)
                self.view.root.after(0, self.view.log, "⏸️ Cuenta regresiva cancelada - Proceso permanece pausado")
            else:
                estado_global.set_pausado(True)
                self.view.root.after(0, self.view.estado.set, "Pausado")
                self.view.root.after(0, self.view.estado_label.configure, {"foreground": "orange"})
                self.view.root.after(0, self.view.log, "⏸️ Proceso pausado")

                # Mostrar ventana emergente de pausa
                self.view.root.after(0, self.view.mostrar_ventana_pausa, self.linea_actual, self.linea_maxima)

            self.view.root.after(0, self._actualizar_estado_botones)

    def reanudar_proceso(self):
        """
        Inicia la cuenta regresiva para reanudar el proceso.
        """
        if estado_global.ejecutando and estado_global.pausado:
            # Cerrar ventana de pausa si está abierta
            self.view.root.after(0, self.view.cerrar_ventana_pausa)

            self.view.root.after(0, self.view.btn_reanudar.config, {"state": tk.DISABLED})
            estado_global.set_en_cuenta_regresiva(True)

            self.view.root.after(0, self.view.estado.set, "Reanudando en 3 segundos...")
            self.view.root.after(0, self.view.estado_label.configure, {"foreground": "blue"})
            self.view.root.after(0, self.view.log, "Reanudando en 3 segundos...")

            # Iniciar cuenta regresiva en un hilo separado
            threading.Thread(target=self._cuenta_regresiva_reanudacion, daemon=True).start()

    def reanudar_desde_ventana(self):
        """
        Reanuda el proceso desde la ventana de pausa (sin cuenta regresiva).
        """
        estado_global.set_pausado(False)
        self.view.root.after(0, self.view.estado.set, "Ejecutando...")
        self.view.root.after(0, self.view.estado_label.configure, {"foreground": "green"})
        self.view.root.after(0, self.view.log, "▶️ Proceso reanudado desde ventana de pausa")
        self.view.root.after(0, self._actualizar_estado_botones)

    def detener_proceso(self):
        """
        Detiene el proceso inmediatamente.
        """
        # Cerrar ventana de pausa si está abierta
        self.view.root.after(0, self.view.cerrar_ventana_pausa)

        estado_global.set_detener_inmediato(True)
        estado_global.set_ejecutando(False)

        # Detener manager de guardado AHK
        self.save_manager.stop_ahk()

        self.view.root.after(0, self.view.estado.set, "Detenido")
        self.view.root.after(0, self.view.estado_label.configure, {"foreground": "red"})
        self.view.root.after(0, self.view.log, "⏹️ Proceso detenido inmediatamente")
        self.view.root.after(0, self._actualizar_estado_lineas)
        self.view.root.after(0, self._actualizar_estado_botones)

    def escribir_prueba_a(self):
        """
        Función para probar la escritura de la última columna del CSV.
        """
        if not self.csv_file:
            self.view.root.after(0, messagebox.showerror, "Error", "Primero seleccione un archivo CSV")
            return

        try:
            df = pd.read_csv(self.csv_file)
            if len(df) == 0:
                self.view.root.after(0, messagebox.showerror, "Error", "El CSV está vacío")
                return

            ultima_columna = df.columns[-1]
            if pd.isna(ultima_columna):
                texto_a_escribir = ""
            else:
                texto_a_escribir = str(ultima_columna)

            self.view.log(f"Escribiendo: '{texto_a_escribir}'")
            self.view.log("⚠️ Coloque el cursor en la posición deseada - escribiendo en 5 segundos...")

            # Contador regresivo en la UI
            for i in range(5, 0, -1):
                self.view.root.after(0, self.view.estado.set, f"Escribiendo en {i} segundos...")
                time.sleep(1)

            x, y = pyautogui.position()
            self.view.log(f"Posición del cursor: ({x}, {y})")

            ahk_writer = AHKWriter()
            if ahk_writer.start_ahk():
                exito = ahk_writer.ejecutar_escritura_ahk(x, y, texto_a_escribir)
                ahk_writer.stop_ahk()

                if exito:
                    self.view.log("✅ Texto escrito exitosamente")
                    self.view.root.after(0, self.view.estado.set, "Listo")
                else:
                    self.view.log("❌ Error al escribir texto")
                    self.view.root.after(0, self.view.estado.set, "Error")
            else:
                self.view.log("❌ No se pudo iniciar AHKWriter")
                self.view.root.after(0, self.view.estado.set, "Error")

        except Exception as e:
            self.view.log(f"❌ Error al escribir PRUEBA A: {e}")
            self.view.root.after(0, self.view.estado.set, "Error")

    def configurar_kml(self, nuevo_nombre):
        """
        Cambia el nombre base para los archivos KML.
        """
        if nuevo_nombre:
            self.kml_filename = nuevo_nombre
            self.view.log(f"Nombre KML configurado a: {self.kml_filename}")

    def guardar_progreso_manual(self):
        """
        Guarda el progreso actual usando AHK o pyautogui.
        """
        try:
            self.view.log("💾 Guardando progreso manualmente...")
            if not self.save_manager.start_ahk():
                self.view.log("❌ No se pudo iniciar AHK para guardar")
                return False

            if not self.save_manager.trigger_save():
                self.view.log("❌ No se pudo guardar con AHK")
                return False

            self.view.log("✅ Progreso guardado exitosamente con AHK")
            return True
        except Exception as e:
            self.view.log(f"❌ Error al guardar progreso: {e}")
            return False

    def mostrar_estado_actual(self):
        """
        Muestra un cuadro de diálogo con el estado actual.
        """
        if estado_global.ejecutando:
            lineas_restantes = self.linea_maxima - self.linea_actual
            estado_linea = "EN PROCESO" if estado_global.linea_en_proceso else "ESPERANDO"
            mensaje = f"Línea actual: {self.linea_actual}\nLíneas restantes: {lineas_restantes}\nEstado línea: {estado_linea}"

            if self.info_lote_actual:
                mensaje += f"\n\nInfo Lote Actual:\n{self.view.info_lote.get()}"

            self.view.root.after(0, messagebox.showinfo, "Estado Actual", mensaje)

    # ----------------------------------------------------------------------
    # Métodos auxiliares para actualizar la vista
    # ----------------------------------------------------------------------
    def _actualizar_estado_lineas(self):
        """Actualiza las etiquetas de línea actual y líneas restantes en la vista."""
        self.view.root.after(0, self.view.linea_actual.set, str(self.linea_actual))
        lineas_rest = max(0, self.linea_maxima - self.linea_actual)
        self.view.root.after(0, self.view.lineas_restantes.set, str(lineas_rest))

    def _actualizar_estado_botones(self):
        """Actualiza el estado de los botones en la vista según el estado global."""
        estado_actual = self.view.estado.get()
        if "Reanudando en" in estado_actual:
            self.view.root.after(0, self.view.actualizar_estado_botones, True, False, True)
            return

        self.view.root.after(0, self.view.actualizar_estado_botones,
                             estado_global.ejecutando,
                             estado_global.pausado,
                             estado_global.en_cuenta_regresiva)

    def _actualizar_info_lote(self, linea_actual, datos):
        """
        Actualiza la información del lote en la vista.
        """
        self.info_lote_actual = {
            'linea': linea_actual,
            'datos': datos if datos is not None else {},
            'timestamp': time.time()
        }

        if datos is not None and hasattr(datos, 'iloc') and len(datos) > 0:
            id_valor = datos.iloc[0] if len(datos) > 0 else ''
            if pd.isna(id_valor):
                id_valor = ""

            info_text = f"Línea {linea_actual}: ID={id_valor}"

            if len(datos) > 1:
                col2_valor = datos.iloc[1]
                if pd.isna(col2_valor):
                    col2_valor = ""
                if col2_valor:
                    info_text += f", Col2={col2_valor}"

            if len(datos) > 3:
                col4_valor = datos.iloc[3]
                if pd.isna(col4_valor):
                    col4_valor = ""
                if col4_valor:
                    info_text += f", Col4={col4_valor}"

            if len(datos) > 0:
                ultima_col_valor = datos.iloc[-1]
                if pd.isna(ultima_col_valor):
                    ultima_col_valor = ""
                if ultima_col_valor:
                    info_text += f", Resumen={ultima_col_valor}"

            self.view.root.after(0, self.view.info_lote.set, info_text)
        else:
            self.view.root.after(0, self.view.info_lote.set, f"Línea {linea_actual}: Datos vacíos o no encontrados")

    # ----------------------------------------------------------------------
    # Métodos de cuenta regresiva para reanudación
    # ----------------------------------------------------------------------
    def _cuenta_regresiva_reanudacion(self):
        """Ejecuta la cuenta regresiva de 3 segundos para reanudar."""
        try:
            for i in range(3, 0, -1):
                if estado_global.detener_inmediato:
                    self.view.root.after(0, self._cancelar_reanudacion, "Reanudación cancelada (proceso detenido)")
                    return

                self.view.root.after(0, self.view.estado.set, f"Reanudando en {i} segundos...")
                time.sleep(1)

            if not estado_global.detener_inmediato:
                self.view.root.after(0, self._completar_reanudacion)
            else:
                self.view.root.after(0, self._cancelar_reanudacion,
                                    "Reanudación cancelada (proceso detenido durante la cuenta regresiva)")
        except Exception as e:
            self.view.root.after(0, self._cancelar_reanudacion, f"Error en cuenta regresiva: {e}")

    def _completar_reanudacion(self):
        """Completa la reanudación después de la cuenta regresiva."""
        estado_global.set_en_cuenta_regresiva(False)
        estado_global.set_pausado(False)
        self.view.root.after(0, self.view.estado.set, "Ejecutando...")
        self.view.root.after(0, self.view.estado_label.configure, {"foreground": "green"})
        self.view.root.after(0, self.view.log, "▶️ Proceso reanudado")
        self.view.root.after(0, self._actualizar_estado_botones)

    def _cancelar_reanudacion(self, mensaje):
        """Cancela la reanudación y actualiza la UI."""
        estado_global.set_en_cuenta_regresiva(False)
        self.view.root.after(0, self.view.log, mensaje)
        self.view.root.after(0, self._actualizar_estado_botones)

    # ----------------------------------------------------------------------
    # Métodos de ejecución de los programas (llamados desde el hilo)
    # ----------------------------------------------------------------------
    def _ejecutar_procesos(self):
        """
        Bucle principal de ejecución. Se ejecuta en un hilo separado.
        """
        lotes_desde_ultimo_guardado = 0

        try:
            while self.linea_actual <= self.linea_maxima and estado_global.verificar_continuar():
                if estado_global.esperar_si_pausado():
                    break

                estado_global.set_linea_en_proceso(True)

                self.view.root.after(0, self.view.log, f"🔄 Procesando línea {self.linea_actual}/{self.linea_maxima}")
                self._actualizar_estado_lineas()

                # --- Programa 1 ---
                self.view.root.after(0, self.view.log, "Iniciando Programa 1 - Procesador CSV")
                resultado1, linea_procesada, datos_lote = self._ejecutar_programa1(self.linea_actual)

                if not estado_global.verificar_continuar():
                    break

                self._actualizar_info_lote(self.linea_actual, datos_lote)

                if not resultado1 or linea_procesada is None:
                    self.view.root.after(0, self.view.log, f"⚠️ Programa 1 no procesó la línea {self.linea_actual} (ID no encontrado después de 2 intentos). Saltando al siguiente lote...")
                                        
                    estado_global.set_linea_en_proceso(False)
                    self.linea_actual += 1
                    lotes_desde_ultimo_guardado += 1
                    continue


                if not estado_global.verificar_continuar():
                    break

                # --- Programa 2 ---
                self.view.root.after(0, self.view.log, "Iniciando Programa 2 - Automatización NSE")
                resultado2 = self._ejecutar_programa2(linea_procesada)

                if not resultado2:
                    self.view.root.after(0, self.view.log, f"❌ Programa 2 falló en línea {self.linea_actual}")
                    estado_global.set_linea_en_proceso(False)
                    self.linea_actual += 1
                    lotes_desde_ultimo_guardado += 1
                    continue

                if not estado_global.verificar_continuar():
                    break

                # --- Programa 3 ---
                self.view.root.after(0, self.view.log, "Iniciando Programa 3 - Servicios NSE")
                resultado3 = self._ejecutar_programa3(linea_procesada)

                if not resultado3:
                    self.view.root.after(0, self.view.log, f"❌ Programa 3 falló en línea {self.linea_actual}")
                    estado_global.set_linea_en_proceso(False)
                    self.linea_actual += 1
                    lotes_desde_ultimo_guardado += 1
                    continue

                if not estado_global.verificar_continuar():
                    break

                # --- Programa 4 ---
                self.view.root.after(0, self.view.log, "Iniciando Programa 4 - Automatización GE")
                resultado4 = self._ejecutar_programa4(linea_procesada)

                if resultado4:
                    self.view.root.after(0, self.view.log, f"✅ Línea {self.linea_actual} procesada exitosamente")
                else:
                    self.view.root.after(0, self.view.log, f"⚠️ Línea {self.linea_actual} completada con advertencias")

                estado_global.set_linea_en_proceso(False)
                lotes_desde_ultimo_guardado += 1

                # Guardar cada 10 lotes
                if lotes_desde_ultimo_guardado >= 10:
                    self.view.root.after(0, self.view.log, "📁 Guardando progreso después de 10 lotes...")
                    if self.save_manager.is_running:
                        if not self.save_manager.trigger_save():
                            self.view.root.after(0, self.view.log,
                                                 "⚠️ No se pudo guardar con AHK. Usando pyautogui como respaldo...")
                            pyautogui.hotkey('ctrl', 's')
                    else:
                        self.view.root.after(0, self.view.log, "⚠️ AHK no está corriendo. Usando pyautogui...")
                        pyautogui.hotkey('ctrl', 's')

                    for _ in range(6):
                        if estado_global.esperar_si_pausado():
                            break
                        time.sleep(0.3)
                    self.view.root.after(0, self.view.log, "✅ Progreso guardado exitosamente")
                    lotes_desde_ultimo_guardado = 0

                self.linea_actual += 1
                self._actualizar_estado_lineas()

                # Espera entre lotes
                for _ in range(4):
                    if estado_global.esperar_si_pausado():
                        break
                    time.sleep(1)

            # Guardado final
            if not estado_global.detener_inmediato and self.linea_actual > self.linea_maxima:
                self.view.root.after(0, self.view.log, "📁 Guardando progreso final al completar todos los lotes...")
                if self.save_manager.is_running:
                    if not self.save_manager.trigger_save():
                        self.view.root.after(0, self.view.log,
                                             "⚠️ No se pudo guardar con AHK. Usando pyautogui como respaldo...")
                        pyautogui.hotkey('ctrl', 's')
                else:
                    self.view.root.after(0, self.view.log, "⚠️ AHK no está corriendo. Usando pyautogui...")
                    pyautogui.hotkey('ctrl', 's')

                time.sleep(6)
                self.view.root.after(0, self.view.log, "✅ Progreso final guardado exitosamente")

            # Mensaje final
            if estado_global.detener_inmediato:
                self.view.root.after(0, self.view.log, "🛑 Proceso detenido inmediatamente por usuario")
                self.view.root.after(0, self.view.estado.set, "Detenido")
                self.view.root.after(0, self.view.estado_label.configure, {"foreground": "red"})
            elif estado_global.ejecutando and self.linea_actual > self.linea_maxima:
                self.view.root.after(0, self.view.log, "🎉 Proceso completado exitosamente")
                self.view.root.after(0, self.view.estado.set, "Completado")
                self.view.root.after(0, self.view.estado_label.configure, {"foreground": "green"})
            elif not estado_global.ejecutando:
                self.view.root.after(0, self.view.log, "Proceso detenido por el usuario")

        except Exception as e:
            self.view.root.after(0, self.view.log, f"❌ Error en ejecución: {e}")
            self.view.root.after(0, self.view.estado.set, "Error")
            self.view.root.after(0, self.view.estado_label.configure, {"foreground": "red"})
            estado_global.set_linea_en_proceso(False)

        finally:
            estado_global.set_ejecutando(False)
            self.save_manager.stop_ahk()
            self.view.root.after(0, self._actualizar_estado_botones)
            self.view.root.after(0, self.view.cerrar_ventana_pausa)

    # ----------------------------------------------------------------------
    # Ejecutores específicos de cada programa
    # ----------------------------------------------------------------------
    def _ejecutar_programa1(self, linea):
        """
        Ejecuta el Programa 1 (ProcesadorCSV) para la línea dada.
        Retorna (exito, linea_procesada, datos_lote)
        """
        try:
            logger.info("=" * 60)
            logger.info("INICIANDO PROGRAMA 1 - PROCESADOR CSV")
            logger.info("=" * 60)

            procesador = ProcesadorCSV(self.csv_file)

            if estado_global.esperar_si_pausado():
                return False, None, None
            time.sleep(0.5)

            if not procesador.cargar_csv():
                logger.warning("⚠️ CSV vacío detectado. Continuando sin procesar...")
                return True, None, None

            if linea > len(procesador.df):
                logger.warning(f"⚠️ Línea {linea} no existe en CSV. Saltando...")
                return True, None, None

            # Filtrar solo la línea actual
            procesador.df = procesador.df.iloc[linea-1:linea]
            datos_lote = procesador.df.iloc[0] if len(procesador.df) > 0 else None

            resultado, linea_procesada = procesador.procesar_todo()

            if resultado and linea_procesada:
                logger.info(f"✅ Programa 1 completado. Línea procesada: {linea_procesada}")
                return True, linea_procesada, datos_lote
            else:
                logger.error("❌ Programa 1 falló o no encontró ID")
                return False, None, datos_lote

        except Exception as e:
            logger.error(f"❌ Error en Programa 1: {e}")
            return False, None, None

    def _ejecutar_programa2(self, linea_procesada):
        """
        Ejecuta el Programa 2 (NSEAutomation) para la línea procesada.
        """
        try:
            logger.info("\n" + "=" * 60)
            logger.info("INICIANDO PROGRAMA 2 - AUTOMATIZACIÓN NSE")
            logger.info("=" * 60)

            nse = NSEAutomation(self.csv_file, linea_especifica=linea_procesada)
            nse.is_running = True

            if not os.path.exists(nse.csv_file):
                logger.warning(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
                return True

            try:
                df = pd.read_csv(nse.csv_file)
                if len(df) == 0:
                    logger.warning("⚠️ CSV vacío. Saltando Programa 2...")
                    return True
            except:
                pass

            logger.info(f"🎯 Procesando línea: {linea_procesada}")

            for _ in range(1):
                if estado_global.esperar_si_pausado():
                    return False
                time.sleep(0.5)

            resultado = nse.execute_nse_script()

            if resultado:
                logger.info("✅ Programa 2 finalizado exitosamente")
            else:
                logger.warning("⚠️ Programa 2 completado con advertencias")

            return resultado

        except Exception as e:
            logger.error(f"❌ Error en Programa 2: {e}")
            return False

    def _ejecutar_programa3(self, linea_procesada):
        """
        Ejecuta el Programa 3 (NSEServicesAutomation) para la línea procesada.
        """
        try:
            logger.info("\n" + "=" * 60)
            logger.info("INICIANDO PROGRAMA 3 - SERVICIOS NSE")
            logger.info("=" * 60)

            nse_services = NSEServicesAutomation(self.csv_file, linea_especifica=linea_procesada)

            if not os.path.exists(nse_services.csv_file):
                logger.warning(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
                return True

            logger.info(f"🎯 Procesando línea: {linea_procesada}")

            if not nse_services.iniciar_ahk():
                logger.warning("⚠️ No se pudieron iniciar los servicios AHK. Continuando...")
                return True

            nse_services.is_running = True
            resultado = nse_services.procesar_linea_especifica()

            if resultado:
                logger.info(f"✅ Programa 3 completado exitosamente")
            else:
                logger.warning(f"⚠️ Programa 3 completado con advertencias")

            nse_services.detener_ahk()
            return resultado

        except Exception as e:
            logger.error(f"❌ Error en Programa 3: {e}")
            return False

    def _ejecutar_programa4(self, linea_procesada):
        """
        Ejecuta el Programa 4 (GEAutomation) para la línea procesada.
        """
        try:
            logger.info("\n" + "=" * 60)
            logger.info("INICIANDO PROGRAMA 4 - AUTOMATIZACIÓN GE")
            logger.info("=" * 60)

            ge_auto = GEAutomation(self.csv_file, linea_especifica=linea_procesada, kml_base_name=self.kml_filename)
            ge_auto.is_running = True

            if not os.path.exists(ge_auto.csv_file):
                logger.warning(f"⚠️ Archivo CSV no encontrado. Continuando sin procesar...")
                return True

            try:
                df = pd.read_csv(ge_auto.csv_file)
                if len(df) == 0:
                    logger.warning("⚠️ CSV vacío. Saltando Programa 4...")
                    return True
            except:
                pass

            logger.info(f"🎯 Procesando línea: {linea_procesada}")
            logger.info(f"📁 Archivo KML base: {ge_auto.kml_base_name}")

            for _ in range(1):
                if estado_global.esperar_si_pausado():
                    return False
                time.sleep(0.5)

            success = ge_auto.perform_actions()

            if success:
                logger.info("✅ Programa 4 finalizado exitosamente")
            else:
                logger.warning("⚠️ Programa 4 completado con advertencias")

            return success

        except Exception as e:
            logger.error(f"❌ Error en Programa 4: {e}")
            return False