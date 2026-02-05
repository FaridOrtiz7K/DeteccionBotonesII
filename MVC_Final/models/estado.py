import threading

class EstadoEjecucion:
    def __init__(self):
        self.ejecutando = False
        self.pausado = False
        self.detener_inmediato = False
        self.linea_en_proceso = False
        self.en_cuenta_regresiva = False
        self.condition = threading.Condition()
    
    def set_ejecutando(self, valor):
        with self.condition:
            self.ejecutando = valor
            if not valor:
                self.pausado = False
                self.linea_en_proceso = False
                self.en_cuenta_regresiva = False
            self.condition.notify_all()
    
    def set_pausado(self, valor):
        with self.condition:
            self.pausado = valor
            self.condition.notify_all()
    
    def set_detener_inmediato(self, valor):
        with self.condition:
            self.detener_inmediato = valor
            if valor:
                self.ejecutando = False
                self.pausado = False
                self.en_cuenta_regresiva = False
            self.condition.notify_all()
    
    def set_en_cuenta_regresiva(self, valor):
        with self.condition:
            self.en_cuenta_regresiva = valor
            self.condition.notify_all()
    
    def set_linea_en_proceso(self, valor):
        with self.condition:
            self.linea_en_proceso = valor
    
    def esperar_si_pausado(self):
        with self.condition:
            while (self.pausado or self.en_cuenta_regresiva) and self.ejecutando and not self.detener_inmediato:
                self.condition.wait()
            return not self.ejecutando or self.detener_inmediato
    
    def verificar_continuar(self):
        with self.condition:
            return self.ejecutando and not self.detener_inmediato and not self.en_cuenta_regresiva