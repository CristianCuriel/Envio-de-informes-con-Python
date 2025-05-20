

class StatusReporter:
    def __init__(self, status: list[tuple[bool, str]]):
        self.status = status

    def view_satus_report(self):
        for i, (exito, mensaje) in enumerate(self.status, 1):
            estado_str = "✅ Éxito" if exito else "❌ Error"
            print(f"{i}. {estado_str} - {mensaje}")

