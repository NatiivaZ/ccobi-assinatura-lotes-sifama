"""Logger usado na automação de assinatura de lotes."""

import os
from datetime import datetime


class Logger:
    """Mantém o log passando pela tela e, quando configurado, também por arquivo."""

    def __init__(self, callback=None, log_file: str = None):
        self.callback = callback
        self.log_file = log_file
        if log_file:
            pasta = os.path.dirname(log_file)
            if pasta:
                os.makedirs(pasta, exist_ok=True)
            with open(log_file, "w", encoding="utf-8") as arquivo_log:
                arquivo_log.write("Log de Automação — Assinatura de Lotes DOU\n")
                arquivo_log.write(f"Iniciado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                arquivo_log.write("=" * 60 + "\n\n")

    def log(self, mensagem: str, tipo: str = "INFO"):
        """Registra a mensagem no formato que a interface já espera."""
        linha = f"[{datetime.now().strftime('%H:%M:%S')}] [{tipo}] {mensagem}"
        print(linha)
        if self.callback:
            self.callback(linha, tipo)
        if self.log_file:
            try:
                with open(self.log_file, "a", encoding="utf-8") as arquivo_log:
                    arquivo_log.write(linha + "\n")
            except Exception:
                pass
