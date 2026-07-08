"""PORTFOLIO — automacao operacional omitida."""
from __future__ import annotations
from portfolio_omitted import omit


class AutomacaoAssinaturaLotes:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs

    def iniciar(self, *args, **kwargs):
        omit("AutomacaoAssinaturaLotes.iniciar")

    def executar(self, *args, **kwargs):
        omit("AutomacaoAssinaturaLotes.executar")

    def assinar(self, *args, **kwargs):
        omit("AutomacaoAssinaturaLotes.assinar")

    def processar(self, *args, **kwargs):
        omit("AutomacaoAssinaturaLotes.processar")

    def fechar(self, *args, **kwargs):
        omit("AutomacaoAssinaturaLotes.fechar")

    def __getattr__(self, name):
        def _missing(*args, **kwargs):
            omit(f"AutomacaoAssinaturaLotes.{name}")

        return _missing
