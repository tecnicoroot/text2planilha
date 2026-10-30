"""Pacote de conversão do Relatório de Conferências de Glosas (TXT -> XLSX)."""
from .conversor import (
    RelatorioInvalido,
    converter,
    converter_bytes,
    gerar_planilha,
)

__all__ = ["RelatorioInvalido", "converter", "converter_bytes", "gerar_planilha"]
__version__ = "1.0"
