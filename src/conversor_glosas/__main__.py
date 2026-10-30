"""Linha de comando:  python -m conversor_glosas ENTRADA.txt [SAIDA.xlsx]"""
import sys
from pathlib import Path

from .conversor import converter


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    entrada = sys.argv[1]
    saida = sys.argv[2] if len(sys.argv) > 2 else str(Path(entrada).with_suffix(".xlsx"))
    totais = converter(entrada, saida)
    print(f"Arquivo gerado: {saida}")
    for aba, qtd in totais.items():
        print(f"  {aba:.<30} {qtd}")


if __name__ == "__main__":
    main()
