#!/usr/bin/env python3
"""
Converte o TXT do "Relatório de Conferências de Glosas" em uma planilha XLSX
com as abas: RESUMO, DETALHAMENTO, ITENS COM VALOR DIFERENTE, MAPA DE SELEÇÃO e Glosa.

Uso:
    python conferencia_glosas_xlsx.py ENTRADA.txt [SAIDA.xlsx]

Requisito:  pip install openpyxl
"""
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
RE_NUM_BR = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d+|-?\d+,\d+")
RE_DATA = re.compile(r"^\d{2}/\d{2}/\d{4}$")
RE_DATAHORA = re.compile(r"^\d{2}/\d{2}/\d{4} \d{2}:\d{2}$")


def ler_linhas(caminho):
    """Lê o TXT tentando UTF-8 e, se falhar, Latin-1/CP1252."""
    dados = Path(caminho).read_bytes()
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            texto = dados.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    # remove caracteres de controle (ex.: \x02) que o Excel não aceita
    texto = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", texto)
    return [l.rstrip("\r") for l in texto.split("\n")]


def num_br(txt):
    """'1.234,56' -> 1234.56 ; vazio -> None"""
    txt = (txt or "").strip()
    if not txt:
        return None
    try:
        return float(txt.replace(".", "").replace(",", "."))
    except ValueError:
        return txt


def valor(txt):
    """Converte texto em data, inteiro ou mantém texto."""
    txt = (txt or "").strip()
    if txt == "":
        return None
    if RE_DATA.match(txt):
        return datetime.strptime(txt, "%d/%m/%Y")
    if RE_DATAHORA.match(txt):
        return datetime.strptime(txt, "%d/%m/%Y %H:%M")
    # inteiro sem zero à esquerda (evita perder zeros de códigos)
    if txt.isdigit() and (txt == "0" or not txt.startswith("0")):
        return int(txt)
    return txt


def posicoes(cabecalho, rotulos):
    """Posição inicial de cada rótulo na linha de cabeçalho."""
    pos, inicio = [], 0
    for r in rotulos:
        p = cabecalho.index(r, inicio)
        pos.append(p)
        inicio = p + len(r)
    return pos


def fatiar(linha, pos):
    """Fatia a linha de largura fixa conforme as posições."""
    campos = []
    for i, p in enumerate(pos):
        fim = pos[i + 1] if i + 1 < len(pos) else None
        campos.append(linha[p:fim].strip() if fim else linha[p:].strip())
    return campos


# --------------------------------------------------------------------------- #
# Seções
# --------------------------------------------------------------------------- #
def localizar_secoes(linhas):
    idx = {}
    for i, l in enumerate(linhas):
        u = l.upper()
        if "CONFERÊNCIAS DE GLOSAS - RESUMO" in u or "CONFERENCIAS DE GLOSAS - RESUMO" in u:
            idx["resumo"] = i
        elif "CONFERÊNCIAS DE GLOSAS - DETALHAMENTO" in u or "CONFERENCIAS DE GLOSAS - DETALHAMENTO" in u:
            idx["detalhe"] = i
        elif "ITENS COM VALOR DIFERENTE DA TABELA" in u and l.strip().startswith("---"):
            idx["valor_dif"] = i
        elif "MAPA DE SELE" in u and l.strip().startswith("---"):
            idx["mapa"] = i
        elif re.match(r"^\s*-+\s*Glosa\s*-+\s*$", l):
            idx["glosa"] = i
    ordem = sorted(idx.items(), key=lambda kv: kv[1])
    faixas = {}
    for n, (nome, ini) in enumerate(ordem):
        fim = ordem[n + 1][1] if n + 1 < len(ordem) else len(linhas)
        faixas[nome] = linhas[ini + 1:fim]
    return faixas


# ---------------------------- RESUMO --------------------------------------- #
CAB_RESUMO = ["ANO", "PER", "GUIA SGU", "ID GUIA", "PRINCIPAL", "UNID",
              "CARTEIRA", "NOME", "VL.COBRADO", "VL.PRINCIPAL", "VL.DIFERENÇA"]
RE_LINHA_RESUMO = re.compile(r"^\s*\d{4}\s+\d{1,2}\s+\d+")


def parse_resumo(linhas):
    pos = None
    saida = []
    for l in linhas:
        if pos is None and "ANO" in l and "GUIA SGU" in l and "VL.COBRADO" in l:
            pos = posicoes(l, ["ANO", "PER", "GUIA SGU", "ID GUIA",
                               "PRINCIPAL", "UNID", "CARTEIRA", "NOME"])
            continue
        if pos is None or not RE_LINHA_RESUMO.match(l):
            continue
        nums = list(RE_NUM_BR.finditer(l))[-3:]          # 3 valores finais
        corte = nums[0].start() if nums else len(l)
        campos = fatiar(l[:corte], pos)
        saida.append([valor(c) for c in campos] + [num_br(m.group()) for m in nums])
    return saida


# --------------------------- DETALHAMENTO ---------------------------------- #
CAB_DETALHE = ["ANO", "PER", "GUIA SGU", "ID GUIA", "AUDITORIA", "INTEGRAÇÃO",
               "PRINCIPAL", "UNID", "CARTEIRA", "NOME", "DT.INTERNACAO", "DT.ALTA",
               "MOTIVO.ALTA", "SOLICITANTE", "ACOMODACAO", "HOSP.DIA", "CID",
               "GLOSA", "MANUAL", "COD.ITEM", "TP.ITEM.VERSAO", "DESCRICAO",
               "QT.COBR", "QT.PAGA", "VL.COBR", "VL.PAGO", "ERRO", "PRESTADOR",
               "NIVEL", "DT.REALZ", "HR.REALZ", "JUSTIFICATIVA", "AUDITOR",
               "SITUAÇÃO"]   # coluna extra: GLOSADO / ITEM AJUSTADO

ROT_GUIA = ["ANO", "PER", "GUIA SGU", "ID GUIA", "AUDITORIA", "INTEGRA",
            "PRINCIPAL", "UNID", "CARTEIRA", "NOME"]
ROT_INTERN = ["DT.INTERNACAO", "DT.ALTA", "MOTIVO.ALTA", "SOLICITANTE",
              "ACOMODACAO", "HOSP.DIA", "CID"]
ROT_ITEM_ESQ = ["GLOSA", "MANUAL", "COD.ITEM", "TP.ITEM.VERSAO", "DESCRICAO"]
ROT_ITEM_DIR = ["ERRO", "PRESTADOR", "NIVEL", "DT.REALZ", "HR.REALZ",
                "JUSTIFICATIVA", "AUDITOR"]


def parse_detalhe(linhas):
    saida = []
    guia = [None] * len(ROT_GUIA)
    intern = [None] * len(ROT_INTERN)
    pos_guia = pos_int = None
    pos_esq = pos_dir = None
    largura_desc = 40
    espera = None          # "guia" | "intern" -> próxima linha é de dados
    ajustado = False

    for l in linhas:
        s = l.strip()
        if not s or set(s) <= {"=", "-"}:
            continue

        if s.startswith("ANO") and "AUDITORIA" in l:
            pos_guia = posicoes(l, ROT_GUIA)
            espera, ajustado = "guia", False
            intern = [None] * len(ROT_INTERN)
            continue
        if s.startswith("DT.INTERNACAO"):
            pos_int = posicoes(l, ROT_INTERN)
            espera = "intern"
            continue
        if s.startswith("GLOSA") and "COD.ITEM" in l:
            pos_esq = posicoes(l, ROT_ITEM_ESQ)
            pos_dir = posicoes(l, [r for r in ROT_ITEM_DIR if r in l])
            # alguns blocos vêm com o cabeçalho cortado (sem AUDITOR)
            if "AUDITOR" not in l and "JUSTIFICATIVA" in l:
                pos_dir.append(l.index("JUSTIFICATIVA") + 51)
            largura_desc = l.index("QT.COBR") - pos_esq[-1] - 5
            espera = None
            continue
        if "Itens Ajustados" in l:
            ajustado = True
            continue

        if espera == "guia":
            guia = [valor(c) for c in fatiar(l, pos_guia)]
            espera = None
            continue
        if espera == "intern":
            intern = [valor(c) for c in fatiar(l, pos_int)]
            espera = None
            continue

        # Linha de item: tem COD.ITEM preenchido e 4 valores numéricos
        if pos_esq is None or not l[pos_esq[2]:pos_esq[3]].strip():
            continue
        ini_num = pos_esq[-1] + largura_desc
        fim_num = pos_dir[0]
        nums = RE_NUM_BR.findall(l[ini_num:fim_num])
        if len(nums) != 4:
            continue
        esq = fatiar(l[:ini_num], pos_esq)
        dir_ = fatiar(l, pos_dir)
        dir_ += [""] * (len(ROT_ITEM_DIR) - len(dir_))   # itens ajustados não têm justificativa/auditor

        linha = (guia + intern
                 + [valor(c) if i != 4 else (c or None) for i, c in enumerate(esq)]
                 + [num_br(n) for n in nums]
                 + [valor(c) if i not in (4, 5, 6) else (c or None) for i, c in enumerate(dir_)]
                 + ["ITEM AJUSTADO" if ajustado else "GLOSADO"])
        saida.append(linha)
    return saida


# ------------------- ITENS COM VALOR DIFERENTE ----------------------------- #
CAB_VALOR_DIF = ["COD.ITEM", "DESC.ITEM", "VL.INFOR", "VL.PAGO", "DT.EXECUÇÃO"]


def parse_valor_dif(linhas):
    pos, saida = None, []
    for l in linhas:
        if pos is None and "COD.ITEM" in l and "VL.INFOR" in l:
            pos = posicoes(l, ["COD.ITEM", "DESC.ITEM", "VL.INFOR", "VL.PAGO", "DT.EXEC"])
            continue
        if pos is None or not l.strip() or set(l.strip()) <= {"-", "="}:
            continue
        c = fatiar(l, pos)
        if not c[0]:
            continue
        saida.append([valor(c[0]), c[1], num_br(c[2]), num_br(c[3]), valor(c[4])])
    return saida


# --------------------------- MAPA DE SELEÇÃO ------------------------------- #
def parse_mapa(linhas):
    saida = []
    for l in linhas:
        if ":" not in l:
            continue
        chave, resto = l.split(":", 1)
        val = resto.strip()
        while val.startswith(":"):                 # "Datas de Integração : : ..."
            val = val[1:].strip()
        v = val if val else None
        if v and v.isdigit():
            v = int(v)
        saida.append([chave.strip(), v])
    return saida


# -------------------------------- GLOSA ------------------------------------ #
def parse_glosa(linhas):
    saida = []
    for l in linhas:
        m = re.match(r"^\s*(\d+)\s+(.+?)\s*$", l)
        if m:
            saida.append([int(m.group(1)), m.group(2)])
    return saida


# --------------------------------------------------------------------------- #
# Escrita do XLSX
# --------------------------------------------------------------------------- #
FONTE = Font(name="Arial", size=10)
FONTE_CAB = Font(name="Arial", size=10, bold=True, color="FFFFFF")
FUNDO_CAB = PatternFill("solid", fgColor="1F4E78")
FMT_MOEDA = "#,##0.00"
FMT_DATA = "DD/MM/YYYY"
FMT_DATAHORA = "DD/MM/YYYY HH:MM"


def escrever_aba(wb, titulo, cabecalho, linhas, col_num=(), largura_max=50):
    ws = wb.create_sheet(titulo)
    if cabecalho:
        ws.append(cabecalho)
        for c in ws[1]:
            c.font, c.fill = FONTE_CAB, FUNDO_CAB
            c.alignment = Alignment(horizontal="center", vertical="center")
    for lin in linhas:
        ws.append(lin)

    larguras = {}
    inicio = 2 if cabecalho else 1
    for row in ws.iter_rows(min_row=inicio):
        for c in row:
            c.font = FONTE
            v = c.value
            if isinstance(v, datetime):
                c.number_format = FMT_DATAHORA if (v.hour or v.minute) else FMT_DATA
                tam = 16 if (v.hour or v.minute) else 11
            elif isinstance(v, float) or c.column in col_num:
                c.number_format = FMT_MOEDA
                tam = len(f"{v:,.2f}") if isinstance(v, (int, float)) else 8
            else:
                tam = len(str(v)) if v is not None else 0
            larguras[c.column] = max(larguras.get(c.column, 0), tam)
    if cabecalho:
        for i, h in enumerate(cabecalho, 1):
            larguras[i] = max(larguras.get(i, 0), len(h))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    for col, w in larguras.items():
        ws.column_dimensions[get_column_letter(col)].width = min(w + 2, largura_max)
    return ws


def converter(entrada, saida):
    linhas = ler_linhas(entrada)
    sec = localizar_secoes(linhas)

    resumo = parse_resumo(sec.get("resumo", []))
    detalhe = parse_detalhe(sec.get("detalhe", []))
    valor_dif = parse_valor_dif(sec.get("valor_dif", []))
    mapa = parse_mapa(sec.get("mapa", []))
    glosa = parse_glosa(sec.get("glosa", []))

    wb = Workbook()
    wb.remove(wb.active)
    escrever_aba(wb, "RESUMO", CAB_RESUMO, resumo)
    escrever_aba(wb, "DETALHAMENTO", CAB_DETALHE, detalhe)
    escrever_aba(wb, "ITENS COM VALOR DIFERENTE", CAB_VALOR_DIF, valor_dif)
    escrever_aba(wb, "MAPA DE SELEÇÃO", None, mapa, largura_max=55)
    escrever_aba(wb, "Glosa", None, glosa, largura_max=70)
    wb.save(saida)

    return {
        "RESUMO": len(resumo),
        "DETALHAMENTO": len(detalhe),
        "ITENS COM VALOR DIFERENTE": len(valor_dif),
        "MAPA DE SELEÇÃO": len(mapa),
        "Glosa": len(glosa),
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    entrada = sys.argv[1]
    saida = sys.argv[2] if len(sys.argv) > 2 else str(Path(entrada).with_suffix(".xlsx"))
    totais = converter(entrada, saida)
    print(f"Arquivo gerado: {saida}")
    for aba, qtd in totais.items():
        print(f"  {aba:.<30} {qtd}")