# -*- coding: utf-8 -*-
"""
Conversor do "Relatório de Conferências de Glosas" (TXT) para XLSX.

Lógica pura, sem interface: usada pelo app desktop e por qualquer outra
interface (Flask, linha de comando etc.).

Uso pela linha de comando (a partir da pasta do projeto):
    python -m conversor_glosas ENTRADA.txt [SAIDA.xlsx]
"""
import re
from datetime import datetime
from io import BytesIO
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


def ler_linhas(dados):
    """Decodifica o conteúdo do TXT (bytes) tentando UTF-8 e, se falhar, CP1252/Latin-1."""
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


def mapear(cabecalho, rotulos):
    """
    Procura cada rótulo na linha de cabeçalho e devolve [(nome, posição), ...]
    em ordem de posição. Não depende da ordem das colunas no relatório, e
    rótulos que não existem no cabeçalho são simplesmente ignorados.

    rotulos: {nome_da_coluna: [grafia1, grafia2, ...]}
    """
    achados = []
    for nome, grafias in rotulos.items():
        for g in grafias:
            m = re.search(r"(?<!\S)" + re.escape(g) + r"(?!\S)", cabecalho)
            if m:
                achados.append((nome, m.start()))
                break
    return sorted(achados, key=lambda x: x[1])


def fatiar(linha, mapa):
    """Recorta a linha de largura fixa conforme o mapa e devolve {nome: texto}."""
    campos = {}
    for i, (nome, p) in enumerate(mapa):
        fim = mapa[i + 1][1] if i + 1 < len(mapa) else None
        campos[nome] = linha[p:fim].strip()
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
ROT_RESUMO = {n: [n] for n in CAB_RESUMO[:8]}
RE_LINHA_RESUMO = re.compile(r"^\s*\d{4}\s+\d{1,2}\s+\d+")


def parse_resumo(linhas):
    mapa = None
    saida = []
    for l in linhas:
        if mapa is None and "GUIA SGU" in l and "VL.COBRADO" in l:
            mapa = mapear(l, ROT_RESUMO)
            continue
        if mapa is None or not RE_LINHA_RESUMO.match(l):
            continue
        nums = list(RE_NUM_BR.finditer(l))[-3:]          # 3 valores finais
        corte = nums[0].start() if nums else len(l)
        c = fatiar(l[:corte], mapa)
        saida.append([valor(c.get(n)) for n in CAB_RESUMO[:8]]
                     + [num_br(m.group()) for m in nums])
    return saida


# --------------------------- DETALHAMENTO ---------------------------------- #
CAB_DETALHE = ["ANO", "PER", "GUIA SGU", "ID GUIA", "AUDITORIA", "INTEGRAÇÃO",
               "PRINCIPAL", "UNID", "CARTEIRA", "NOME", "DT.INTERNACAO", "DT.ALTA",
               "MOTIVO.ALTA", "SOLICITANTE", "ACOMODACAO", "HOSP.DIA", "CID",
               "GLOSA", "MANUAL", "COD.ITEM", "TP.ITEM.VERSAO", "DESCRICAO",
               "QT.COBR", "QT.PAGA", "VL.COBR", "VL.PAGO", "ERRO", "PRESTADOR",
               "NIVEL", "DT.REALZ", "HR.REALZ", "JUSTIFICATIVA", "AUDITOR",
               "SITUAÇÃO"]   # coluna extra: GLOSADO / ITEM AJUSTADO

ROT_GUIA = {"ANO": ["ANO"], "PER": ["PER"], "GUIA SGU": ["GUIA SGU"], "ID GUIA": ["ID GUIA"],
            "AUDITORIA": ["AUDITORIA"], "INTEGRAÇÃO": ["INTEGRAÇÃO", "INTEGRACAO"],
            "PRINCIPAL": ["PRINCIPAL"], "UNID": ["UNID"], "CARTEIRA": ["CARTEIRA"],
            "NOME": ["NOME"]}
ROT_INTERN = {n: [n] for n in ["DT.INTERNACAO", "DT.ALTA", "MOTIVO.ALTA", "SOLICITANTE",
                               "ACOMODACAO", "HOSP.DIA", "CID"]}
ROT_ITEM_ESQ = {n: [n] for n in ["GLOSA", "MANUAL", "COD.ITEM", "TP.ITEM.VERSAO", "DESCRICAO"]}
ROT_ITEM_DIR = {n: [n] for n in ["ERRO", "PRESTADOR", "NIVEL", "DT.REALZ", "HR.REALZ",
                                 "JUSTIFICATIVA", "AUDITOR"]}
TEXTO_LIVRE = {"TP.ITEM.VERSAO", "DESCRICAO", "PRESTADOR", "NIVEL", "HR.REALZ",
               "JUSTIFICATIVA", "AUDITOR", "SOLICITANTE", "NOME", "CARTEIRA"}
LARGURA_JUSTIFICATIVA = 51


def _conv(nome, texto):
    """Converte o texto de um campo; campos de texto livre ficam como texto."""
    if nome in TEXTO_LIVRE:
        return texto or None
    return valor(texto)


def parse_detalhe(linhas):
    saida = []
    guia, intern = {}, {}
    mapa_guia = mapa_int = mapa_esq = mapa_dir = None
    ini_num = fim_num = 0
    espera = None          # "guia" | "intern" -> próxima linha é de dados
    ajustado = False

    for l in linhas:
        s = l.strip()
        if not s or set(s) <= {"=", "-"}:
            continue

        if s.startswith("ANO") and "AUDITORIA" in l:
            mapa_guia = mapear(l, ROT_GUIA)
            espera, ajustado, intern = "guia", False, {}
            continue
        if s.startswith("DT.INTERNACAO"):
            mapa_int = mapear(l, ROT_INTERN)
            espera = "intern"
            continue
        if s.startswith("GLOSA") and "COD.ITEM" in l:
            mapa_esq = mapear(l, ROT_ITEM_ESQ)
            mapa_dir = mapear(l, ROT_ITEM_DIR)
            nomes_dir = [n for n, _ in mapa_dir]
            # layout antigo com cabeçalho cortado: AUDITOR não aparece, mas vem
            # logo depois da JUSTIFICATIVA (que tem largura fixa)
            if "AUDITOR" not in nomes_dir and nomes_dir and nomes_dir[-1] == "JUSTIFICATIVA":
                mapa_dir.append(("AUDITOR", mapa_dir[-1][1] + LARGURA_JUSTIFICATIVA))
            pos_desc = dict(mapa_esq)["DESCRICAO"]
            largura_desc = l.index("QT.COBR") - pos_desc - 5
            ini_num = pos_desc + largura_desc                 # números começam após a descrição
            fim_num = min(p for _, p in mapa_dir)             # e terminam na 1ª coluna da direita
            espera = None
            continue
        if "Itens Ajustados" in l:
            ajustado = True
            continue

        if espera == "guia":
            guia = {n: _conv(n, t) for n, t in fatiar(l, mapa_guia).items()}
            espera = None
            continue
        if espera == "intern":
            intern = {n: _conv(n, t) for n, t in fatiar(l, mapa_int).items()}
            espera = None
            continue

        # Linha de item: tem COD.ITEM preenchido e 4 valores numéricos
        if mapa_esq is None:
            continue
        esq = fatiar(l[:ini_num], mapa_esq)
        if not esq.get("COD.ITEM"):
            continue
        nums = RE_NUM_BR.findall(l[ini_num:fim_num])
        if len(nums) != 4:
            continue
        dir_ = fatiar(l, mapa_dir)

        campos = {**guia, **intern}
        campos.update({n: _conv(n, t) for n, t in esq.items()})
        campos.update(zip(["QT.COBR", "QT.PAGA", "VL.COBR", "VL.PAGO"], map(num_br, nums)))
        campos.update({n: _conv(n, t) for n, t in dir_.items()})
        campos["SITUAÇÃO"] = "ITEM AJUSTADO" if ajustado else "GLOSADO"
        saida.append([campos.get(n) for n in CAB_DETALHE])
    return saida


# ------------------- ITENS COM VALOR DIFERENTE ----------------------------- #
ROT_VALOR_DIF = {"GUIA SGU": ["GUIA SGU"], "ID GUIA": ["ID GUIA"], "COD.ITEM": ["COD.ITEM"],
                 "DESC.ITEM": ["DESC.ITEM"], "VL.INFOR": ["VL.INFOR", "VL.INFORMADO"],
                 "VL.PAGO": ["VL.PAGO"], "DT.EXECUÇÃO": ["DT.EXECUÇÃO", "DT.EXECUCAO"]}


def parse_valor_dif(linhas):
    """Devolve (cabeçalho, linhas). As colunas seguem as que existem no relatório."""
    mapa, saida = None, []
    for l in linhas:
        if mapa is None and "COD.ITEM" in l and "VL.INFOR" in l:
            mapa = mapear(l, ROT_VALOR_DIF)
            continue
        if mapa is None or not l.strip() or set(l.strip()) <= {"-", "="}:
            continue
        c = fatiar(l, mapa)
        if not c.get("COD.ITEM"):
            continue
        saida.append([num_br(c[n]) if n in ("VL.INFOR", "VL.PAGO")
                      else (c[n] or None) if n == "DESC.ITEM"
                      else valor(c[n]) for n, _ in mapa])
    cabecalho = [n for n, _ in mapa] if mapa else list(ROT_VALOR_DIF)[2:]
    return cabecalho, saida


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


class RelatorioInvalido(ValueError):
    """O conteúdo não parece ser um Relatório de Conferências de Glosas."""


def gerar_planilha(dados):
    """
    Núcleo da conversão: recebe o conteúdo do TXT (bytes) e devolve
    (Workbook do openpyxl, dicionário com a quantidade de linhas por aba).
    Não lê nem grava arquivos, por isso serve para desktop, web (Flask) ou linha de comando.
    """
    linhas = ler_linhas(dados)
    sec = localizar_secoes(linhas)

    resumo = parse_resumo(sec.get("resumo", []))
    detalhe = parse_detalhe(sec.get("detalhe", []))
    cab_valor_dif, valor_dif = parse_valor_dif(sec.get("valor_dif", []))
    mapa = parse_mapa(sec.get("mapa", []))
    glosa = parse_glosa(sec.get("glosa", []))

    if not resumo and not detalhe:
        raise RelatorioInvalido("O arquivo não parece ser um Relatório de Conferências de Glosas.")

    wb = Workbook()
    wb.remove(wb.active)
    escrever_aba(wb, "RESUMO", CAB_RESUMO, resumo)
    escrever_aba(wb, "DETALHAMENTO", CAB_DETALHE, detalhe)
    escrever_aba(wb, "ITENS COM VALOR DIFERENTE", cab_valor_dif, valor_dif)
    escrever_aba(wb, "MAPA DE SELEÇÃO", None, mapa, largura_max=55)
    escrever_aba(wb, "Glosa", None, glosa, largura_max=70)

    totais = {
        "RESUMO": len(resumo),
        "DETALHAMENTO": len(detalhe),
        "ITENS COM VALOR DIFERENTE": len(valor_dif),
        "MAPA DE SELEÇÃO": len(mapa),
        "Glosa": len(glosa),
    }
    return wb, totais


def converter(entrada, saida):
    """Arquivo TXT -> arquivo XLSX. Devolve as quantidades por aba."""
    wb, totais = gerar_planilha(Path(entrada).read_bytes())
    wb.save(saida)
    return totais


def converter_bytes(dados):
    """Conteúdo TXT (bytes) -> conteúdo XLSX (bytes). Útil para upload/download na web."""
    wb, totais = gerar_planilha(dados)
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), totais
