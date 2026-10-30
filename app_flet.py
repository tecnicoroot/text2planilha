#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor de Glosas - interface desktop em Flet (Flet 1.0).
A lógica de conversão fica no pacote conversor_glosas.

Executar:  python app_flet.py
"""
import asyncio
import os
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

import flet as ft

from conversor_glosas import RelatorioInvalido, converter

APP_NOME = "Conversor de Glosas"
APP_VERSAO = "2.0"

# --------------------------------------------------------------------------- #
# Identidade visual
# --------------------------------------------------------------------------- #
VERDE = "#0F6E56"        # cor principal (faixa lateral, botão principal)
VERDE_ESCURO = "#0A4F3E"
VERDE_CLARO = "#E3F1EC"
TINTA = "#16302A"        # texto principal
TINTA_SUAVE = "#5B6F69"  # texto secundário
FUNDO = "#F5F7F6"
LINHA = "#DCE5E1"
BRANCO = "#FFFFFF"
OK = "#2E9E6B"
ERRO = "#B9402A"
FONTE = "Segoe UI"

# Status possíveis de um arquivo na fila
PENDENTE, CONVERTENDO, CONCLUIDO, FALHOU = "pendente", "convertendo", "concluido", "falhou"


@dataclass
class Item:
    """Um TXT na fila de conversão."""
    entrada: Path
    status: str = PENDENTE
    saida: Path | None = None
    mensagem: str = ""
    totais: dict = field(default_factory=dict)


def abrir_no_sistema(caminho):
    """Abre arquivo ou pasta com o programa padrão do sistema."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(caminho)])
        else:
            subprocess.Popen(["xdg-open", str(caminho)])
    except Exception:
        pass


def tamanho_legivel(n):
    for unidade in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {unidade}"
        n /= 1024
    return f"{n:.1f} TB"


# --------------------------------------------------------------------------- #
# Aplicação
# --------------------------------------------------------------------------- #
class ConversorApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.itens: list[Item] = []
        self.pasta_destino: Path | None = None
        self.convertendo = False

        self._configurar_pagina()
        self.picker = ft.FilePicker()          # serviço do Flet: registra-se sozinho na página
        self._montar()
        self.atualizar_fila()

    # ------------------------------------------------------------------ #
    # Página e layout
    # ------------------------------------------------------------------ #
    def _configurar_pagina(self):
        p = self.page
        p.title = f"{APP_NOME} {APP_VERSAO}"
        p.theme_mode = ft.ThemeMode.LIGHT
        p.theme = ft.Theme(color_scheme_seed=VERDE, font_family=FONTE)
        p.bgcolor = FUNDO
        p.padding = 0
        p.window.width = 980
        p.window.height = 680
        p.window.min_width = 820
        p.window.min_height = 560

    def _montar(self):
        self.page.add(
            ft.Row(
                [self._painel_lateral(), self._area_principal()],
                spacing=0,
                expand=True,
                vertical_alignment=ft.CrossAxisAlignment.STRETCH,
            )
        )

    def _painel_lateral(self):
        def passo(numero, titulo, texto):
            return ft.Row(
                [
                    ft.Container(
                        ft.Text(str(numero), color=VERDE, weight=ft.FontWeight.BOLD, size=13),
                        width=26, height=26, border_radius=13, bgcolor=BRANCO,
                        alignment=ft.Alignment.CENTER,
                    ),
                    ft.Column(
                        [
                            ft.Text(titulo, color=BRANCO, weight=ft.FontWeight.W_600, size=14),
                            ft.Text(texto, color="#CFE6DE", size=12.5),
                        ],
                        spacing=2, expand=True,
                    ),
                ],
                spacing=12,
                vertical_alignment=ft.CrossAxisAlignment.START,
            )

        self.sw_mesma_pasta = ft.Switch(
            value=True, active_color=BRANCO, active_track_color="#5FB394",
            on_change=self._ao_mudar_destino,
        )
        self.sw_abrir = ft.Switch(value=True, active_color=BRANCO, active_track_color="#5FB394")
        self.txt_pasta = ft.Text("", color="#CFE6DE", size=12, max_lines=2,
                                 overflow=ft.TextOverflow.ELLIPSIS, visible=False)
        self.btn_pasta = ft.TextButton(
            "Escolher pasta", icon=ft.Icons.FOLDER_OPEN_ROUNDED, visible=False,
            style=ft.ButtonStyle(color=BRANCO), on_click=self._escolher_pasta,
        )

        def opcao(texto, switch):
            return ft.Row(
                [ft.Text(texto, color=BRANCO, size=13, expand=True), switch],
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            )

        return ft.Container(
            width=290,
            bgcolor=VERDE,
            padding=ft.Padding.symmetric(horizontal=24, vertical=28),
            content=ft.Column(
                [
                    ft.Icon(ft.Icons.TABLE_VIEW_ROUNDED, color=BRANCO, size=34),
                    ft.Text(APP_NOME, color=BRANCO, size=22, weight=ft.FontWeight.BOLD),
                    ft.Text("Relatório de conferência de glosas em TXT para planilha Excel.",
                            color="#CFE6DE", size=13),
                    ft.Container(height=18),
                    passo(1, "Selecione os TXT", "Um ou vários relatórios de uma vez."),
                    passo(2, "Escolha o destino", "Mesma pasta do TXT ou outra pasta."),
                    passo(3, "Converta", "Cada TXT vira uma planilha com 5 abas."),
                    ft.Container(expand=True),
                    ft.Divider(color="#2C8A70", height=1),
                    opcao("Salvar na pasta do TXT", self.sw_mesma_pasta),
                    self.txt_pasta,
                    self.btn_pasta,
                    opcao("Abrir ao terminar", self.sw_abrir),
                    ft.Text(f"Versão {APP_VERSAO}", color="#8CC4B1", size=11),
                ],
                spacing=10,
                expand=True,
            ),
        )

    def _area_principal(self):
        self.lista = ft.ListView(spacing=8, expand=True, padding=ft.Padding.only(right=8))
        self.vazio = self._estado_vazio()

        self.lbl_fila = ft.Text("", color=TINTA, size=20, weight=ft.FontWeight.W_600)
        self.lbl_resumo = ft.Text("", color=TINTA_SUAVE, size=13)
        self.btn_adicionar = ft.OutlinedButton(
            "Adicionar TXT", icon=ft.Icons.ADD_ROUNDED, on_click=self._adicionar,
            style=ft.ButtonStyle(color=VERDE, side=ft.BorderSide(1, LINHA),
                                 shape=ft.RoundedRectangleBorder(radius=8)),
        )
        self.btn_limpar = ft.TextButton(
            "Limpar fila", on_click=self._limpar, style=ft.ButtonStyle(color=TINTA_SUAVE),
        )

        self.progresso = ft.ProgressBar(value=0, color=VERDE, bgcolor=LINHA,
                                        bar_height=6, border_radius=3, expand=True)
        self.lbl_progresso = ft.Text("", color=TINTA_SUAVE, size=12.5)
        self.btn_converter = ft.FilledButton(
            "Converter para Excel", icon=ft.Icons.PLAY_ARROW_ROUNDED,
            on_click=self._converter,
            style=ft.ButtonStyle(bgcolor={ft.ControlState.DEFAULT: VERDE,
                                          ft.ControlState.HOVERED: VERDE_ESCURO,
                                          ft.ControlState.DISABLED: "#A9C2BA"},
                                 color=BRANCO,
                                 padding=ft.Padding.symmetric(horizontal=22, vertical=18),
                                 shape=ft.RoundedRectangleBorder(radius=8)),
        )

        cabecalho = ft.Row(
            [
                ft.Column([self.lbl_fila, self.lbl_resumo], spacing=2, expand=True),
                self.btn_limpar,
                self.btn_adicionar,
            ],
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        rodape = ft.Container(
            bgcolor=BRANCO,
            border=ft.Border.all(1, LINHA),
            border_radius=10,
            padding=ft.Padding.symmetric(horizontal=18, vertical=14),
            content=ft.Row(
                [
                    ft.Column([self.lbl_progresso, ft.Row([self.progresso])],
                              spacing=6, expand=True),
                    self.btn_converter,
                ],
                spacing=20,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
        )
        self.corpo = ft.Stack([self.lista, self.vazio], expand=True)

        return ft.Container(
            expand=True,
            padding=ft.Padding.symmetric(horizontal=28, vertical=24),
            content=ft.Column([cabecalho, self.corpo, rodape], spacing=16, expand=True),
        )

    def _estado_vazio(self):
        return ft.Container(
            expand=True,
            border=ft.Border.all(1.5, LINHA),
            border_radius=12,
            bgcolor=BRANCO,
            alignment=ft.Alignment.CENTER,
            on_click=self._adicionar,
            ink=True,
            content=ft.Column(
                [
                    ft.Icon(ft.Icons.UPLOAD_FILE_ROUNDED, size=48, color=VERDE),
                    ft.Text("Nenhum relatório na fila", size=16, weight=ft.FontWeight.W_600,
                            color=TINTA),
                    ft.Text("Clique aqui ou em “Adicionar TXT” para escolher os arquivos.",
                            size=13, color=TINTA_SUAVE),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=6,
                tight=True,
            ),
        )

    # ------------------------------------------------------------------ #
    # Fila de arquivos
    # ------------------------------------------------------------------ #
    def _linha_item(self, item: Item):
        cores = {
            PENDENTE: (TINTA_SUAVE, "#EEF2F0", "Na fila"),
            CONVERTENDO: (VERDE, VERDE_CLARO, "Convertendo…"),
            CONCLUIDO: (OK, "#E2F4EB", "Pronto"),
            FALHOU: (ERRO, "#FBE9E5", "Não convertido"),
        }
        cor, fundo, rotulo = cores[item.status]

        if item.status == CONVERTENDO:
            icone = ft.ProgressRing(width=22, height=22, stroke_width=2.5, color=VERDE)
        else:
            icone = ft.Icon(
                {PENDENTE: ft.Icons.DESCRIPTION_OUTLINED,
                 CONCLUIDO: ft.Icons.CHECK_CIRCLE_ROUNDED,
                 FALHOU: ft.Icons.ERROR_OUTLINE_ROUNDED}[item.status],
                color=cor, size=24,
            )

        try:
            detalhe = f"{tamanho_legivel(item.entrada.stat().st_size)}  ·  {item.entrada.parent}"
        except OSError:
            detalhe = str(item.entrada.parent)

        textos = [
            ft.Text(item.entrada.name, size=14, weight=ft.FontWeight.W_600, color=TINTA,
                    max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
            ft.Text(detalhe, size=12, color=TINTA_SUAVE, max_lines=1,
                    overflow=ft.TextOverflow.ELLIPSIS),
        ]
        if item.status == CONCLUIDO and item.totais:
            textos.append(ft.Row(
                [self._chip(f"{aba}: {qtd:,}".replace(",", ".")) for aba, qtd in item.totais.items()],
                spacing=6, wrap=True, run_spacing=6,
            ))
        if item.status == FALHOU:
            textos.append(ft.Text(item.mensagem, size=12.5, color=ERRO))

        acoes = []
        if item.status == CONCLUIDO:
            acoes.append(ft.IconButton(ft.Icons.OPEN_IN_NEW_ROUNDED, icon_color=VERDE,
                                       tooltip="Abrir planilha",
                                       on_click=lambda e, i=item: abrir_no_sistema(i.saida)))
            acoes.append(ft.IconButton(ft.Icons.FOLDER_OUTLINED, icon_color=TINTA_SUAVE,
                                       tooltip="Abrir pasta",
                                       on_click=lambda e, i=item: abrir_no_sistema(i.saida.parent)))
        if not self.convertendo:
            acoes.append(ft.IconButton(ft.Icons.CLOSE_ROUNDED, icon_color=TINTA_SUAVE,
                                       tooltip="Tirar da fila",
                                       on_click=lambda e, i=item: self._remover(i)))

        return ft.Container(
            bgcolor=BRANCO,
            border=ft.Border.all(1, LINHA),
            border_radius=10,
            padding=ft.Padding.symmetric(horizontal=16, vertical=12),
            content=ft.Row(
                [
                    icone,
                    ft.Column(textos, spacing=4, expand=True),
                    ft.Container(
                        ft.Text(rotulo, size=12, color=cor, weight=ft.FontWeight.W_600),
                        bgcolor=fundo, border_radius=20,
                        padding=ft.Padding.symmetric(horizontal=10, vertical=4),
                    ),
                    *acoes,
                ],
                spacing=14,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
        )

    @staticmethod
    def _chip(texto):
        return ft.Container(
            ft.Text(texto, size=11.5, color=VERDE_ESCURO),
            bgcolor=VERDE_CLARO, border_radius=6,
            padding=ft.Padding.symmetric(horizontal=8, vertical=3),
        )

    def atualizar_fila(self):
        self.lista.controls = [self._linha_item(i) for i in self.itens]
        tem_itens = bool(self.itens)
        self.vazio.visible = not tem_itens
        self.lista.visible = tem_itens

        n = len(self.itens)
        prontos = sum(i.status == CONCLUIDO for i in self.itens)
        falhas = sum(i.status == FALHOU for i in self.itens)
        self.lbl_fila.value = "Fila de conversão"
        if not n:
            self.lbl_resumo.value = "Os relatórios escolhidos aparecem aqui."
        else:
            partes = [f"{n} arquivo{'s' if n > 1 else ''}"]
            if prontos:
                partes.append(f"{prontos} pronto{'s' if prontos > 1 else ''}")
            if falhas:
                partes.append(f"{falhas} com erro")
            self.lbl_resumo.value = "  ·  ".join(partes)

        pendentes = any(i.status in (PENDENTE, FALHOU) for i in self.itens)
        self.btn_converter.disabled = self.convertendo or not pendentes
        self.btn_adicionar.disabled = self.convertendo
        self.btn_limpar.visible = tem_itens and not self.convertendo
        self.page.update()

    # ------------------------------------------------------------------ #
    # Ações
    # ------------------------------------------------------------------ #
    def adicionar_caminhos(self, caminhos):
        existentes = {i.entrada for i in self.itens}
        for c in caminhos:
            p = Path(c)
            if p not in existentes:
                self.itens.append(Item(p))
                existentes.add(p)
        self.atualizar_fila()

    async def _adicionar(self, e):
        if self.convertendo:
            return
        arquivos = await self.picker.pick_files(
            dialog_title="Selecione o(s) relatório(s) TXT",
            allowed_extensions=["txt", "TXT"],
            allow_multiple=True,
        )
        if arquivos:
            self.adicionar_caminhos([f.path for f in arquivos if f.path])

    def _remover(self, item):
        self.itens.remove(item)
        self.atualizar_fila()

    def _limpar(self, e):
        self.itens.clear()
        self.progresso.value = 0
        self.lbl_progresso.value = ""
        self.atualizar_fila()

    def _ao_mudar_destino(self, e):
        outra = not self.sw_mesma_pasta.value
        self.btn_pasta.visible = outra
        self.txt_pasta.visible = outra and self.pasta_destino is not None
        self.page.update()

    async def _escolher_pasta(self, e):
        pasta = await self.picker.get_directory_path(dialog_title="Pasta onde salvar as planilhas")
        if pasta:
            self.pasta_destino = Path(pasta)
            self.txt_pasta.value = pasta
            self.txt_pasta.visible = True
            self.page.update()

    def _destino(self, entrada: Path) -> Path:
        pasta = entrada.parent if self.sw_mesma_pasta.value else self.pasta_destino
        return pasta / (entrada.stem + ".xlsx")

    def _aviso(self, texto, cor=TINTA):
        self.page.show_dialog(ft.SnackBar(ft.Text(texto, color=BRANCO), bgcolor=cor,
                                          behavior=ft.SnackBarBehavior.FLOATING, width=460))

    async def _converter(self, e):
        if not self.sw_mesma_pasta.value and self.pasta_destino is None:
            self._aviso("Escolha a pasta de destino no painel ao lado.", ERRO)
            return

        fila = [i for i in self.itens if i.status in (PENDENTE, FALHOU)]
        if not fila:
            return

        self.convertendo = True
        self.progresso.value = 0
        for n, item in enumerate(fila, 1):
            item.status = CONVERTENDO
            item.mensagem = ""
            self.lbl_progresso.value = f"Convertendo {n} de {len(fila)}: {item.entrada.name}"
            self.atualizar_fila()

            item.saida = self._destino(item.entrada)
            try:
                # roda fora da tela para a janela não congelar
                item.totais = await asyncio.to_thread(converter, str(item.entrada), str(item.saida))
                item.status = CONCLUIDO
            except RelatorioInvalido as erro:
                item.status, item.mensagem = FALHOU, str(erro)
            except PermissionError:
                item.status = FALHOU
                item.mensagem = (f"Não foi possível salvar {item.saida.name}. "
                                 "Se ele estiver aberto no Excel, feche e converta de novo.")
            except FileNotFoundError:
                item.status, item.mensagem = FALHOU, "Arquivo ou pasta não encontrado."
            except Exception as erro:
                item.status, item.mensagem = FALHOU, f"Erro inesperado: {erro}"

            self.progresso.value = n / len(fila)

        self.convertendo = False
        ok = [i for i in fila if i.status == CONCLUIDO]
        falhas = len(fila) - len(ok)
        self.lbl_progresso.value = (f"Concluído: {len(ok)} convertido{'s' if len(ok) != 1 else ''}"
                                    + (f", {falhas} com erro." if falhas else "."))
        self.atualizar_fila()

        if falhas:
            self._aviso(f"{falhas} arquivo(s) não foram convertidos. Veja o motivo na fila.", ERRO)
        else:
            self._aviso("Planilhas geradas com sucesso.", OK)

        if ok and self.sw_abrir.value:
            abrir_no_sistema(ok[0].saida if len(ok) == 1 else ok[0].saida.parent)


def main(page: ft.Page):
    ConversorApp(page)


if __name__ == "__main__":
    ft.run(main)
