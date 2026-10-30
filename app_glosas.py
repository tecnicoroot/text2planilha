#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor de Glosas - interface gráfica.
Converte um ou vários TXT do "Relatório de Conferências de Glosas" em XLSX.
"""
import os
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from conferencia_glosas_xlsx import converter

APP_NOME = "Conversor de Glosas"
APP_VERSAO = "1.0"

COR_PRIMARIA = "#1F4E78"
COR_FUNDO = "#F4F6F9"
COR_TEXTO = "#1E1E1E"
COR_OK = "#2E7D32"
COR_ERRO = "#C62828"


def abrir_no_sistema(caminho):
    """Abre arquivo/pasta com o programa padrão do sistema."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", caminho])
        else:
            subprocess.Popen(["xdg-open", caminho])
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NOME} v{APP_VERSAO}")
        self.geometry("780x620")
        self.minsize(640, 480)
        self.configure(bg=COR_FUNDO)

        self.arquivos = []
        self.pasta_saida = tk.StringVar(value="")
        self.mesma_pasta = tk.BooleanVar(value=True)
        self.abrir_ao_final = tk.BooleanVar(value=True)
        self.ultimo_gerado = None

        self._estilos()
        self._montar_tela()

    # ------------------------------------------------------------------ #
    def _estilos(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure(".", background=COR_FUNDO, foreground=COR_TEXTO,
                     font=("Segoe UI", 10))
        st.configure("TFrame", background=COR_FUNDO)
        st.configure("Card.TFrame", background="white", relief="solid", borderwidth=1)
        st.configure("Card.TLabel", background="white")
        st.configure("Card.TCheckbutton", background="white")
        st.configure("Titulo.TLabel", background=COR_PRIMARIA, foreground="white",
                     font=("Segoe UI", 16, "bold"))
        st.configure("Sub.TLabel", background=COR_PRIMARIA, foreground="#D6E4F0",
                     font=("Segoe UI", 10))
        st.configure("Secao.TLabel", background="white", font=("Segoe UI", 10, "bold"))
        st.configure("Primario.TButton", background=COR_PRIMARIA, foreground="white",
                     font=("Segoe UI", 11, "bold"), padding=(18, 8))
        st.map("Primario.TButton",
               background=[("disabled", "#9AA9B8"), ("active", "#2C6AA0")])
        st.configure("TButton", padding=(10, 5))
        st.configure("Horizontal.TProgressbar", background=COR_PRIMARIA)

    def _montar_tela(self):
        # Cabeçalho
        topo = tk.Frame(self, bg=COR_PRIMARIA, height=70)
        topo.pack(fill="x")
        ttk.Label(topo, text=APP_NOME, style="Titulo.TLabel").pack(anchor="w", padx=20, pady=(12, 0))
        ttk.Label(topo, text="Transforma o relatório TXT de conferência de glosas em planilha Excel",
                  style="Sub.TLabel").pack(anchor="w", padx=20, pady=(0, 12))

        corpo = ttk.Frame(self, padding=16)
        corpo.pack(fill="both", expand=True)

        # 1. Arquivos
        card1 = ttk.Frame(corpo, style="Card.TFrame", padding=12)
        card1.pack(fill="both", expand=True)
        ttk.Label(card1, text="1. Arquivos TXT", style="Secao.TLabel").grid(row=0, column=0, sticky="w")

        self.lista = tk.Listbox(card1, height=6, activestyle="none", selectmode="extended",
                                font=("Segoe UI", 9), relief="flat", bg="#FAFBFC",
                                highlightthickness=1, highlightcolor=COR_PRIMARIA,
                                highlightbackground="#D0D7DE")
        self.lista.grid(row=1, column=0, rowspan=3, sticky="nsew", pady=(6, 0))
        ttk.Button(card1, text="Adicionar…", command=self.adicionar).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=(6, 2))
        ttk.Button(card1, text="Remover", command=self.remover).grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=2)
        ttk.Button(card1, text="Limpar", command=self.limpar).grid(row=3, column=1, sticky="new", padx=(10, 0), pady=2)
        card1.columnconfigure(0, weight=1)
        card1.rowconfigure(3, weight=1)

        # 2. Saída
        card2 = ttk.Frame(corpo, style="Card.TFrame", padding=12)
        card2.pack(fill="x", pady=(12, 0))
        ttk.Label(card2, text="2. Onde salvar", style="Secao.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(card2, text="Na mesma pasta do TXT", variable=self.mesma_pasta,
                        style="Card.TCheckbutton", command=self._alternar_pasta
                        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 4))
        self.ent_pasta = ttk.Entry(card2, textvariable=self.pasta_saida, state="disabled")
        self.ent_pasta.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.btn_pasta = ttk.Button(card2, text="Escolher pasta…", command=self.escolher_pasta, state="disabled")
        self.btn_pasta.grid(row=2, column=2, padx=(10, 0))
        ttk.Checkbutton(card2, text="Abrir a planilha ao terminar", variable=self.abrir_ao_final,
                        style="Card.TCheckbutton").grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        card2.columnconfigure(0, weight=1)

        # 3. Ação
        acao = ttk.Frame(corpo)
        acao.pack(fill="x", pady=(14, 0))
        self.btn_converter = ttk.Button(acao, text="Converter para Excel", style="Primario.TButton",
                                        command=self.iniciar)
        self.btn_converter.pack(side="left")
        self.progresso = ttk.Progressbar(acao, mode="determinate")
        self.progresso.pack(side="left", fill="x", expand=True, padx=(14, 0))

        # Status
        self.status = tk.Text(corpo, height=6, relief="flat", bg="white", font=("Consolas", 9),
                              highlightthickness=1, highlightbackground="#D0D7DE", state="disabled")
        self.status.pack(fill="both", expand=False, pady=(12, 0))
        self.status.tag_configure("ok", foreground=COR_OK)
        self.status.tag_configure("erro", foreground=COR_ERRO)
        self._log("Adicione um ou mais arquivos TXT e clique em “Converter para Excel”.")

    # ------------------------------------------------------------------ #
    def _log(self, msg, tag=None):
        self.status.configure(state="normal")
        self.status.insert("end", msg + "\n", tag)
        self.status.see("end")
        self.status.configure(state="disabled")

    def _alternar_pasta(self):
        estado = "disabled" if self.mesma_pasta.get() else "normal"
        self.ent_pasta.configure(state=estado)
        self.btn_pasta.configure(state=estado)

    def adicionar(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecione o(s) relatório(s) TXT",
            filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")])
        for c in caminhos:
            if c not in self.arquivos:
                self.arquivos.append(c)
                self.lista.insert("end", c)

    def remover(self):
        for i in reversed(self.lista.curselection()):
            self.lista.delete(i)
            del self.arquivos[i]

    def limpar(self):
        self.lista.delete(0, "end")
        self.arquivos.clear()

    def escolher_pasta(self):
        p = filedialog.askdirectory(title="Pasta onde salvar as planilhas")
        if p:
            self.pasta_saida.set(p)

    # ------------------------------------------------------------------ #
    def iniciar(self):
        if not self.arquivos:
            messagebox.showwarning(APP_NOME, "Adicione pelo menos um arquivo TXT.")
            return
        if not self.mesma_pasta.get() and not self.pasta_saida.get():
            messagebox.showwarning(APP_NOME, "Escolha a pasta onde salvar as planilhas.")
            return
        self.btn_converter.configure(state="disabled")
        self.progresso.configure(value=0, maximum=len(self.arquivos))
        threading.Thread(target=self._processar, args=(list(self.arquivos),), daemon=True).start()

    def _destino(self, entrada):
        pasta = Path(entrada).parent if self.mesma_pasta.get() else Path(self.pasta_saida.get())
        return pasta / (Path(entrada).stem + ".xlsx")

    def _processar(self, arquivos):
        ok = erros = 0
        for n, entrada in enumerate(arquivos, 1):
            saida = self._destino(entrada)
            self.after(0, self._log, f"Convertendo {Path(entrada).name}…")
            try:
                totais = converter(entrada, str(saida))
                if not isinstance(totais, dict):
                    # versão antiga do conferencia_glosas_xlsx.py não devolve contagens
                    totais = {}
                if totais and not totais.get("RESUMO") and not totais.get("DETALHAMENTO"):
                    raise ValueError("O arquivo não parece ser um Relatório de Conferências de Glosas.")
                resumo = ", ".join(f"{k}: {v}" for k, v in totais.items()) or "contagens indisponíveis"
                self.after(0, self._log, f"  ✔ {saida.name}  ({resumo})", "ok")
                self.ultimo_gerado = str(saida)
                ok += 1
            except PermissionError:
                self.after(0, self._log, f"  ✖ Não consegui salvar {saida.name}. "
                                         "Ele está aberto no Excel? Feche e tente de novo.", "erro")
                erros += 1
            except Exception as e:
                self.after(0, self._log, f"  ✖ Erro: {e}", "erro")
                traceback.print_exc()
                erros += 1
            self.after(0, self.progresso.configure, {"value": n})
        self.after(0, self._finalizar, ok, erros)

    def _finalizar(self, ok, erros):
        self.btn_converter.configure(state="normal")
        self._log(f"Concluído: {ok} convertido(s), {erros} com erro.", "ok" if not erros else "erro")
        if ok and self.abrir_ao_final.get():
            if ok == 1 and self.ultimo_gerado:
                abrir_no_sistema(self.ultimo_gerado)
            else:
                abrir_no_sistema(str(Path(self.ultimo_gerado).parent))
        if erros:
            messagebox.showwarning(APP_NOME, f"{erros} arquivo(s) não foram convertidos. Veja os detalhes na tela.")


if __name__ == "__main__":
    App().mainloop()