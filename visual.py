"""
Visual da janela: tema Sun Valley (sv_ttk, visual do Windows 11 pros
botões/campos/interruptores) + a paleta do programa.

Os "cartões" são tk.Frame com a mesma cor de fundo do tema (#FAFAFA) e borda
fina, sobre uma página um pouco mais escura - assim os widgets do tema
(que pintam o próprio fundo com #FAFAFA) se encaixam sem emenda.
"""

import tkinter as tk
from tkinter import ttk

PAGINA = "#ECEFF3"       # fundo da janela
CARTAO = "#FAFAFA"       # = fundo dos widgets do tema
BORDA = "#D9DEE5"
BORDA_FORTE = "#B7C0CC"
TEXTO = "#1C2430"
TEXTO_2 = "#5F6B7A"
CABECALHO = "#1E2A38"
CABECALHO_2 = "#9AA8B8"
MARCA = "#E87722"         # laranja de destaque (quadros, faixa do cabeçalho)
AZUL = "#005FB8"          # destaque do tema
VERDE = "#1E8E3E"
VERDE_FUNDO = "#E3F2E7"
LARANJA_FUNDO = "#FDEBDD"
CINZA = "#8A94A3"
CINZA_FUNDO = "#E8EBEF"
VERMELHO = "#C5221F"
VERMELHO_FUNDO = "#FBE4E3"

FONTE = "Segoe UI"


def aplicar_tema(raiz):
    """Liga o Sun Valley (se disponível) e cria os estilos do programa."""
    try:
        import sv_ttk

        sv_ttk.set_theme("light")
    except Exception:  # sem o tema, cai no visual padrão do Windows
        st = ttk.Style(raiz)
        if "vista" in st.theme_names():
            st.theme_use("vista")
    raiz.configure(bg=PAGINA)
    st = ttk.Style(raiz)
    st.configure("Legenda.TLabel", font=(FONTE, 8, "bold"), foreground=TEXTO_2)
    st.configure("Titulo.TLabel", font=(FONTE, 13, "bold"), foreground=TEXTO)
    st.configure("Sub.TLabel", font=(FONTE, 9), foreground=TEXTO_2)
    st.configure("Status.TLabel", font=(FONTE, 10), foreground=TEXTO)
    st.configure("Proposta.TEntry", padding=(10, 6))
    st.configure("Grande.Accent.TButton", font=(FONTE, 11, "bold"), padding=(18, 8))
    st.configure("Grande.TButton", padding=(12, 6))
    return st


def cartao(pai, **kw):
    """Moldura clara com borda fina (o "cartão" da tela)."""
    return tk.Frame(pai, bg=CARTAO, highlightthickness=1, highlightbackground=BORDA,
                    highlightcolor=BORDA, **kw)


def chip(pai, texto="", fundo=CINZA_FUNDO, cor=TEXTO_2, bg=CARTAO):
    """Etiqueta arredondada-ish (contadores)."""
    return tk.Label(pai, text=texto, bg=fundo, fg=cor, font=(FONTE, 9, "bold"), padx=10, pady=3)


class Etapas(tk.Frame):
    """Indicador ① Buscar → ② Escolher → ③ Gerar."""

    NOMES = ("Buscar a proposta", "Escolher e ordenar as fotos", "Gerar o Excel")

    def __init__(self, pai):
        super().__init__(pai, bg=PAGINA)
        self.itens = []
        for n, nome in enumerate(self.NOMES, 1):
            if n > 1:
                tk.Frame(self, bg=BORDA_FORTE, height=2, width=46).pack(side="left", padx=10)
            num = tk.Label(self, text=str(n), width=2, font=(FONTE, 9, "bold"), fg="white", bg=CINZA)
            num.pack(side="left")
            txt = tk.Label(self, text=nome, font=(FONTE, 10), bg=PAGINA, fg=TEXTO_2)
            txt.pack(side="left", padx=(6, 0))
            self.itens.append((num, txt))
        self.marcar(1)

    def marcar(self, atual, tudo_feito=False):
        for n, (num, txt) in enumerate(self.itens, 1):
            if n < atual or tudo_feito:
                num.config(text="✓", bg=VERDE)
                txt.config(fg=VERDE, font=(FONTE, 10))
            elif n == atual:
                num.config(text=str(n), bg=AZUL)
                txt.config(fg=TEXTO, font=(FONTE, 10, "bold"))
            else:
                num.config(text=str(n), bg=CINZA)
                txt.config(fg=TEXTO_2, font=(FONTE, 10))
