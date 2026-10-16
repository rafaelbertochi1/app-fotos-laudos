"""
Janela do Downloader de Fotos (vira o .exe).

Duas etapas: "Buscar fotos" (baixa tudo pra um cache e mostra a galeria -
galeria.py - onde a pessoa escolhe, ordena e manda até 4 fotos pra aba
"Mapa e fotos locais") e "Gerar Excel". O trabalho pesado roda numa thread
separada e conversa com a janela por uma fila (Tkinter não aceita ser
mexido de fora da thread principal).
"""

import json
import os
import queue
import shutil
import threading
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import comum
import pasta_publica
import robo
import trabalho
import intranet
from perfis import PADRAO, PERFIS
from central_gestao import ErroCentralGestao, ErroLogin, apagar_token, obter_token
from galeria import Galeria, gerar_miniaturas
import visual

TITULO = "Downloader de Fotos - Central de Gestão"


def carregar_config():
    try:
        with open(comum.CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def salvar_config(cfg):
    try:
        with open(comum.CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _mensagem_erro(e):
    """Traduz os erros mais comuns do Excel/arquivos pra algo que o usuário entenda."""
    if isinstance(e, PermissionError):
        return ("Não consegui gravar o arquivo. Se o Excel desta proposta estiver aberto, "
                f"feche-o e tente de novo.\n\n({e.filename})")
    if "pywintypes" in type(e).__module__ or type(e).__name__ == "com_error":
        return ("O Excel deu erro ao montar a planilha. Feche outras janelas do Excel "
                f"e tente de novo.\n\n({str(e)[:300]})")
    return str(e) or type(e).__name__


class JanelaLogin(tk.Toplevel):
    """Grava e-mail/senha da Central de Gestão neste PC (comum.gravar_login)."""

    def __init__(self, pai, ao_salvar=None):
        super().__init__(pai)
        self.title("Configurar login da Central de Gestão")
        self.resizable(False, False)
        self.transient(pai)
        self.grab_set()
        self.ao_salvar = ao_salvar

        corpo = ttk.Frame(self, padding=16)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text="Login da Central de Gestão", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(corpo, text="Ficam guardados só neste computador, criptografados pelo\n"
                              "Windows. Você só precisa fazer isso uma vez.",
                  foreground="#555").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 12))

        ttk.Label(corpo, text="E-mail:").grid(row=2, column=0, sticky="w")
        self.usuario = ttk.Entry(corpo, width=34)
        self.usuario.grid(row=2, column=1, pady=4)
        self.usuario.insert(0, comum.usuario_configurado() or "")
        ttk.Label(corpo, text="Senha:").grid(row=3, column=0, sticky="w")
        self.senha = ttk.Entry(corpo, width=34, show="•")
        self.senha.grid(row=3, column=1, pady=4)

        botoes = ttk.Frame(corpo)
        botoes.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(botoes, text="Esquecer login", command=self.esquecer).pack(side="left", padx=(0, 16))
        ttk.Button(botoes, text="Cancelar", command=self.destroy).pack(side="left", padx=4)
        ttk.Button(botoes, text="Salvar", style="Accent.TButton", command=self.salvar).pack(side="left")

        (self.senha if self.usuario.get() else self.usuario).focus_set()
        self.bind("<Return>", lambda e: self.salvar())
        self.bind("<Escape>", lambda e: self.destroy())

    def salvar(self):
        usuario, senha = self.usuario.get().strip(), self.senha.get()
        if not usuario or not senha:
            messagebox.showwarning(TITULO, "Preencha e-mail e senha.", parent=self)
            return
        try:
            comum.gravar_login(usuario, senha)
            comum.apagar_sessao()
            apagar_token()  # força entrar de novo com o login novo
        except Exception as e:
            messagebox.showerror(TITULO, f"Não consegui gravar o login neste computador:\n\n{e}\n\n"
                                         f"Pasta: {comum.PASTA_DADOS}", parent=self)
            return

        # já testa na Central de Gestão, pra senha errada aparecer agora e não na hora de rodar
        self.config(cursor="watch")
        self.update()
        try:
            obter_token(log=lambda m: None)
            messagebox.showinfo(TITULO, "Login salvo e testado na Central de Gestão com sucesso.", parent=self)
        except ErroLogin as e:
            self.config(cursor="")
            messagebox.showerror(TITULO, f"{e}\n\nConfira o e-mail e a senha.", parent=self)
            return
        except Exception as e:
            messagebox.showwarning(TITULO, f"Login salvo, mas não consegui testá-lo agora:\n\n{e}", parent=self)
        if self.ao_salvar:
            self.ao_salvar()
        self.destroy()

    def esquecer(self):
        if messagebox.askyesno(TITULO, "Apagar o usuário e a senha gravados neste computador?", parent=self):
            comum.apagar_login()
            apagar_token()
            if self.ao_salvar:
                self.ao_salvar()
            self.destroy()


class JanelaLoginIntranet(tk.Toplevel):
    """Grava usuário/senha da Intranet neste PC (intranet.gravar_login): é
    de lá que o programa descobre o tipo de operação do Banco D (o modelo)."""

    def __init__(self, pai, ao_salvar=None):
        super().__init__(pai)
        self.title("Login da Intranet")
        self.resizable(False, False)
        self.transient(pai)
        self.grab_set()
        self.ao_salvar = ao_salvar

        corpo = ttk.Frame(self, padding=16)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text="Login da Intranet", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(corpo, text="Usado só para ler, na requisição da proposta, o tipo de operação\n"
                              "do Banco D e já escolher o modelo do Excel. Fica guardado só\n"
                              "neste computador, criptografado pelo Windows.",
                  foreground="#555").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 12))
        ttk.Label(corpo, text="Usuário:").grid(row=2, column=0, sticky="w")
        self.usuario = ttk.Entry(corpo, width=34)
        self.usuario.grid(row=2, column=1, pady=4)
        self.usuario.insert(0, intranet.usuario_configurado() or "")
        ttk.Label(corpo, text="Senha:").grid(row=3, column=0, sticky="w")
        self.senha = ttk.Entry(corpo, width=34, show="•")
        self.senha.grid(row=3, column=1, pady=4)

        botoes = ttk.Frame(corpo)
        botoes.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(botoes, text="Esquecer login", command=self.esquecer).pack(side="left", padx=(0, 16))
        ttk.Button(botoes, text="Cancelar", command=self.destroy).pack(side="left", padx=4)
        ttk.Button(botoes, text="Salvar", style="Accent.TButton", command=self.salvar).pack(side="left")
        (self.senha if self.usuario.get() else self.usuario).focus_set()
        self.bind("<Return>", lambda e: self.salvar())
        self.bind("<Escape>", lambda e: self.destroy())

    def salvar(self):
        usuario, senha = self.usuario.get().strip(), self.senha.get()
        if not usuario or not senha:
            messagebox.showwarning(TITULO, "Preencha usuário e senha.", parent=self)
            return
        self.config(cursor="watch")
        self.update()
        try:
            intranet.testar_login(usuario, senha)  # testa antes de gravar: senha errada aparece agora
        except intranet.ErroLoginIntranet as e:
            self.config(cursor="")
            messagebox.showerror(TITULO, f"{e}.\n\nConfira o usuário e a senha da Intranet.", parent=self)
            return
        except Exception as e:
            self.config(cursor="")
            if not messagebox.askyesno(TITULO, f"Não consegui testar o login agora:\n\n{e}\n\nGravar assim mesmo?",
                                       parent=self):
                return
        try:
            intranet.gravar_login(usuario, senha)
        except Exception as e:
            self.config(cursor="")
            messagebox.showerror(TITULO, f"Não consegui gravar o login neste computador:\n\n{e}", parent=self)
            return
        messagebox.showinfo(TITULO, "Login da Intranet salvo e testado.", parent=self)
        if self.ao_salvar:
            self.ao_salvar()
        self.destroy()

    def esquecer(self):
        if messagebox.askyesno(TITULO, "Apagar o login da Intranet gravado neste computador?", parent=self):
            intranet.apagar_login()
            if self.ao_salvar:
                self.ao_salvar()
            self.destroy()


class JanelaSecoes(tk.Toplevel):
    """Lista as seções criadas pela pessoa (os 4 padrões do modelo não saem)."""

    def __init__(self, pai, galeria):
        super().__init__(pai)
        self.galeria = galeria
        aba = galeria.perfil.locais.nome
        self.title(f"Seções da aba {aba}")
        self.resizable(False, False)
        self.transient(pai)
        self.grab_set()
        corpo = ttk.Frame(self, padding=16)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text=f"Seções da aba “{aba}” ({galeria.perfil.banco})",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w")
        padroes = ", ".join(galeria.perfil.rotulos_locais().values())
        ttk.Label(corpo, foreground="#555", justify="left", wraplength=420,
                  text=f"Padrões do modelo (sempre existem): {padroes}.\n\n"
                       "Seções novas (cada uma vira um quadro novo, com o nome embaixo):").pack(anchor="w", pady=(4, 6))
        self.lista = tk.Listbox(corpo, height=8, width=44, font=("Segoe UI", 10), activestyle="none")
        self.lista.pack(fill="x")
        botoes = ttk.Frame(corpo)
        botoes.pack(fill="x", pady=(10, 0))
        ttk.Button(botoes, text="+ Nova seção…", command=self._nova).pack(side="left")
        ttk.Button(botoes, text="Apagar selecionada", command=self._apagar).pack(side="left", padx=8)
        ttk.Button(botoes, text="Fechar", command=self.destroy).pack(side="right")
        self._preencher()

    def _preencher(self):
        self.lista.delete(0, "end")
        for nome in self.galeria.secoes_extras:
            self.lista.insert("end", nome)
        if not self.galeria.secoes_extras:
            self.lista.insert("end", "(nenhuma ainda)")

    def _nova(self):
        self.galeria.nova_secao(pai=self)
        self._preencher()

    def _apagar(self):
        sel = [self.lista.get(i) for i in self.lista.curselection()]
        sel = [n for n in sel if n in self.galeria.secoes_extras]
        if sel and messagebox.askyesno(TITULO, f"Apagar a seção “{sel[0]}”?\n\nFotos que estavam nela voltam "
                                               "para a aba principal.", parent=self):
            self.galeria.remover_secoes(sel)
            self._preencher()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{TITULO}  (v{comum.VERSAO}" + (f" - {comum.ROTULO_EDICAO})" if comum.ROTULO_EDICAO else ")"))
        try:
            self.iconbitmap(os.path.join(comum.PASTA_RECURSOS, "assets", "icone.ico"))
        except tk.TclError:
            pass
        self.minsize(980, 700)
        self.geometry("1280x860")
        try:
            self.state("zoomed")  # abre maximizada: mais espaço pras fotos
        except tk.TclError:
            pass
        self.cfg = carregar_config()
        self._migrar_config()
        self.perfil = PERFIS.get(self.cfg.get("banco"), PERFIS[PADRAO])
        self.modelos = robo.listar_modelos(self.perfil)
        self.fila = queue.Queue()
        self.rodando = False
        self.ultimo_excel = None
        self.arquivo_log = None
        self.busca = None  # {"proposta", "itens", "cache", "inicio", "retomado"} da última busca
        self._trabalho = None  # trabalho salvo da proposta aberta (trabalho.py): {"proposta", "ultimo"}
        self._salvar_agendado = None
        self._versao_avisada = None

        visual.aplicar_tema(self)
        self._montar()
        self._atualizar_login()
        self.report_callback_exception = self._erro_na_janela
        self.after(100, self._ler_fila)
        self.after(1500, self._procurar_versao_nova)
        pasta_publica.enviar_em_segundo_plano()  # erros/uso que ficaram sem enviar da última vez
        threading.Thread(target=trabalho.limpar_antigos, daemon=True).start()
        self.protocol("WM_DELETE_WINDOW", self._fechar)

    # ------------------------------------------------------------ layout
    def _montar(self):
        self._montar_cabecalho()
        corpo = tk.Frame(self, bg=visual.PAGINA)
        corpo.pack(fill="both", expand=True, padx=18, pady=(14, 12))
        self._montar_busca(corpo)
        self.etapas = visual.Etapas(corpo)
        self.etapas.pack(anchor="w", pady=(12, 10))
        self._montar_galeria(corpo)
        self._montar_acoes(corpo)
        self._galeria_mudou()

    def _montar_cabecalho(self):
        cab = tk.Frame(self, bg=visual.CABECALHO, height=62)
        cab.pack(fill="x")
        cab.pack_propagate(False)
        try:
            from PIL import Image, ImageTk

            with Image.open(os.path.join(comum.PASTA_RECURSOS, "assets", "icone.ico")) as im:
                self._logo = ImageTk.PhotoImage(im.convert("RGBA").resize((34, 34), Image.LANCZOS))
            tk.Label(cab, image=self._logo, bg=visual.CABECALHO).pack(side="left", padx=(18, 10))
        except Exception:
            pass
        titulos = tk.Frame(cab, bg=visual.CABECALHO)
        titulos.pack(side="left")
        tk.Label(titulos, text="Downloader de Fotos", bg=visual.CABECALHO, fg="white",
                 font=(visual.FONTE, 15, "bold")).pack(anchor="w")
        tk.Label(titulos, text="Fotos da Central de Gestão direto no modelo de Excel do banco", bg=visual.CABECALHO,
                 fg=visual.CABECALHO_2, font=(visual.FONTE, 9)).pack(anchor="w")
        # direita: versão + login
        tk.Label(cab, text=f"v{comum.VERSAO}", bg=visual.CABECALHO, fg=visual.CABECALHO_2,
                 font=(visual.FONTE, 9)).pack(side="right", padx=(10, 18))
        # aparece só quando a pasta pública tem um .exe mais novo (_procurar_versao_nova)
        self.lnk_versao = tk.Label(cab, bg="#FFD666", fg="#1E1F22", font=(visual.FONTE, 9, "bold"),
                                   cursor="hand2", padx=10, pady=3)
        self.lnk_versao.bind("<Button-1>", lambda e: self._abrir_versao_nova())
        if comum.ROTULO_EDICAO:  # .exe de teste: deixa claro qual edição é
            tk.Label(cab, text=f"  {comum.ROTULO_EDICAO}  ", bg=visual.MARCA, fg="white",
                     font=(visual.FONTE, 9, "bold")).pack(side="right", padx=(10, 0))
        self.lnk_login = tk.Label(cab, text="Configurar login", bg=visual.CABECALHO, fg="#8EC5FF",
                                  font=(visual.FONTE, 9, "underline"), cursor="hand2")
        self.lnk_login.pack(side="right", padx=(10, 6))
        self.lnk_login.bind("<Button-1>", lambda e: self._abrir_login())
        self.lbl_login = tk.Label(cab, bg="#2B3A4C", fg="white", font=(visual.FONTE, 9), padx=10, pady=3)
        self.lbl_login.pack(side="right")
        # login da Intranet (tipo de operação do Banco D -> modelo)
        self.lnk_intranet = tk.Label(cab, bg=visual.CABECALHO, fg="#8EC5FF", font=(visual.FONTE, 9, "underline"),
                                 cursor="hand2")
        self.lnk_intranet.pack(side="right", padx=(10, 14))
        self.lnk_intranet.bind("<Button-1>", lambda e: JanelaLoginIntranet(self, ao_salvar=self._atualizar_login))
        tk.Frame(self, bg=visual.MARCA, height=3).pack(fill="x")

    def _montar_busca(self, pai):
        caixa = visual.cartao(pai)
        caixa.pack(fill="x")
        dentro = ttk.Frame(caixa, padding=(16, 12, 16, 14))
        dentro.pack(fill="x")
        dentro.columnconfigure(3, weight=1)

        # banco: botões lado a lado
        ttk.Label(dentro, text="BANCO", style="Legenda.TLabel").grid(row=0, column=0, sticky="w")
        bancos = ttk.Frame(dentro)
        bancos.grid(row=1, column=0, sticky="w", padx=(0, 22))
        self.var_banco = tk.StringVar(value=self.perfil.banco)
        for nome in PERFIS:
            ttk.Radiobutton(bancos, text=nome, value=nome, variable=self.var_banco, style="Toggle.TButton",
                            command=self._banco_mudou, width=10).pack(side="left")

        ttk.Label(dentro, text="MODELO DO EXCEL", style="Legenda.TLabel").grid(row=0, column=1, sticky="w")
        self.var_modelo = tk.StringVar()
        self.cmb_modelo = ttk.Combobox(dentro, textvariable=self.var_modelo, state="readonly", width=30)  # cabe "Consórcio reforma sem ampliação"
        self.cmb_modelo.grid(row=1, column=1, sticky="w", padx=(0, 22))
        self.cmb_modelo.bind("<<ComboboxSelected>>", lambda e: setattr(self, "_modelo_tocado", True))
        self._modelo_tocado = False
        self._preencher_modelos()
        # Banco D: de onde veio o modelo (tipo de operação achado na Intranet/planilha do banco)
        self.var_tipo = tk.StringVar(value="")
        self.lbl_tipo = ttk.Label(dentro, textvariable=self.var_tipo, style="Sub.TLabel")
        self.lbl_tipo.grid(row=2, column=1, columnspan=3, sticky="w", pady=(6, 0))

        ttk.Label(dentro, text="Nº DA PROPOSTA", style="Legenda.TLabel").grid(row=0, column=2, sticky="w")
        self.var_proposta = tk.StringVar()
        self.ent_proposta = ttk.Entry(dentro, textvariable=self.var_proposta, font=(visual.FONTE, 13), width=14)
        self.ent_proposta.grid(row=1, column=2, sticky="w")
        self.ent_proposta.bind("<Return>", lambda e: self._buscar())
        self.ent_proposta.focus_set()
        self.btn_buscar = ttk.Button(dentro, text="Buscar fotos", style="Grande.Accent.TButton", command=self._buscar)
        self.btn_buscar.grid(row=1, column=3, sticky="w", padx=(10, 0))

        ttk.Label(dentro, text="SALVAR EM", style="Legenda.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
        pasta = ttk.Frame(dentro)
        pasta.grid(row=3, column=0, columnspan=4, sticky="ew")
        pasta.columnconfigure(0, weight=1)
        self.var_destino = tk.StringVar(value=self.cfg.get("destino") or os.path.join(
            os.path.expanduser("~"), "Documents", "Fotos Propostas"))
        ttk.Entry(pasta, textvariable=self.var_destino).grid(row=0, column=0, sticky="ew")
        ttk.Button(pasta, text="Escolher pasta…", command=self._escolher_pasta).grid(row=0, column=1, padx=(8, 0))

    def _montar_galeria(self, pai):
        topo = visual.cartao(pai)
        topo.pack(fill="x")
        linha = ttk.Frame(topo, padding=(14, 8, 10, 4))
        linha.pack(fill="x")
        self.var_titulo_galeria = tk.StringVar(value="Fotos")
        ttk.Label(linha, textvariable=self.var_titulo_galeria, style="Titulo.TLabel").pack(side="left")
        self.chips = tk.Frame(linha, bg=visual.CARTAO)
        self.chips.pack(side="left", padx=(14, 0))
        self.chip_principal = visual.chip(self.chips, fundo=visual.VERDE_FUNDO, cor=visual.VERDE)
        self.chip_quadros = visual.chip(self.chips, fundo=visual.LARANJA_FUNDO, cor=visual.MARCA)
        self.chip_fora = visual.chip(self.chips)
        self.chip_erro = visual.chip(self.chips, fundo=visual.VERMELHO_FUNDO, cor=visual.VERMELHO)
        self.chip_obs = visual.chip(self.chips, fundo="#FFF1C2", cor="#7A5B00")  # fotos com observação
        self.var_resumo = tk.StringVar(value="")  # texto completo (log/testes)

        self.btn_secoes = ttk.Button(linha, text="Seções…", command=self._gerenciar_secoes)
        self.botoes_galeria = [
            self.btn_secoes,
            ttk.Button(linha, text="Sugerir quadros", command=self._sugerir),
            ttk.Button(linha, text="↺  Ordem original", command=lambda: self.galeria.ordem_original()),
            ttk.Button(linha, text="Desmarcar todas", command=lambda: self.galeria.marcar_todas(False)),
            ttk.Button(linha, text="Marcar todas", command=lambda: self.galeria.marcar_todas(True)),
            # fotos do computador além das da Central de Gestão (não vão pra Central de Gestão, só pro Excel)
            ttk.Button(linha, text="+  Fotos do computador…", style="Accent.TButton",
                       command=self._adicionar_fotos_pc),
        ]
        for b in self.botoes_galeria:
            b.pack(side="right", padx=(6, 0))
        self.lbl_ajuda = ttk.Label(topo, style="Sub.TLabel", wraplength=1500, justify="left", padding=(14, 0, 14, 8))
        self.lbl_ajuda.pack(fill="x")

        self.galeria = Galeria(pai, ao_mudar=self._galeria_mudou, perfil=self.perfil,
                               secoes_extras=self._secoes_do_banco(),
                               ao_mudar_secoes=self._secoes_mudaram)
        self.galeria.pack(fill="both", expand=True, pady=(8, 8))
        self._ajustar_ao_banco()

    def _montar_acoes(self, pai):
        caixa = visual.cartao(pai)
        caixa.pack(fill="x")
        dentro = ttk.Frame(caixa, padding=(14, 10, 14, 10))
        dentro.pack(fill="x")
        dentro.columnconfigure(0, weight=1)
        self.var_status = tk.StringVar(value="Digite o nº da proposta e clique em “Buscar fotos”.")
        self.lbl_status = ttk.Label(dentro, textvariable=self.var_status, style="Status.TLabel")
        self.lbl_status.grid(row=0, column=0, sticky="w")
        self.barra = ttk.Progressbar(dentro, maximum=1.0, mode="determinate")
        self.barra.grid(row=1, column=0, sticky="ew", pady=(6, 0), padx=(0, 18))
        self.btn_detalhes = ttk.Button(dentro, text="Detalhes ▾", style="Toolbutton", command=self._alternar_log)
        self.btn_detalhes.grid(row=0, column=1, rowspan=2, padx=(0, 6))
        self.btn_pasta = ttk.Button(dentro, text="Abrir pasta", command=self._abrir_pasta, state="disabled")
        self.btn_pasta.grid(row=0, column=2, rowspan=2, padx=(0, 6))
        self.btn_excel = ttk.Button(dentro, text="Abrir Excel", command=self._abrir_excel, state="disabled")
        self.btn_excel.grid(row=0, column=3, rowspan=2, padx=(0, 12))
        self.btn_gerar = ttk.Button(dentro, text="Gerar Excel com as fotos selecionadas",
                                    style="Grande.Accent.TButton", command=self._gerar, state="disabled")
        self.btn_gerar.grid(row=0, column=4, rowspan=2)

        # log (escondido atrás de "Detalhes")
        self.txt_log = ScrolledText(caixa, height=5, font=("Consolas", 8), state="disabled",
                                    relief="flat", borderwidth=0, wrap="word", bg="#F3F5F8")

    def _alternar_log(self):
        if self.txt_log.winfo_ismapped():
            self.txt_log.pack_forget()
            self.btn_detalhes.config(text="Detalhes ▾")
        else:
            self.txt_log.pack(fill="x", padx=14, pady=(0, 10))
            self.btn_detalhes.config(text="Detalhes ▴")

    # ------------------------------------------------------------ galeria
    def _galeria_mudou(self):
        if not getattr(self, "galeria", None) or not self.galeria.cartoes:
            self.var_resumo.set("Nenhuma proposta carregada.")
            self.var_titulo_galeria.set("Fotos")
            for c in (self.chip_principal, self.chip_quadros, self.chip_fora, self.chip_erro, self.chip_obs):
                c.pack_forget()
            if hasattr(self, "btn_gerar"):
                self.btn_gerar.config(state="disabled")
            for b in getattr(self, "botoes_galeria", []):
                b.state(["disabled"])
            return
        internas, locais = self.galeria.selecao()
        total = len(self.galeria.cartoes)
        fora = total - len(set(internas) | set(locais.values()))  # Banco D: foto de quadro conta nas duas
        rotulos = self.perfil.rotulos_locais()
        padroes = sum(1 for k in locais if k in rotulos)
        novas = len(locais) - padroes
        nao_baixaram = sum(1 for it in self.busca["itens"] if not it["caminho"])
        self.var_titulo_galeria.set(f"Proposta {self.busca['proposta']}  ·  {total} fotos")

        def mostrar(chip, texto):
            if texto:
                chip.config(text=texto)
                if not chip.winfo_ismapped():
                    chip.pack(side="left", padx=(0, 6))
            else:
                chip.pack_forget()

        mostrar(self.chip_principal, f"{len(internas)}  na aba {self.perfil.fotos.nome}")
        mostrar(self.chip_quadros, f"{padroes}/{len(rotulos)}  na aba {self.perfil.locais.nome}"
                + (f"  (+{novas} novas)" if novas else "") if self.perfil.locais else "")
        mostrar(self.chip_fora, f"{fora}  fora do Excel" if fora else "")
        mostrar(self.chip_erro, f"{nao_baixaram}  não baixaram" if nao_baixaram else "")
        com_obs = sum(1 for it in self.busca["itens"] if it.get("observacao"))
        mostrar(self.chip_obs, f"{com_obs}  com observação (OBS)" if com_obs else "")
        texto = f"{self.perfil.banco} • proposta {self.busca['proposta']}: {total} fotos  •  {len(internas)} na aba principal"
        if self.perfil.locais:
            texto += (f"  •  {padroes}/{len(rotulos)} na aba {self.perfil.locais.nome}"
                      + (f" (+{novas} em seções novas)" if novas else ""))
        texto += f"  •  {fora} fora do Excel"
        if nao_baixaram:
            texto += f"  •  {nao_baixaram} não baixaram"
        self.var_resumo.set(texto)
        for b in self.botoes_galeria:
            b.state(["!disabled"])
        if not (self.perfil.locais and self.perfil.locais.aceita_secoes_novas):
            self.btn_secoes.state(["disabled"])
        if not self.rodando:
            self.btn_gerar.config(state="normal" if (internas or locais) else "disabled")
        self._agendar_salvar_trabalho()

    # ------------------------------------------------------------ banco
    def _migrar_config(self):
        """Config de versões só-Banco C -> por banco."""
        if "secoes_extras" in self.cfg:
            self.cfg.setdefault("secoes_por_banco", {}).setdefault("Banco C", self.cfg.pop("secoes_extras"))
        if "modelo" in self.cfg:
            antigo = self.cfg.pop("modelo")
            antigo = PERFIS["Banco C"].nomes_antigos_modelos.get(antigo, antigo)
            self.cfg.setdefault("modelo_por_banco", {}).setdefault("Banco C", antigo)

    def _secoes_do_banco(self):
        return self.cfg.get("secoes_por_banco", {}).get(self.perfil.banco, [])

    def _preencher_modelos(self):
        nomes = list(self.modelos)
        self.cmb_modelo.config(values=nomes)
        salvo = self.cfg.get("modelo_por_banco", {}).get(self.perfil.banco)
        salvo = self.perfil.nomes_antigos_modelos.get(salvo, salvo)
        self.var_modelo.set(salvo if salvo in nomes else (nomes[0] if nomes else ""))

    def _banco_mudou(self):
        if self.var_banco.get() == self.perfil.banco:
            return
        self.perfil = PERFIS[self.var_banco.get()]
        self.var_tipo.set("")  # o aviso do tipo de operação era do outro banco
        self.modelos = robo.listar_modelos(self.perfil)
        self._preencher_modelos()
        self.galeria.definir_perfil(self.perfil, self._secoes_do_banco())
        self._ajustar_ao_banco()
        self.cfg["banco"] = self.perfil.banco
        salvar_config(self.cfg)
        if not self.modelos:
            messagebox.showwarning(TITULO, f"Nenhum modelo de Excel do {self.perfil.banco} encontrado "
                                           f"(pasta modelos\\{self.perfil.pasta_modelos}).")

    def _ajustar_ao_banco(self):
        """Textos/botões que dependem do banco (aba de quadros, seções novas)."""
        loc = self.perfil.locais
        ajuda = ("Clique na foto para ver grande  ·  “Incluir” tira/põe no Excel  ·  arraste (ou ◀ ▶) para "
                 "mudar a ordem  ·  clique no nome para renomear")
        if loc:
            ajuda += f"  ·  na caixa de baixo: aba principal ou um quadro da “{loc.nome}”"
            if loc.aceita_secoes_novas:
                ajuda += " (ou “+ Nova seção…”)"
                self.btn_secoes.pack(side="right", padx=(6, 0), before=self.botoes_galeria[1])
            else:
                self.btn_secoes.pack_forget()
            if self.perfil.quadros_tambem_na_principal:
                ajuda += f"  ·  a foto do quadro fica também na aba “{self.perfil.fotos.nome}”"
        self.lbl_ajuda.config(text=ajuda)
        self._galeria_mudou()

    def _secoes_mudaram(self, lista):
        self.cfg.setdefault("secoes_por_banco", {})[self.perfil.banco] = list(lista)
        salvar_config(self.cfg)

    def _gerenciar_secoes(self):
        JanelaSecoes(self, self.galeria)

    def _sugerir(self):
        if self.busca:
            self.galeria.aplicar_sugestao(robo.sugerir_locais(self.busca["itens"], self.perfil))

    # ------------------------------------------------------------ ações
    def _escolher_pasta(self):
        pasta = filedialog.askdirectory(title="Onde salvar as fotos?",
                                        initialdir=self.var_destino.get() or os.path.expanduser("~"))
        if pasta:
            self.var_destino.set(os.path.normpath(pasta))

    def _abrir_login(self):
        JanelaLogin(self, ao_salvar=self._atualizar_login)

    def _atualizar_login(self):
        usuario = comum.usuario_configurado()
        if usuario:
            self.lbl_login.config(text=f"●  {usuario}", fg="#A7E3B4")
            self.lnk_login.config(text="Trocar login")
        else:
            self.lbl_login.config(text="●  Login não configurado", fg="#FFC08A")
            self.lnk_login.config(text="Configurar login")
        if hasattr(self, "lnk_intranet"):
            self.lnk_intranet.config(text="Intranet ✓" if intranet.usuario_configurado() else "Configurar Intranet")

    # ------------------------------------------------------------ versão nova (pasta pública)
    VERIFICAR_VERSAO_MS = 30 * 60 * 1000  # de novo a cada 30 min (gente que deixa aberto o dia todo)

    def _procurar_versao_nova(self):
        """Olha a pasta pública em segundo plano (rede lenta não trava a janela)."""
        threading.Thread(target=lambda: self.fila.put(("versao", pasta_publica.versao_mais_nova())),
                         daemon=True).start()
        self.after(self.VERIFICAR_VERSAO_MS, self._procurar_versao_nova)

    def _versao_nova(self, achada):
        if not achada:
            return
        versao, caminho = achada
        self._exe_novo = caminho
        self.lnk_versao.config(text=f"↑  Versão {versao} disponível - clique para abrir")  # ⬆ não existe na Segoe UI
        if not self.lnk_versao.winfo_ismapped():
            self.lnk_versao.pack(side="right", padx=(10, 4))
        if self._versao_avisada != versao and not self.rodando and not self.busca:
            self._versao_avisada = versao  # pergunta uma vez; depois fica só a faixa amarela
            self._abrir_versao_nova(f"Saiu a versão {versao} do programa (esta é a {comum.VERSAO}).\n\n")

    def _abrir_versao_nova(self, intro=""):
        if self.rodando:
            messagebox.showinfo(TITULO, "Espere terminar o que está fazendo para abrir a versão nova.")
            return
        aviso = ("\n\nO que você fez na proposta aberta fica salvo: é só buscar de novo na versão nova."
                 if self.busca else "")
        if messagebox.askyesno(TITULO, f"{intro}Abrir a versão nova agora? Esta janela fecha.{aviso}"):
            try:
                os.startfile(self._exe_novo)
            except OSError as e:
                messagebox.showerror(TITULO, f"Não consegui abrir {self._exe_novo}:\n\n{e}")
                return
            self._fechar(perguntar=False)

    # ------------------------------------------------------------ erros (vão pra pasta pública)
    def _erro_na_janela(self, tipo, valor, tb):
        """Erro dentro de um clique/evento da janela (senão sumiria calado)."""
        detalhe = "".join(traceback.format_exception(tipo, valor, tb))
        self._log(detalhe)
        pasta_publica.erro(f"erro na janela: {valor}", detalhe)
        messagebox.showerror(TITULO, f"Aconteceu um erro inesperado:\n\n{valor}\n\n"
                                     "O responsável já foi avisado. Se continuar, feche e abra o programa.")

    def _contexto(self, proposta, etapa):
        """O que vai junto de um erro, se der (pasta_publica.erro)."""
        pasta_publica.contexto.clear()
        pasta_publica.contexto.update(proposta=proposta, etapa=etapa, banco=self.perfil.banco,
                                      modelo=self.var_modelo.get(), arquivo_log=self.arquivo_log)

    def _status(self, texto, tipo=""):
        """Mensagem da barra de baixo; tipo "ok" (verde) ou "erro" (vermelho)."""
        self.var_status.set(texto)
        self.lbl_status.config(foreground={"ok": visual.VERDE, "erro": visual.VERMELHO}.get(tipo, visual.TEXTO))

    def _abrir_pasta(self):
        if self.ultimo_excel:
            os.startfile(os.path.dirname(self.ultimo_excel))

    def _abrir_excel(self):
        if self.ultimo_excel and os.path.exists(self.ultimo_excel):
            os.startfile(self.ultimo_excel)

    def _ocupado(self, sim, texto_botao=None, buscando=False):
        self.rodando = sim
        self.btn_buscar.config(state="disabled" if sim else "normal",
                               text="Buscando…" if sim and buscando else "Buscar fotos")
        if sim:
            self.btn_gerar.config(state="disabled", text=texto_botao or "Trabalhando…")
            self.config(cursor="watch")
        else:
            self.btn_gerar.config(text="Gerar Excel com as fotos selecionadas")
            self.config(cursor="")
            self._galeria_mudou()

    def _novo_log(self, proposta, etapa):
        self._limpar_log()
        os.makedirs(comum.PASTA_LOGS, exist_ok=True)
        self.arquivo_log = os.path.join(
            comum.PASTA_LOGS, f"proposta_{proposta}_{etapa}_{datetime.now():%Y%m%d_%H%M%S}.txt")

    def _buscar(self):
        if self.rodando:
            return
        proposta = robo.limpar_proposta(self.var_proposta.get())
        if not proposta:
            messagebox.showwarning(TITULO, "Digite o número da proposta.")
            self.ent_proposta.focus_set()
            return
        if not comum.usuario_configurado():
            if messagebox.askyesno(TITULO, "O login da Central de Gestão ainda não foi configurado.\n\n"
                                           "Configurar agora?"):
                self._abrir_login()
            return
        self._descartar_busca()
        self.ultimo_excel = None
        self.btn_excel.config(state="disabled")
        self.btn_pasta.config(state="disabled")
        self.barra["value"] = 0
        self._novo_log(proposta, "busca")
        self._contexto(proposta, "buscar fotos")
        self._inicio_busca = datetime.now()
        self.etapas.marcar(1)
        self._status(f"Buscando a proposta {proposta} na Central de Gestão…")
        self._ocupado(True, buscando=True)
        threading.Thread(target=self._trabalho_buscar, args=(proposta,), daemon=True).start()
        self.var_tipo.set("")
        self._consulta_banco = None
        if self.perfil.pre_preenchimento:  # Banco D: descobre o modelo em paralelo com as fotos
            self._modelo_tocado = False
            self._aviso_tipo(f"Procurando o tipo de operação da proposta {proposta}…", "")
            # o resultado (tipo, junção...) também vai pro Excel: _trabalho_gerar espera por ele
            self._consulta_banco = {"proposta": proposta, "pronto": threading.Event(), "dados": None}
            threading.Thread(target=self._trabalho_tipo, args=(proposta, self.perfil, self._consulta_banco),
                             daemon=True).start()

    def _trabalho_tipo(self, proposta, perfil, consulta):
        try:
            dados = robo.descobrir_modelo(proposta, perfil, log=lambda m: self.fila.put(("log", m)))
            consulta["dados"] = dados
            self.fila.put(("tipo_ok", (proposta, perfil.banco, dados)))
        except Exception as e:
            self.fila.put(("tipo_erro", (proposta, perfil.banco, str(e))))
        finally:
            consulta["pronto"].set()

    def _dados_do_banco(self, proposta, log):
        """O que veio da planilha do banco pra esta proposta ({} se nada).
        Se a consulta ainda estiver rodando, espera um pouco (o Excel sairia
        sem a junção)."""
        c = getattr(self, "_consulta_banco", None)
        if not c or c["proposta"] != proposta:
            return {}
        if not c["pronto"].is_set():
            log("[MODELO] Esperando a consulta à Intranet terminar (junção)...")
            c["pronto"].wait(30)
        return c["dados"] or {}

    def _aviso_tipo(self, texto, tipo):
        self.var_tipo.set(texto)
        self.lbl_tipo.config(foreground={"ok": visual.VERDE, "erro": visual.MARCA}.get(tipo, visual.TEXTO_2))

    def _tipo_descoberto(self, proposta, banco, dados):
        if banco != self.perfil.banco or proposta != robo.limpar_proposta(self.var_proposta.get()):
            return  # a pessoa já trocou de banco/proposta
        tipo, modelo = dados.get("tipo_operacao"), dados.get("modelo")
        self._log(f"[MODELO] {dados['fonte']}: tipo de operação {tipo!r} -> modelo {modelo or '(nenhum)'}")
        if not modelo:
            self._aviso_tipo(f"⚠  {dados['fonte']}: tipo de operação “{tipo}” não tem modelo - escolha o modelo.", "erro")
        elif self._modelo_tocado and self.var_modelo.get() != modelo:
            self._aviso_tipo(f"⚠  {dados['fonte']}: “{tipo}” → o modelo certo seria {modelo} "
                             f"(você escolheu {self.var_modelo.get()}).", "erro")
        else:
            self.var_modelo.set(modelo)
            self._aviso_tipo(f"✓  {dados['fonte']}: {tipo} → modelo {modelo} selecionado.", "ok")

    def _tipo_nao_achado(self, proposta, banco, motivo):
        if banco != self.perfil.banco or proposta != robo.limpar_proposta(self.var_proposta.get()):
            return
        self._log(f"[MODELO] Não deu pra descobrir o tipo de operação: {motivo}")
        if "não configurado" in motivo:
            texto = "⚠  Configure o login da Intranet (no alto, à direita) para o modelo vir sozinho - escolha o modelo."
        else:
            texto = "⚠  Não achei o tipo de operação desta proposta (veja Detalhes) - confira o modelo."
        self._aviso_tipo(texto, "erro")

    def _trabalho_buscar(self, proposta):
        try:
            proposta, itens, cache = robo.buscar(
                proposta, log=lambda m: self.fila.put(("log", m)),
                progresso=lambda f, t: self.fila.put(("prog", (f * 0.9, t))))
            # trabalho salvo desta proposta: já prepara as fotos do computador que ele usa
            # (a pergunta "retomar?" é feita na janela; se a pessoa não quiser, elas ficam de fora)
            salvo, fotos_pc = trabalho.ler(proposta), []
            if salvo:
                guardadas = trabalho.fotos_pc_salvas(proposta, salvo)
                if guardadas:
                    fotos_pc, _ = robo.adicionar_fotos_pc(list(itens), cache, guardadas)
            self.fila.put(("prog", (0.95, "Preparando as miniaturas...")))
            miniaturas = gerar_miniaturas(itens + fotos_pc)
            self.fila.put(("busca_ok", (proposta, itens, cache, miniaturas, salvo, fotos_pc)))
        except Exception as e:
            self._enfileirar_erro(e)

    def _mostrar_busca(self, proposta, itens, cache, miniaturas, salvo=None, fotos_pc=()):
        self.busca = {"proposta": proposta, "itens": itens, "cache": cache,
                      "inicio": getattr(self, "_inicio_busca", datetime.now()), "retomado": False}
        self._trabalho = None
        # a aba "Mapa e fotos locais" começa vazia: quem escolhe as 4 é a pessoa
        # (o botão "Sugerir fotos locais" preenche pela legenda, se ela quiser)
        self.galeria.carregar(itens, miniaturas, {})
        self.barra["value"] = 1.0
        retomada = self._perguntar_retomar(proposta, itens, miniaturas, salvo, fotos_pc) if salvo else ""
        # daqui pra frente, cada mudança na galeria fica salva (_agendar_salvar_trabalho)
        self._trabalho = {"proposta": proposta, "ultimo": self.galeria.estado()}
        baixadas = sum(1 for it in itens if it["caminho"])
        # observações das fotos na Central de Gestão: avisa logo (pedido de 14/10/2026)
        # (ambiente, observação) sem repetir, na ordem da Central de Gestão - o mesmo ambiente pode ter
        # várias observações (ex.: "Banheiro Suíte: 1" e "Banheiro Suíte: 2")
        ambientes = list(dict.fromkeys((it["legenda"].rsplit(" (", 1)[0], it["observacao"])
                                       for it in itens if it.get("observacao")))
        obs = (f"  •  ⚠ {len(ambientes)} observação(ões) nas fotos (selo OBS)" if ambientes
               else "  •  nenhuma foto com observação")
        self._status(f"{baixadas} foto(s) prontas{retomada}{obs}. Escolha, ordene e clique em “Gerar Excel”.")
        self.etapas.marcar(2)
        self._ocupado(False)
        for falha in robo.falhas_de(itens):
            self._log(f"[AVISO] Não baixou: {falha}")
        if ambientes:
            lista = "\n".join(f"•  {amb}:  {texto}" for amb, texto in ambientes[:15])
            mais = f"\n… e mais {len(ambientes) - 15}" if len(ambientes) > 15 else ""
            messagebox.showinfo(TITULO, f"A proposta {proposta} tem {len(ambientes)} observação(ões) nas fotos:"
                                        f"\n\n{lista}{mais}\n\nAs fotos estão marcadas com OBS na galeria "
                                        "(clique na foto para ver a observação).")

    # ------------------------------------------------------------ trabalho salvo (trabalho.py)
    def _perguntar_retomar(self, proposta, itens, miniaturas, salvo, fotos_pc):
        """Proposta já trabalhada neste PC: pergunta se volta como estava.
        Devolve o texto pra barra de status ("" se não retomou)."""
        quando, oque = trabalho.resumo(salvo)
        banco = salvo.get("banco")
        outro_banco = f"\n(na época, banco {banco})" if banco and banco != self.perfil.banco else ""
        if not messagebox.askyesno(
                TITULO, f"Você já trabalhou na proposta {proposta} neste computador - {quando}:\n"
                        f"{oque}.{outro_banco}\n\n"
                        "Retomar de onde parou?\n\n"
                        "Sim = volta as fotos escolhidas, a ordem, os nomes e as seções.\n"
                        "Não = começa do zero, na ordem da Central de Gestão (o que estava salvo só é "
                        "substituído quando você mexer em alguma foto)."):
            self._log(f"[TRABALHO] Havia trabalho salvo ({quando}); a pessoa preferiu começar do zero.")
            return ""
        if fotos_pc:
            itens.extend(fotos_pc)
            self.galeria.carregar(itens, miniaturas, {})
        retomadas, sumidas, novas = self.galeria.aplicar_estado(salvo)
        self.busca["retomado"] = True
        self._log(f"[TRABALHO] Retomado ({quando}): {oque}; {retomadas} foto(s) reconhecidas, "
                  f"{sumidas} não existem mais, {novas} nova(s) na Central de Gestão.")
        avisos = []
        if sumidas:
            avisos.append(f"{sumidas} foto(s) da última vez não existem mais na Central de Gestão")
        if novas:
            avisos.append(f"{novas} foto(s) novas na Central de Gestão foram para o fim da galeria")
        if avisos:
            messagebox.showinfo(TITULO, "Trabalho retomado.\n\n• " + "\n• ".join(avisos) + ".")
        return "  •  ✓ trabalho anterior retomado"

    def _agendar_salvar_trabalho(self):
        """Salva um instante depois da última mudança (arrastar várias vezes = um salvamento)."""
        if not self._trabalho:
            return
        if self._salvar_agendado:
            self.after_cancel(self._salvar_agendado)
        self._salvar_agendado = self.after(800, self._salvar_trabalho)

    def _salvar_trabalho(self, **extra):
        """Grava o estado da galeria se mudou (ou se veio `extra`, ex.: gerado_em).
        Nunca atrapalha a pessoa: erro aqui só vai pro log."""
        self._salvar_agendado = None
        t = self._trabalho
        if not t or not self.busca or self.busca["proposta"] != t["proposta"] or not self.galeria.cartoes:
            return
        if self.rodando:  # ex.: fotos do computador sendo copiadas - salva quando terminar
            self._agendar_salvar_trabalho()
            return
        estado = self.galeria.estado()
        if estado == t["ultimo"] and not extra:
            return  # nada mudou desde o último salvamento (ou desde que abriu)
        try:
            trabalho.salvar(t["proposta"], estado, banco=self.perfil.banco, **extra)
            t["ultimo"] = estado
        except Exception as e:
            self._log(f"[TRABALHO] [AVISO] Não consegui salvar o trabalho desta proposta: {e}")

    def _adicionar_fotos_pc(self):
        """Escolhe fotos no computador e põe na galeria (no fim, já incluídas)."""
        if self.rodando or not self.busca:
            return
        tipos = " ".join("*" + e for e in robo.EXTENSOES_FOTO_PC)
        arquivos = filedialog.askopenfilenames(
            parent=self, title="Adicionar fotos do computador",
            initialdir=self.cfg.get("pasta_fotos_pc") or os.path.expanduser("~"),
            filetypes=[("Fotos", tipos), ("Todos os arquivos", "*.*")])
        if not arquivos:
            return
        self.cfg["pasta_fotos_pc"] = os.path.dirname(arquivos[0])
        salvar_config(self.cfg)
        self._status(f"Adicionando {len(arquivos)} foto(s) do computador…")
        self._ocupado(True)
        threading.Thread(target=self._trabalho_fotos_pc, args=(list(arquivos),), daemon=True).start()

    def _trabalho_fotos_pc(self, arquivos):
        try:
            proposta = self.busca["proposta"]
            # cópia guardada no trabalho salvo: a foto volta mesmo se a original for apagada
            novos, recusados = robo.adicionar_fotos_pc(self.busca["itens"], self.busca["cache"], arquivos,
                                                       guardar=lambda a: trabalho.guardar_foto_pc(proposta, a))
            self.fila.put(("pc_ok", (novos, recusados, gerar_miniaturas(novos))))
        except Exception as e:
            self._enfileirar_erro(e)

    def _fotos_pc_prontas(self, novos, recusados, miniaturas):
        self._ocupado(False)
        self.galeria.adicionar(novos, miniaturas)
        for it in novos:
            self._log(f"[PC] Foto do computador adicionada: {it['legenda']}")
        for arquivo, motivo in recusados:
            self._log(f"[PC] [AVISO] Não adicionada: {os.path.basename(arquivo)} - {motivo}")
        self._status(f"{len(novos)} foto(s) do computador adicionada(s) no fim da galeria."
                     + (f" {len(recusados)} recusada(s) - veja Detalhes." if recusados else ""),
                     "erro" if recusados and not novos else "")
        if recusados:
            messagebox.showwarning(TITULO, "Estes arquivos não foram adicionados (não são fotos que o programa "
                                           "consegue abrir):\n\n"
                                   + "\n".join(os.path.basename(a) for a, _ in recusados[:15]))

    def _gerar(self):
        if self.rodando or not self.busca:
            return
        destino = self.var_destino.get().strip()
        modelo = self.modelos.get(self.var_modelo.get())
        if not destino:
            messagebox.showwarning(TITULO, "Escolha a pasta onde salvar.")
            return
        if not modelo:
            messagebox.showerror(TITULO, "Nenhum modelo de Excel (.xlsm) encontrado na pasta 'modelos'.")
            return
        internas, locais = self.galeria.selecao()
        if not internas and not locais:
            messagebox.showwarning(TITULO, "Selecione pelo menos uma foto.")
            return
        legendas = self.galeria.legendas()
        proposta = self.busca["proposta"]
        faltam = [r for k, r in self.perfil.rotulos_locais().items() if k not in locais]
        if faltam and not messagebox.askyesno(
                TITULO, f"Estes quadros da aba “{self.perfil.locais.nome}” vão ficar sem foto:\n\n• "
                        + "\n• ".join(faltam) + "\n\nGerar assim mesmo?"):
            return
        _, _, excel = robo.caminhos_saida(proposta, destino, modelo, self.perfil)
        if os.path.exists(excel) and not messagebox.askyesno(
                TITULO, f"Já existe o Excel desta proposta com este modelo:\n{excel}\n\n"
                        "Substituir (e trocar as fotos da pasta pelas selecionadas agora)?"):
            return
        self.cfg.update(destino=destino, banco=self.perfil.banco)
        self.cfg.setdefault("modelo_por_banco", {})[self.perfil.banco] = self.var_modelo.get()
        salvar_config(self.cfg)
        self.barra["value"] = 0
        self._novo_log(proposta, "excel")
        self._contexto(proposta, "gerar Excel")
        itens = self.busca["itens"]
        self._gerando = {  # vai pro registro de uso (pasta_publica.uso) quando o Excel ficar pronto
            "banco": self.perfil.banco, "modelo": self.var_modelo.get(), "proposta": proposta,
            "fotos_central": sum(1 for it in itens if it.get("origem") != "pc"),
            "fotos_no_excel": len(set(internas) | set(locais.values())),
            "fotos_do_pc": sum(1 for it in itens if it.get("origem") == "pc"),
            "observacoes": sum(1 for it in itens if it.get("observacao")),
            "retomado": "sim" if self.busca.get("retomado") else "não",
        }
        self._ocupado(True, "Gerando o Excel…")
        threading.Thread(target=self._trabalho_gerar,
                         args=(proposta, destino, modelo, internas, locais, legendas, self.perfil),
                         daemon=True).start()

    def _trabalho_gerar(self, proposta, destino, modelo, internas, locais, legendas, perfil):
        try:
            log = lambda m: self.fila.put(("log", m))
            dados = robo.dados_da_busca(self.busca["cache"])
            if perfil.pre_preenchimento:
                dados["banco"] = self._dados_do_banco(proposta, log)
            excel = robo.gerar(proposta, destino, modelo, self.busca["itens"], internas, locais,
                               log=log, progresso=lambda f, t: self.fila.put(("prog", (f, t))), legendas=legendas,
                               perfil=perfil, dados=dados)
            self.fila.put(("excel_ok", excel))
        except Exception as e:
            self._enfileirar_erro(e)

    def _enfileirar_erro(self, e):
        if isinstance(e, (ErroCentralGestao, ValueError)):  # erro "conhecido": a mensagem já é amigável
            self.fila.put(("log", f"[ERRO] {e}"))
            self.fila.put(("erro", (str(e), isinstance(e, ErroLogin))))
        else:
            self.fila.put(("log", traceback.format_exc()))
            self.fila.put(("erro", (_mensagem_erro(e), False)))
            if not isinstance(e, PermissionError):  # Excel da proposta aberto: não é defeito do programa
                pasta_publica.erro(_mensagem_erro(e))

    def _descartar_busca(self):
        if self._salvar_agendado:  # salva a última mudança da proposta anterior antes de trocar
            self.after_cancel(self._salvar_agendado)
            self._salvar_trabalho()
        self._trabalho = None
        if self.busca:
            shutil.rmtree(self.busca["cache"], ignore_errors=True)
        self.busca = None
        self.galeria.limpar()

    # ------------------------------------------------------------ fila
    def _ler_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self._log(dado)
                elif tipo == "prog":
                    self.barra["value"], texto = dado
                    self._status(texto)
                elif tipo == "busca_ok":
                    self._mostrar_busca(*dado)
                elif tipo == "tipo_ok":
                    self._tipo_descoberto(*dado)
                elif tipo == "tipo_erro":
                    self._tipo_nao_achado(*dado)
                elif tipo == "pc_ok":
                    self._fotos_pc_prontas(*dado)
                elif tipo == "excel_ok":
                    self._excel_pronto(dado)
                elif tipo == "erro":
                    self._falhou(*dado)
                elif tipo == "versao":
                    self._versao_nova(dado)
        except queue.Empty:
            pass
        self.after(100, self._ler_fila)

    def _excel_pronto(self, excel):
        self._ocupado(False)
        self.ultimo_excel = excel
        self.btn_excel.config(state="normal")
        self.btn_pasta.config(state="normal")
        self.etapas.marcar(3, tudo_feito=True)
        self._status(f"✓  {os.path.basename(excel)} gerado. Dá pra mudar a seleção e gerar de novo.", "ok")
        self._salvar_trabalho(gerado_em=datetime.now().isoformat(timespec="seconds"), excel=excel)
        if self.busca:
            pasta_publica.uso(**getattr(self, "_gerando", {}),
                              segundos=round((datetime.now() - self.busca["inicio"]).total_seconds()))
        if messagebox.askyesno(TITULO, f"Pronto!\n\n{excel}\n\nAbrir o Excel agora?"):
            self._abrir_excel()

    def _falhou(self, mensagem, de_login):
        self._ocupado(False)
        self.barra["value"] = 0
        self._status("Não foi possível concluir - veja a mensagem.", "erro")
        if de_login:
            if messagebox.askyesno(TITULO, f"{mensagem}\n\nAbrir a configuração do login agora?"):
                self._abrir_login()
        else:
            messagebox.showerror(TITULO, f"Não foi possível concluir:\n\n{mensagem}")

    def _log(self, msg):
        linha = f"{datetime.now():%H:%M:%S}  {msg}".rstrip() + "\n"
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", linha)
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")
        if self.arquivo_log:
            try:
                with open(self.arquivo_log, "a", encoding="utf-8") as f:
                    f.write(linha)
            except Exception:
                pass

    def _limpar_log(self):
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")

    def _fechar(self, perguntar=True):
        if perguntar and self.rodando and not messagebox.askyesno(TITULO, "Ainda está trabalhando. Fechar mesmo assim?"):
            return
        if self._salvar_agendado:  # a última mudança ainda não tinha sido salva
            self.after_cancel(self._salvar_agendado)
            self.rodando = False
            self._salvar_trabalho()
        if self.busca:
            shutil.rmtree(self.busca["cache"], ignore_errors=True)
        self.destroy()


def modo_linha_de_comando(args):
    """DownloaderFotos.exe --proposta 123 --destino C:\\pasta [--modelo Nome]
    Roda sem janela e grava o log em %LOCALAPPDATA%\\DownloaderFotos\\logs
    (o .exe não tem console). Serve pra testar o .exe e pra suporte."""
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--proposta", required=True)
    ap.add_argument("--destino", required=True)
    ap.add_argument("--modelo", default=None)
    ap.add_argument("--banco", default=PADRAO, choices=list(PERFIS))
    ap.add_argument("--log", default=None, help="arquivo do log (padrão: na pasta de logs) - usado pelos testes")
    a = ap.parse_args(args)
    pasta_publica.ATIVO = False  # testes/suporte não entram no uso nem nos erros da pasta pública
    perfil = PERFIS[a.banco]
    modelos = robo.listar_modelos(perfil)
    os.makedirs(comum.PASTA_LOGS, exist_ok=True)
    caminho_log = a.log or os.path.join(comum.PASTA_LOGS, f"cli_{a.proposta}_{datetime.now():%Y%m%d_%H%M%S}.txt")
    with open(caminho_log, "w", encoding="utf-8") as log:
        escrever = lambda m: (log.write(f"{datetime.now():%H:%M:%S}  {m}\n"), log.flush())
        escrever(f"modelos: {list(modelos)}")
        achado = {}
        if perfil.pre_preenchimento:  # Banco D: tipo de operação (modelo, se não veio --modelo) e junção
            try:
                achado = robo.descobrir_modelo(robo.limpar_proposta(a.proposta), perfil, log=escrever)
                a.modelo = a.modelo or achado.get("modelo")
                escrever(f"[MODELO] {achado['fonte']}: {achado.get('tipo_operacao')!r} -> {a.modelo}")
            except Exception as e:
                escrever(f"[MODELO] não descobri o tipo de operação: {e}")
        try:
            excel, falhas = robo.executar(a.proposta, a.destino, modelos[a.modelo or next(iter(modelos))],
                                          log=escrever, perfil=perfil, dados_banco=achado)
            escrever(f"RESULTADO OK {excel} falhas={len(falhas)}")
            return 2 if falhas else 0
        except Exception:
            escrever("RESULTADO ERRO\n" + traceback.format_exc())
            return 1


if __name__ == "__main__":
    import sys

    if "--proposta" in sys.argv:
        raise SystemExit(modo_linha_de_comando(sys.argv[1:]))
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # texto nítido em tela com zoom
    except Exception:
        pass
    App().mainloop()
