"""
Galeria de escolha das fotos (parte da janela principal).

Cada foto vira um "cartão" com miniatura, nome (editável), "Incluir" e
seção de destino: aba principal ou um quadro da aba de quadros do banco
(Banco C: "Mapa e fotos locais"; Banco D: "PAG.1") - os do modelo ou seções
novas criadas pela pessoa em "+ Nova seção…" (ficam salvas como padrão
neste PC; cada uma vira um quadro novo, com o nome dela).

A ordem começa igual à da Central de Gestão. Arrastar a miniatura muda a ordem:
o próprio cartão vai junto com o mouse, uma barra azul mostra onde ele vai
cair e a lista rola sozinha perto das bordas. Também dá pra usar ◀ ▶.
Clique (sem arrastar) abre a foto grande dentro do programa.

Desempenho (08-09/10/2026): a galeria é UM Canvas e cada cartão é um grupo
de itens desenhados nele (retângulos, textos, a miniatura) - não widgets.
Com widgets, 100 fotos davam ~1.000 janelinhas do Windows e a rolagem
redesenhava uma por uma ("parece que está carregando de novo"). Desenhado,
rolar é só deslizar a imagem. O nome é editado num campo que aparece só na
hora do clique; a seção, num menu.
"""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, simpledialog, ttk

from PIL import Image, ImageEnhance, ImageOps, ImageTk

import visual
from perfis import BANCO_C

MINIATURA = (184, 138)
W = MINIATURA[0] + 30   # largura do cartão
H = 272                 # altura do cartão
GAP = 14                # espaço entre cartões
FAIXA_H = 24
LARGURA_CARTAO = W
COR_FUNDO = visual.PAGINA
COR_CARTAO = visual.CARTAO
COR_BORDA = visual.BORDA
COR_LOCAL = visual.MARCA
COR_ALVO = visual.AZUL
COR_FORA = "#C9CED6"
COR_CAMPO = "#FFFFFF"
COR_HOVER = "#E6EAF0"
COR_OBS = "#FFD666"  # foto com observação na Central de Gestão (selo no cartão e faixa na foto grande)
# faixa colorida do topo do cartão: (fundo, texto)
FAIXA = {
    "principal": (visual.VERDE, "white"),
    "local": (visual.MARCA, "white"),
    "fora": ("#A3ABB7", "white"),
    "erro": (visual.VERMELHO, "white"),
    "movendo": (visual.AZUL, "white"),
}

DESTINO_INTERNA = "Aba principal"
ANIMACAO_MS = 180        # duração do deslize dos cartões
ANIMACAO_QUADROS = 12
NOVA_SECAO = "+ Nova seção…"
PREFIXO_EXTRA = "extra:"

# posições (relativas ao canto do cartão)
Y_FOTO = FAIXA_H + 8
Y_NOME = Y_FOTO + MINIATURA[1] + 8          # campo do nome: 24 de altura
Y_LINHA = Y_NOME + 24 + 16                  # centro da linha do interruptor
Y_DEST = Y_LINHA + 16                       # campo da seção: 26 de altura


def gerar_miniaturas(itens):
    """Roda fora da thread da janela: {indice: (normal, apagada)} em PIL."""
    miniaturas = {}
    for it in itens:
        if not it["caminho"]:
            continue
        try:
            with Image.open(it["caminho"]) as img:
                img.draft("RGB", (MINIATURA[0] * 2, MINIATURA[1] * 2))  # JPEG grande: decodifica já reduzido
                img = ImageOps.exif_transpose(img).convert("RGB")
                img.thumbnail(MINIATURA, Image.LANCZOS)
            fundo = Image.new("RGB", MINIATURA, "#EEF1F4")
            fundo.paste(img, ((MINIATURA[0] - img.width) // 2, (MINIATURA[1] - img.height) // 2))
            apagada = ImageEnhance.Brightness(ImageOps.grayscale(fundo).convert("RGB")).enhance(1.35)
            miniaturas[it["indice"]] = (fundo, apagada)
        except Exception:
            pass
    return miniaturas


def _arredondado(cv, x1, y1, x2, y2, r, **kw):
    """Retângulo de cantos arredondados (polígono suavizado)."""
    p = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
         x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(p, smooth=True, **kw)


def _encurtar(texto, fonte, largura):
    """Corta com "…" pra caber em `largura` pixels."""
    if fonte.measure(texto) <= largura:
        return texto
    while texto and fonte.measure(texto + "…") > largura:
        texto = texto[:-1]
    return texto.rstrip() + "…"


class Cartao:
    """Uma foto da galeria: estado (variáveis Tk) + os itens desenhados no Canvas."""

    def __init__(self, galeria, item, imagens, x, y):
        self.galeria = galeria
        cv = galeria.canvas
        self.item = item
        self.indice = item["indice"]
        self.imagens = imagens  # (PhotoImage normal, PhotoImage apagada) ou None
        self.var_incluir = tk.BooleanVar(master=cv, value=bool(item["caminho"]))
        self.var_destino = tk.StringVar(master=cv, value=DESTINO_INTERNA)
        self.var_legenda = tk.StringVar(master=cv, value=item["legenda"])
        self._destino_anterior = DESTINO_INTERNA
        self.tag = f"c{self.indice}"
        self.x, self.y = x, y
        self._estado = None
        self._borda = COR_BORDA
        self._desenhar()
        self.var_legenda.trace_add("write", lambda *a: self._desenhar_nome())

    # ------------------------------------------------------------ desenho
    def _desenhar(self):
        g, cv, t, x, y = self.galeria, self.galeria.canvas, self.tag, self.x, self.y
        ok = bool(self.item["caminho"])
        self.i_fundo = _arredondado(cv, x, y, x + W, y + H, 7, fill=COR_CARTAO, outline=COR_BORDA, width=2,
                                    tags=(t, "cartao"))
        self.i_faixa = cv.create_rectangle(x + 2, y + 2, x + W - 2, y + FAIXA_H, width=0,
                                           fill=FAIXA["principal"][0], tags=(t, "cartao"))
        self.i_rotulo = cv.create_text(x + 10, y + FAIXA_H / 2 + 1, anchor="w", fill="white", font=g.f_faixa,
                                       tags=(t, "cartao"))
        # nº na Central de Gestão; foto escolhida no computador não tem
        self.i_num = cv.create_text(x + W - 10, y + FAIXA_H / 2 + 1, anchor="e",
                                    text="DO PC" if self.item.get("origem") == "pc" else f"#{self.indice + 1}",
                                    fill="#EAF2EC", font=g.f_num, tags=(t, "cartao"))
        if self.imagens:
            self.i_foto = cv.create_image(x + W / 2, y + Y_FOTO, anchor="n", image=self.imagens[0],
                                          tags=(t, "cartao", f"{t}:foto"))
            if self.item.get("observacao"):  # selo: esta foto tem observação na Central de Gestão
                _arredondado(cv, x + 19, y + Y_FOTO + 5, x + 64, y + Y_FOTO + 23, 4, fill=COR_OBS,
                             outline="#C9A227", width=1, tags=(t, "cartao", f"{t}:foto"))
                cv.create_text(x + 41.5, y + Y_FOTO + 14, text="OBS", fill="#1E1F22", font=g.f_faixa,
                               tags=(t, "cartao", f"{t}:foto"))
        else:
            self.i_foto = cv.create_rectangle(x + 15, y + Y_FOTO, x + W - 15, y + Y_FOTO + MINIATURA[1],
                                              fill=visual.VERMELHO_FUNDO, outline="", tags=(t, "cartao"))
            cv.create_text(x + W / 2, y + Y_FOTO + MINIATURA[1] / 2, width=W - 50, justify="center",
                           text=f"Não baixou\n\n{self.item['erro'] or ''}", fill=visual.VERMELHO,
                           font=g.f_txt, tags=(t, "cartao"))
        # nome
        _arredondado(cv, x + 12, y + Y_NOME, x + W - 12, y + Y_NOME + 24, 5, fill=COR_CAMPO, outline=COR_BORDA,
                     tags=(t, "cartao", f"{t}:nome"))
        self.i_nome = cv.create_text(x + W / 2 - 6, y + Y_NOME + 12, font=g.f_nome, fill=visual.TEXTO,
                                     tags=(t, "cartao", f"{t}:nome"))
        self.i_volta = cv.create_text(x + W - 22, y + Y_NOME + 12, text="↺", font=g.f_btn, fill=visual.AZUL,
                                      state="hidden", tags=(t, "cartao", f"{t}:volta"))
        # interruptor "Incluir" (a linha toda é clicável, não só a bolinha)
        cv.create_rectangle(x + 10, y + Y_LINHA - 12, x + W - 74, y + Y_LINHA + 12, fill=COR_CARTAO, width=0,
                            tags=(t, "cartao", f"{t}:sw"))
        self.i_sw = _arredondado(cv, x + 14, y + Y_LINHA - 9, x + 48, y + Y_LINHA + 9, 9, width=1,
                                 tags=(t, "cartao", f"{t}:sw"))
        self.i_bola = cv.create_oval(0, 0, 1, 1, width=0, tags=(t, "cartao", f"{t}:sw"))
        self.i_sw_txt = cv.create_text(x + 56, y + Y_LINHA, anchor="w", text="Incluir", font=g.f_txt,
                                       fill=visual.TEXTO, tags=(t, "cartao", f"{t}:sw"))
        # ◀ ▶
        for nome, cx, seta in (("esq", W - 56, "◀"), ("dir", W - 26, "▶")):
            _arredondado(cv, x + cx - 13, y + Y_LINHA - 11, x + cx + 13, y + Y_LINHA + 11, 5, fill=COR_CARTAO,
                         outline="", tags=(t, "cartao", f"{t}:{nome}", f"{t}:{nome}bg"))
            cv.create_text(x + cx, y + Y_LINHA, text=seta, font=g.f_seta, fill=visual.TEXTO_2,
                           tags=(t, "cartao", f"{t}:{nome}"))
        # seção de destino
        _arredondado(cv, x + 12, y + Y_DEST, x + W - 12, y + Y_DEST + 26, 5, fill=COR_CAMPO, outline=COR_BORDA,
                     tags=(t, "cartao", f"{t}:dest", f"{t}:destbg"))
        self.i_dest = cv.create_text(x + 22, y + Y_DEST + 13, anchor="w", font=g.f_txt, fill=visual.TEXTO,
                                     tags=(t, "cartao", f"{t}:dest"))
        cv.create_text(x + W - 24, y + Y_DEST + 13, text="▾", font=g.f_btn, fill=visual.TEXTO_2,
                       tags=(t, "cartao", f"{t}:dest"))
        if ok:
            cv.tag_bind(f"{t}:foto", "<ButtonPress-1>", lambda e: g.inicio_arrasto(self.indice, e))
            cv.tag_bind(f"{t}:foto", "<B1-Motion>", g.movendo)
            cv.tag_bind(f"{t}:foto", "<ButtonRelease-1>", g.fim_arrasto)
            cv.tag_bind(f"{t}:foto", "<Enter>", lambda e: self._realce(True))
            cv.tag_bind(f"{t}:foto", "<Leave>", lambda e: self._realce(False))
            cv.tag_bind(f"{t}:sw", "<Button-1>", lambda e: self.alternar())
            cv.tag_bind(f"{t}:nome", "<Button-1>", lambda e: g.editar_nome(self))
            cv.tag_bind(f"{t}:volta", "<Button-1>", lambda e: self.var_legenda.set(self.item["legenda"]))
            cv.tag_bind(f"{t}:dest", "<Button-1>", lambda e: g.menu_destino(self, e))
            for zona in ("sw", "nome", "volta", "dest", "esq", "dir"):
                cv.tag_bind(f"{t}:{zona}", "<Enter>", lambda e, z=zona: g.passou_mouse(self, z, True))
                cv.tag_bind(f"{t}:{zona}", "<Leave>", lambda e, z=zona: g.passou_mouse(self, z, False))
        cv.tag_bind(f"{t}:esq", "<Button-1>", lambda e: g.mover(self.indice, -1))
        cv.tag_bind(f"{t}:dir", "<Button-1>", lambda e: g.mover(self.indice, +1))
        self._desenhar_nome()

    def _desenhar_nome(self):
        """Nome trocado fica em itálico (e aparece o ↺), pra ver que não é o da Central de Gestão."""
        g = self.galeria
        mudou = self.var_legenda.get().strip() != self.item["legenda"]
        fonte = g.f_nome_it if mudou else g.f_nome
        g.canvas.itemconfigure(self.i_nome, text=_encurtar(self.legenda, fonte, W - 52), font=fonte)
        g.canvas.itemconfigure(self.i_volta, state="normal" if mudou else "hidden")

    def mover_para(self, x, y):
        if (x, y) != (self.x, self.y):
            self.galeria.canvas.move(self.tag, x - self.x, y - self.y)
            self.x, self.y = x, y

    # ------------------------------------------------------------ estado
    @property
    def legenda(self):
        return self.var_legenda.get().strip() or self.item["legenda"]

    @property
    def chave_local(self):
        return self.galeria.chave_por_destino().get(self.var_destino.get())

    def alternar(self):
        self.var_incluir.set(not self.var_incluir.get())
        self._mudou_incluir()

    def _mudou_incluir(self):
        if not self.var_incluir.get():
            self.var_destino.set(DESTINO_INTERNA)
        self.galeria.atualizar()

    def _realce(self, ligado):
        if self._borda == COR_BORDA:  # só no estado normal (laranja/azul têm significado)
            self.galeria.canvas.itemconfigure(self.i_fundo, outline=visual.BORDA_FORTE if ligado else COR_BORDA)

    def pintar(self, posicao, arrastando=False):
        incluida = self.var_incluir.get()
        estado = (incluida, self.var_destino.get(), posicao if self.galeria.na_principal(self) else 0, arrastando,
                  self.galeria.prefixo_local, self.galeria.perfil.quadros_tambem_na_principal)
        if estado == self._estado:
            return  # nada mudou neste cartão: não redesenha
        self._estado = estado
        cv, x, y = self.galeria.canvas, self.x, self.y
        ok = bool(self.item["caminho"])
        if self.imagens:
            cv.itemconfigure(self.i_foto, image=self.imagens[0 if incluida and not arrastando else 1])
        local = self.chave_local
        if not ok:
            texto, tipo = "NÃO BAIXOU", "erro"
        elif arrastando:
            texto, tipo = "MOVENDO…", "movendo"
        elif not incluida:
            texto, tipo = "FORA DO EXCEL", "fora"
        elif local and self.galeria.perfil.quadros_tambem_na_principal:  # Banco D: quadro + aba principal
            texto, tipo = f"FOTO {posicao}  +  {self.galeria.prefixo_local.upper()}", "local"
        elif local:
            texto, tipo = f"{self.galeria.prefixo_local.upper()} · {self.galeria.titulo_secao(local).upper()}", "local"
        else:
            texto, tipo = f"FOTO {posicao}", "principal"
        fundo, cor = FAIXA[tipo]
        cv.itemconfigure(self.i_faixa, fill=fundo)
        cv.itemconfigure(self.i_rotulo, text=_encurtar(texto, self.galeria.f_faixa, W - 60), fill=cor)
        self._borda = COR_ALVO if arrastando else (COR_LOCAL if (local and incluida) else COR_BORDA)
        cv.itemconfigure(self.i_fundo, outline=self._borda, width=3 if self._borda != COR_BORDA else 2)
        # interruptor
        ligado = incluida and ok
        cv.itemconfigure(self.i_sw, fill=visual.AZUL if ligado else "#FFFFFF",
                         outline=visual.AZUL if ligado else ("#8A94A3" if ok else "#C9CED6"))
        cx = x + (39 if ligado else 23)
        cv.coords(self.i_bola, cx - 6, y + Y_LINHA - 6, cx + 6, y + Y_LINHA + 6)
        cv.itemconfigure(self.i_bola, fill="white" if ligado else ("#5F6B7A" if ok else "#C9CED6"))
        cv.itemconfigure(self.i_sw_txt, fill=visual.TEXTO if ok else visual.CINZA)
        # seção
        dest = self.var_destino.get()
        cv.itemconfigure(self.i_dest, text=_encurtar(dest, self.galeria.f_txt, W - 62),
                         fill=COR_LOCAL if local else (visual.TEXTO if ok else visual.CINZA))

    def piscar(self, vezes=2):
        """Destaque rápido depois de mover, pra pessoa achar a foto no lugar novo."""
        cv = self.galeria.canvas

        def passo(n):
            if self.galeria.cartoes.get(self.indice) is not self:
                return
            if n % 2 == 0:
                cv.itemconfigure(self.i_fundo, outline=COR_ALVO, width=3)
            else:
                self._estado = None  # força repintar a borda normal
                self.galeria.atualizar()
            if n < vezes * 2 - 1:
                cv.after(160, lambda: passo(n + 1))
        passo(0)


class Visualizador(tk.Toplevel):
    """Foto grande dentro do programa, com anterior/próxima e "Incluir"."""

    def __init__(self, galeria, indice):
        super().__init__(galeria.winfo_toplevel())
        self.galeria = galeria
        self.title("Foto")
        self.configure(bg="#1E1F22")
        larg, alt = int(self.winfo_screenwidth() * 0.82), int(self.winfo_screenheight() * 0.86)
        self.geometry(f"{larg}x{alt}+{(self.winfo_screenwidth() - larg) // 2}+{(self.winfo_screenheight() - alt) // 3}")
        self.transient(galeria.winfo_toplevel())
        self.area = (larg - 40, alt - 170)  # sobra espaço pra faixa da observação

        topo = tk.Frame(self, bg="#1E1F22")
        topo.pack(fill="x", padx=16, pady=(12, 6))
        self.lbl_titulo = tk.Label(topo, bg="#1E1F22", fg="white", font=("Segoe UI", 14, "bold"), anchor="w")
        self.lbl_titulo.pack(side="left")
        self.lbl_info = tk.Label(topo, bg="#1E1F22", fg="#AAB", font=("Segoe UI", 10))
        self.lbl_info.pack(side="right")
        # observação da foto na Central de Gestão ("Observações:" ao abrir a foto lá); só aparece se tiver
        self.lbl_obs = tk.Label(self, bg=COR_OBS, fg="#1E1F22", font=("Segoe UI", 11, "bold"), anchor="w",
                                justify="left", padx=12, pady=6, wraplength=larg - 60)

        self.lbl_foto = tk.Label(self, bg="#1E1F22", cursor="hand2")
        self.lbl_foto.pack(expand=True)
        self.lbl_foto.bind("<Button-1>", lambda e: self.destroy())

        base = tk.Frame(self, bg="#1E1F22")
        base.pack(fill="x", padx=16, pady=(6, 12))
        # takefocus=0: o teclado fica sempre na janela (senão o Espaço com foco
        # na caixinha alternava duas vezes - caixinha + janela - e não mudava nada)
        ttk.Button(base, text="◀  Anterior", command=lambda: self.ir(-1), takefocus=0).pack(side="left")
        ttk.Button(base, text="Próxima  ▶", command=lambda: self.ir(+1), takefocus=0).pack(side="left", padx=8)
        self.var_incluir = tk.BooleanVar()
        tk.Checkbutton(base, text="Incluir no Excel", variable=self.var_incluir, command=self._incluir,
                       bg="#1E1F22", fg="white", selectcolor="#1E1F22", activebackground="#1E1F22",
                       activeforeground="white", font=("Segoe UI", 11), takefocus=0).pack(side="left", padx=20)
        tk.Label(base, text="← → navega  •  Espaço inclui/tira  •  Esc ou clique fecha", bg="#1E1F22",
                 fg="#889", font=("Segoe UI", 9)).pack(side="left", padx=10)
        ttk.Button(base, text="Fechar", command=self.destroy, takefocus=0).pack(side="right")

        self.bind("<Left>", lambda e: self.ir(-1))
        self.bind("<Right>", lambda e: self.ir(+1))
        self.bind("<space>", self._espaco)
        self.bind("<Escape>", lambda e: self.destroy())
        self.mostrar(indice)
        self.after(60, self.focus_force)  # garante que as teclas vão pra esta janela

    def _espaco(self, evento):
        if isinstance(evento.widget, (tk.Checkbutton, tk.Button, ttk.Button)):
            return  # o próprio botão já tratou o Espaço
        self.var_incluir.set(not self.var_incluir.get())
        self._incluir()

    def _navegaveis(self):
        return [i for i in self.galeria.ordem if self.galeria.cartoes[i].item["caminho"]]

    def mostrar(self, indice):
        self.indice = indice
        c = self.galeria.cartoes[indice]
        try:
            with Image.open(c.item["caminho"]) as img:
                img.draft("RGB", self.area)
                img = ImageOps.exif_transpose(img).convert("RGB")
                # ocupa a janela toda (amplia foto pequena também, ex.: 960x720)
                escala = min(self.area[0] / img.width, self.area[1] / img.height)
                img = img.resize((max(1, int(img.width * escala)), max(1, int(img.height * escala))), Image.LANCZOS)
            self._foto = ImageTk.PhotoImage(img)
            self.lbl_foto.config(image=self._foto, text="")
        except Exception as e:
            self.lbl_foto.config(image="", text=f"Não consegui abrir a foto: {e}", fg="white")
        lista = self._navegaveis()
        self.lbl_titulo.config(text=c.legenda)
        origem = "do computador" if c.item.get("origem") == "pc" else f"Central de Gestão #{indice + 1}"
        self.lbl_info.config(text=f"{lista.index(indice) + 1} de {len(lista)}  •  {origem}")
        obs = c.item.get("observacao")
        if obs:
            self.lbl_obs.config(text=f"Observação na Central de Gestão:  {obs}")
            self.lbl_obs.pack(fill="x", padx=16, pady=(0, 6), before=self.lbl_foto)
        else:
            self.lbl_obs.pack_forget()
        self.var_incluir.set(c.var_incluir.get())
        self.title(f"Foto - {c.legenda}")

    def ir(self, passo):
        lista = self._navegaveis()
        self.mostrar(lista[(lista.index(self.indice) + passo) % len(lista)])

    def _incluir(self):
        c = self.galeria.cartoes[self.indice]
        c.var_incluir.set(self.var_incluir.get())
        c._mudou_incluir()


class Galeria(ttk.Frame):
    def __init__(self, pai, ao_mudar=None, secoes_extras=None, ao_mudar_secoes=None, perfil=BANCO_C):
        super().__init__(pai)
        self.perfil = perfil
        self.ao_mudar = ao_mudar or (lambda: None)
        self.ao_mudar_secoes = ao_mudar_secoes or (lambda lista: None)
        self.secoes_extras = list(secoes_extras or [])
        self.cartoes = {}
        self.ordem = []
        self.colunas = 0
        self._margem = 10
        self._arrasto = None
        self._animacao = None
        self._alvo_rolagem = None
        self._editor = None

        self.f_faixa = tkfont.Font(family="Segoe UI", size=8, weight="bold")
        self.f_num = tkfont.Font(family="Segoe UI", size=8)
        self.f_nome = tkfont.Font(family="Segoe UI", size=9, weight="bold")
        self.f_nome_it = tkfont.Font(family="Segoe UI", size=9, weight="bold", slant="italic")
        self.f_txt = tkfont.Font(family="Segoe UI", size=9)
        self.f_btn = tkfont.Font(family="Segoe UI", size=11)
        self.f_seta = tkfont.Font(family="Segoe UI", size=9)

        self.canvas = tk.Canvas(self, bg=COR_FUNDO, highlightthickness=0, yscrollincrement=1)
        barra = ttk.Scrollbar(self, orient="vertical", command=self._barra)
        self.canvas.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", self._redimensionou)
        # roda do mouse: vale em qualquer ponto da galeria (inclusive em cima dos cartões)
        self.bind_all("<MouseWheel>", self._roda, add="+")
        # tela vazia (desenhada também)
        self.canvas.create_text(0, 0, text="🖼", font=("Segoe UI Emoji", 44), fill="#B5BEC9", tags=("vazio", "v1"))
        self.canvas.create_text(0, 0, text="Nenhuma proposta carregada", font=("Segoe UI", 14, "bold"),
                                fill=visual.TEXTO, tags=("vazio", "v2"))
        self.canvas.create_text(0, 0, justify="center", font=("Segoe UI", 10), fill=visual.TEXTO_2,
                                text="Escolha o banco, digite o nº da proposta e clique em “Buscar fotos”.\n"
                                     "As fotos aparecem aqui, na ordem da Central de Gestão, para você escolher,\n"
                                     "ordenar e renomear antes de gerar o Excel.", tags=("vazio", "v3"))
        # barra azul que mostra onde a foto arrastada vai cair
        self.marcador = self.canvas.create_rectangle(0, 0, 0, 0, fill=COR_ALVO, width=0, state="hidden")

    # ------------------------------------------------------------ seções
    def secoes(self):
        """[(chave, título)] na ordem em que os quadros entram no Excel."""
        if not self.perfil.locais:  # banco sem aba de fotos locais: só a aba principal
            return []
        extras = self.secoes_extras if self.perfil.locais.aceita_secoes_novas else []
        return list(self.perfil.rotulos_locais().items()) + [(PREFIXO_EXTRA + n, n) for n in extras]

    @property
    def prefixo_local(self):
        return self.perfil.locais.prefixo if self.perfil.locais else "Local"

    def titulo_secao(self, chave):
        return dict(self.secoes()).get(chave, chave)

    def opcoes_destino(self):
        if not self.perfil.locais:
            return [DESTINO_INTERNA]
        nova = [NOVA_SECAO] if self.perfil.locais.aceita_secoes_novas else []
        return [DESTINO_INTERNA] + [f"{self.prefixo_local} · {t}" for _, t in self.secoes()] + nova

    def definir_perfil(self, perfil, secoes_extras):
        """Troca de banco: seções do novo banco; foto que estava num quadro
        que não existe mais volta pra aba principal."""
        self.perfil = perfil
        self.secoes_extras = list(secoes_extras or [])
        validos = set(self.opcoes_destino())
        for c in self.cartoes.values():
            if c.var_destino.get() not in validos:
                c.var_destino.set(DESTINO_INTERNA)
            c._destino_anterior = c.var_destino.get()
        self.atualizar()

    def chave_por_destino(self):
        return {f"{self.prefixo_local} · {t}": k for k, t in self.secoes()}

    def destino_da_chave(self, chave):
        return f"{self.prefixo_local} · {self.titulo_secao(chave)}"

    def _atualizar_opcoes(self):
        pass  # o menu de seções é montado na hora do clique, sempre atualizado

    def nova_secao(self, pai=None):
        """Pergunta o nome, guarda como padrão e devolve a chave (ou None)."""
        nome = simpledialog.askstring(
            "Nova seção", f"Nome da nova seção.\n\nVira um quadro novo na aba “{self.perfil.locais.nome}”,\n"
                          "com esse nome embaixo (ex.: Fundos do imóvel):", parent=pai or self)
        nome = (nome or "").strip()
        if not nome:
            return None
        existentes = {t.lower() for _, t in self.secoes()}
        if nome.lower() in existentes:
            messagebox.showinfo("Nova seção", f"A seção “{nome}” já existe.", parent=pai or self)
            return next(k for k, t in self.secoes() if t.lower() == nome.lower())
        self.secoes_extras.append(nome)
        self.ao_mudar_secoes(self.secoes_extras)
        self._atualizar_opcoes()
        return PREFIXO_EXTRA + nome

    def remover_secoes(self, nomes):
        for nome in nomes:
            if nome in self.secoes_extras:
                self.secoes_extras.remove(nome)
                for c in self.cartoes.values():
                    if c.var_destino.get() == f"{self.prefixo_local} · {nome}":
                        c.var_destino.set(DESTINO_INTERNA)
        self.ao_mudar_secoes(self.secoes_extras)
        self._atualizar_opcoes()
        self.atualizar()

    # ------------------------------------------------------------ dados
    def carregar(self, itens, miniaturas, locais_sugeridos):
        self.limpar()
        self.canvas.itemconfigure("vazio", state="hidden")
        self._fotos_tk = {}
        self._calcular_colunas(self.canvas.winfo_width())
        for pos, it in enumerate(itens):
            par = miniaturas.get(it["indice"])
            imagens = (ImageTk.PhotoImage(par[0]), ImageTk.PhotoImage(par[1])) if par else None
            self._fotos_tk[it["indice"]] = imagens
            x, y = self._posicao(pos)
            self.cartoes[it["indice"]] = Cartao(self, it, imagens, x, y)
        self.ordem = [it["indice"] for it in itens]
        for chave, idx in locais_sugeridos.items():
            self.cartoes[idx].var_destino.set(self.destino_da_chave(chave))
        self.canvas.tag_raise(self.marcador)
        self._reorganizar()
        self.canvas.yview_moveto(0)

    def adicionar(self, itens, miniaturas):
        """Fotos novas (do computador) no fim da galeria, já incluídas na aba
        principal; rola até elas e pisca, pra pessoa ver onde entraram."""
        for it in itens:
            par = miniaturas.get(it["indice"])
            imagens = (ImageTk.PhotoImage(par[0]), ImageTk.PhotoImage(par[1])) if par else None
            self._fotos_tk[it["indice"]] = imagens
            x, y = self._posicao(len(self.ordem))
            self.cartoes[it["indice"]] = Cartao(self, it, imagens, x, y)
            self.ordem.append(it["indice"])
        if not itens:
            return
        self.canvas.tag_raise(self.marcador)
        self._reorganizar()
        self.atualizar()
        self._mostrar_cartao(itens[-1]["indice"])
        for it in itens:
            self.cartoes[it["indice"]].piscar()

    def limpar(self):
        self._fechar_editor(salvar=False)
        self.canvas.delete("cartao")
        self.cartoes, self.ordem = {}, []
        self.canvas.itemconfigure(self.marcador, state="hidden")
        self.canvas.itemconfigure("vazio", state="normal")
        self._centralizar_vazio()
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), self.canvas.winfo_height()))
        self.ao_mudar()

    def na_principal(self, cartao):
        """A foto (incluída) vai pra aba principal? Banco C: só se não estiver num
        quadro; Banco D: sempre (o quadro da PAG.1 é uma cópia)."""
        return cartao.var_incluir.get() and (not cartao.chave_local or self.perfil.quadros_tambem_na_principal)

    def legendas(self):
        """{indice: legenda (a digitada, se a pessoa trocou)}."""
        return {idx: c.legenda for idx, c in self.cartoes.items()}

    def selecao(self):
        """(internas na ordem, {chave da seção: indice} na ordem das seções).
        No Banco D a foto de um quadro entra também nas internas."""
        internas, por_chave = [], {}
        for idx in self.ordem:
            c = self.cartoes[idx]
            if not c.var_incluir.get():
                continue
            if c.chave_local:
                por_chave[c.chave_local] = idx
            if self.na_principal(c):
                internas.append(idx)
        locais = {k: por_chave[k] for k, _ in self.secoes() if k in por_chave}
        return internas, locais

    # ------------------------------------------------------------ trabalho salvo (trabalho.py)
    def _chave(self, idx):
        return self.cartoes[idx].item.get("chave") or f"#{idx}"

    def estado(self):
        """O que a pessoa fez, por chave da foto (não pela posição): ordem,
        fotos fora do Excel, nomes trocados e quadro/seção de cada uma."""
        cartoes = [(self._chave(i), self.cartoes[i]) for i in self.ordem]
        return {
            "ordem": [k for k, _ in cartoes],
            "fora": [k for k, c in cartoes if c.item["caminho"] and not c.var_incluir.get()],
            "nomes": {k: c.legenda for k, c in cartoes if c.legenda != c.item["legenda"]},
            "destinos": {k: c.chave_local for k, c in cartoes if c.var_incluir.get() and c.chave_local},
        }

    def aplicar_estado(self, estado):
        """Volta um estado salvo. Foto que não existe mais é ignorada; foto
        nova (não estava no estado) fica no fim, incluída. Seção nova que não
        existe mais neste PC é criada de novo. Devolve (retomadas, sumidas, novas)."""
        por_chave = {c.item.get("chave"): idx for idx, c in self.cartoes.items()}
        salvas = [por_chave[k] for k in estado.get("ordem", []) if k in por_chave]
        conhecidas = set(salvas)
        novas = [i for i in self.ordem if i not in conhecidas]
        self.ordem = salvas + novas
        fora, nomes = set(estado.get("fora", [])), estado.get("nomes", {})
        for idx in salvas:
            c, k = self.cartoes[idx], self._chave(idx)
            if c.item["caminho"]:
                c.var_incluir.set(k not in fora)
            c.var_legenda.set(nomes.get(k, c.item["legenda"]))
            c.var_destino.set(DESTINO_INTERNA)
        for k, chave_local in estado.get("destinos", {}).items():
            idx = por_chave.get(k)
            if idx is None or not self.perfil.locais:
                continue
            if chave_local.startswith(PREFIXO_EXTRA) and self.perfil.locais.aceita_secoes_novas:
                nome = chave_local[len(PREFIXO_EXTRA):]
                if nome not in self.secoes_extras:
                    self.secoes_extras.append(nome)
                    self.ao_mudar_secoes(self.secoes_extras)
            if chave_local in dict(self.secoes()):
                self.cartoes[idx].var_destino.set(self.destino_da_chave(chave_local))
        for c in self.cartoes.values():
            c._destino_anterior = c.var_destino.get()
        self._reorganizar()
        self.canvas.yview_moveto(0)
        return len(salvas), len(set(estado.get("ordem", [])) - set(por_chave)), len(novas)

    def marcar_todas(self, valor):
        for c in self.cartoes.values():
            if c.item["caminho"]:
                c.var_incluir.set(valor)
                if not valor:
                    c.var_destino.set(DESTINO_INTERNA)
        self.atualizar()

    def ordem_original(self):
        self.ordem.sort()
        self._reorganizar(animar=True)

    def aplicar_sugestao(self, locais_sugeridos):
        for c in self.cartoes.values():
            if c.chave_local in self.perfil.rotulos_locais():
                c.var_destino.set(DESTINO_INTERNA)
        for chave, idx in locais_sugeridos.items():
            self.cartoes[idx].var_incluir.set(True)
            self.cartoes[idx].var_destino.set(self.destino_da_chave(chave))
        self.atualizar()

    def destino_mudou(self, indice):
        c = self.cartoes[indice]
        if c.var_destino.get() == NOVA_SECAO:
            chave = self.nova_secao(pai=self.winfo_toplevel())
            c.var_destino.set(self.destino_da_chave(chave) if chave else c._destino_anterior)
        if c.chave_local:  # cada quadro aceita uma foto só
            c.var_incluir.set(True)
            for outro in self.cartoes.values():
                if outro is not c and outro.chave_local == c.chave_local:
                    outro.var_destino.set(DESTINO_INTERNA)
        c._destino_anterior = c.var_destino.get()
        self.atualizar()

    def mover(self, indice, passo):
        pos = self.ordem.index(indice)
        nova = max(0, min(len(self.ordem) - 1, pos + passo))
        if nova != pos:
            self.ordem.insert(nova, self.ordem.pop(pos))
            self._reorganizar(animar=True)
            self._mostrar_cartao(indice)
            self.cartoes[indice].piscar()

    def abrir_foto(self, indice):
        Visualizador(self, indice)

    # ------------------------------------------------------------ nome e seção (sob demanda)
    def editar_nome(self, cartao):
        """Campo de texto que aparece só na hora de renomear (em cima do nome)."""
        self._fechar_editor(salvar=True)
        e = ttk.Entry(self.canvas, font=("Segoe UI", 9, "bold"), justify="center")
        e.insert(0, cartao.var_legenda.get())
        e.select_range(0, "end")
        janela = self.canvas.create_window(cartao.x + 10, cartao.y + Y_NOME - 2, anchor="nw", window=e,
                                           width=W - 20, height=28, tags=(cartao.tag, "cartao"))
        self._editor = (cartao, e, janela)
        e.bind("<Return>", lambda ev: self._fechar_editor(salvar=True))
        e.bind("<Escape>", lambda ev: self._fechar_editor(salvar=False))
        e.bind("<FocusOut>", lambda ev: self._fechar_editor(salvar=True))
        e.focus_set()

    def _fechar_editor(self, salvar):
        if not self._editor:
            return
        cartao, e, janela = self._editor
        self._editor = None
        if salvar:
            cartao.var_legenda.set(e.get().strip() or cartao.item["legenda"])
        self.canvas.delete(janela)
        e.destroy()

    def menu_destino(self, cartao, evento):
        m = tk.Menu(self.canvas, tearoff=0, font=("Segoe UI", 10))
        for op in self.opcoes_destino():
            m.add_radiobutton(label=op, value=op, variable=cartao.var_destino,
                              command=lambda: self.destino_mudou(cartao.indice))
        try:
            m.tk_popup(evento.x_root, evento.y_root)
        finally:
            m.grab_release()

    def passou_mouse(self, cartao, zona, dentro):
        """Mãozinha e fundo nos botões desenhados."""
        self.canvas.config(cursor="hand2" if dentro else "")
        cv, t = self.canvas, cartao.tag
        if zona in ("esq", "dir"):
            cv.itemconfigure(f"{t}:{zona}bg", fill=COR_HOVER if dentro else COR_CARTAO)
        elif zona == "dest":
            cv.itemconfigure(f"{t}:destbg", outline=visual.BORDA_FORTE if dentro else COR_BORDA)

    # ------------------------------------------------------------ layout
    def _calcular_colunas(self, largura):
        colunas = max(1, (largura - 20 + GAP) // (W + GAP))
        margem = max(10, (largura - (colunas * W + (colunas - 1) * GAP)) // 2)  # grade centralizada
        mudou = (colunas, margem) != (self.colunas, self._margem)
        self.colunas, self._margem = colunas, margem
        return mudou

    def _posicao(self, pos):
        colunas = self.colunas or 4
        return self._margem + (pos % colunas) * (W + GAP), 12 + (pos // colunas) * (H + GAP)

    def _redimensionou(self, evento):
        self._centralizar_vazio()
        if self._calcular_colunas(evento.width) and self.ordem:
            self._reorganizar()

    def _centralizar_vazio(self):
        cx = max(self.canvas.winfo_width(), 200) / 2
        for tag, y in (("v1", 90), ("v2", 160), ("v3", 210)):
            self.canvas.coords(tag, cx, y)

    def _reorganizar(self, animar=False):
        """Põe cada cartão na sua casa. Com `animar`, os que mudaram de lugar
        deslizam até lá."""
        if self._animacao:
            self.after_cancel(self._animacao)
            self._animacao = None
        if not self.ordem:
            return
        colunas = self.colunas or 4
        linhas = -(-len(self.ordem) // colunas)
        altura = 24 + linhas * (H + GAP)
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(), altura))
        moves = []
        for pos, idx in enumerate(self.ordem):
            c = self.cartoes[idx]
            x1, y1 = self._posicao(pos)
            if (c.x, c.y) == (x1, y1):
                continue
            if animar:
                moves.append((c, c.x, c.y, x1, y1))
            else:
                c.mover_para(x1, y1)
        if moves:
            self._deslizar(moves, 1)
        self.atualizar()

    def _deslizar(self, moves, quadro):
        t = quadro / ANIMACAO_QUADROS
        f = 1 - (1 - t) ** 3  # desacelera no fim (ease-out)
        for c, x0, y0, x1, y1 in moves:
            if self.cartoes.get(c.indice) is c:
                c.mover_para(round(x0 + (x1 - x0) * f), round(y0 + (y1 - y0) * f))
        if quadro < ANIMACAO_QUADROS:
            self._animacao = self.after(ANIMACAO_MS // ANIMACAO_QUADROS,
                                        lambda: self._deslizar(moves, quadro + 1))
        else:
            self._animacao = None

    def atualizar(self):
        posicao = 0
        arrastado = self._arrasto["indice"] if self._arrasto and self._arrasto["ativo"] else None
        for idx in self.ordem:
            c = self.cartoes[idx]
            if self.na_principal(c):
                posicao += 1
            c.pintar(posicao, arrastando=(idx == arrastado))
        self.ao_mudar()

    # ------------------------------------------------------------ rolagem suave
    def _altura_total(self):
        try:
            return float(self.canvas.cget("scrollregion").split()[3])
        except (IndexError, ValueError):
            return float(self.canvas.winfo_height())

    def _barra(self, *args):
        self._alvo_rolagem = None  # a barra de rolagem manda direto
        self.canvas.yview(*args)

    def rolar(self, evento):
        self._rolar_suave(-evento.delta / 120 * 120)

    def _rolar_suave(self, dy):
        total, visivel = self._altura_total(), self.canvas.winfo_height()
        if total <= visivel:
            return
        atual = self.canvas.canvasy(0)
        base = self._alvo_rolagem if self._alvo_rolagem is not None else atual
        animando = self._alvo_rolagem is not None
        self._alvo_rolagem = max(0.0, min(total - visivel, base + dy))
        if not animando:
            self._passo_rolagem()

    def _passo_rolagem(self):
        if self._alvo_rolagem is None:
            return
        total = self._altura_total()
        atual = self.canvas.canvasy(0)
        falta = self._alvo_rolagem - atual
        if abs(falta) < 1.5:
            self.canvas.yview_moveto(self._alvo_rolagem / total)
            self._alvo_rolagem = None
            return
        passo = falta * 0.35
        if abs(passo) < 1:
            passo = 1 if falta > 0 else -1
        self.canvas.yview_moveto((atual + passo) / total)
        self.after(12, self._passo_rolagem)

    def _roda(self, evento):
        for w in (evento.widget, self.winfo_containing(evento.x_root, evento.y_root)):
            while isinstance(w, tk.Misc) and w is not self:
                w = w.master
            if w is self:
                self.rolar(evento)
                return

    def _mostrar_cartao(self, indice):
        """Rola a galeria até o cartão aparecer inteiro (pela casa de destino,
        não pela posição atual - ele pode estar deslizando)."""
        total, visivel = self._altura_total(), self.canvas.winfo_height()
        topo = self._posicao(self.ordem.index(indice))[1]
        atual = self.canvas.canvasy(0)
        if topo - 8 < atual:
            self._rolar_suave(topo - 8 - atual)
        elif topo + H + 8 > atual + visivel:
            self._rolar_suave(topo + H + 8 - atual - visivel)

    # ------------------------------------------------------------ arrastar
    def _posicao_de_soltar(self, cx, cy, arrastado):
        """(cartão de referência, antes?) - a casa mais perto do ponteiro
        (em coordenadas do Canvas), sem contar o próprio cartão arrastado."""
        melhor, dist = None, None
        for pos, idx in enumerate(self.ordem):
            if idx == arrastado:
                continue
            x, y = self._posicao(pos)
            d = ((x + W / 2 - cx) / W) ** 2 + ((y + H / 2 - cy) / H) ** 2
            if dist is None or d < dist:
                melhor, dist = self.cartoes[idx], d
        if melhor is None:
            return None, True
        x, _ = self._posicao(self.ordem.index(melhor.indice))
        return melhor, cx < x + W / 2

    def inicio_arrasto(self, indice, evento):
        c = self.cartoes[indice]
        cx, cy = self.canvas.canvasx(evento.x), self.canvas.canvasy(evento.y)
        self._arrasto = {"indice": indice, "x": evento.x_root, "y": evento.y_root, "ativo": False,
                         "dx": cx - c.x, "dy": cy - c.y, "ref": None, "antes": True, "rolagem": None,
                         "ultimo": (evento.x, evento.y)}

    def movendo(self, evento):
        a = self._arrasto
        if not a:
            return
        if not a["ativo"]:
            if abs(evento.x_root - a["x"]) + abs(evento.y_root - a["y"]) <= 8:
                return
            a["ativo"] = True
            self.canvas.config(cursor="fleur")
            self.canvas.tag_raise(self.cartoes[a["indice"]].tag)
            self.atualizar()
            self._rolar_sozinho()
        a["ultimo"] = (evento.x, evento.y)
        self._acompanhar(evento.x, evento.y)

    def _acompanhar(self, x, y):
        """O cartão arrastado segue o mouse; a barra azul mostra onde cai."""
        a = self._arrasto
        c = self.cartoes[a["indice"]]
        cx, cy = self.canvas.canvasx(x), self.canvas.canvasy(y)
        c.mover_para(cx - a["dx"], cy - a["dy"])
        ref, antes = self._posicao_de_soltar(cx, cy, a["indice"])
        a["ref"], a["antes"] = ref, antes
        if ref is None:
            self.canvas.itemconfigure(self.marcador, state="hidden")
            return
        rx, ry = self._posicao(self.ordem.index(ref.indice))
        mx = rx - GAP / 2 - 3 if antes else rx + W + GAP / 2 - 3
        self.canvas.coords(self.marcador, mx, ry, mx + 6, ry + H)
        self.canvas.itemconfigure(self.marcador, state="normal")
        self.canvas.tag_raise(self.marcador)
        self.canvas.tag_raise(c.tag)

    def _rolar_sozinho(self):
        """Enquanto arrasta, rola a lista se o mouse estiver perto da borda."""
        a = self._arrasto
        if not a or not a["ativo"]:
            return
        y = self.winfo_pointery() - self.canvas.winfo_rooty()
        altura, zona = self.canvas.winfo_height(), 70
        dy = 0
        if y < zona:
            dy = -max(4, (zona - y) // 3)
        elif y > altura - zona:
            dy = max(4, (y - altura + zona) // 3)
        if dy:
            total = self._altura_total()
            atual = self.canvas.canvasy(0)
            novo = max(0.0, min(total - altura, atual + dy))
            if novo != atual:
                self.canvas.yview_moveto(novo / total)
                self._acompanhar(self.winfo_pointerx() - self.canvas.winfo_rootx(),
                                 self.winfo_pointery() - self.canvas.winfo_rooty())
        a["rolagem"] = self.after(16, self._rolar_sozinho)

    def fim_arrasto(self, evento):
        a, self._arrasto = self._arrasto, None
        self.canvas.config(cursor="")
        self.canvas.itemconfigure(self.marcador, state="hidden")
        if not a:
            return
        if a["rolagem"]:
            self.after_cancel(a["rolagem"])
        if not a["ativo"]:  # foi só um clique: abre a foto grande
            self.abrir_foto(a["indice"])
            return
        ref = a["ref"]
        if ref is not None and ref.indice != a["indice"]:
            self.ordem.remove(a["indice"])
            destino = self.ordem.index(ref.indice) + (0 if a["antes"] else 1)
            self.ordem.insert(destino, a["indice"])
        self.cartoes[a["indice"]]._estado = None
        self._reorganizar(animar=True)  # o arrastado desliza do ponto onde foi solto
        self._mostrar_cartao(a["indice"])
        self.cartoes[a["indice"]].piscar()
