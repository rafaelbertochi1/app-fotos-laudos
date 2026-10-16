"""
Perfis por banco: tudo o que muda de um banco pro outro (modelos, abas,
layout dos blocos de foto, quadros da aba local, nome do arquivo). O resto
do programa (busca na Central de Gestão, galeria, login) é igual pra todos.

Pra um banco novo: criar a pasta modelos/<pasta_modelos>/ com os modelos
(.xlsm ou .xlsx) e um Perfil aqui embaixo, registrado em PERFIS. Os números
de linha/coluna saem do próprio modelo (ver como foram medidos os do Banco C e
do Banco D). Exemplos preenchidos ficam em referencia/<banco>/ (fora do
.exe e do Git: têm fotos de cliente).
"""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Grade:
    """Pra modelo que NÃO traz os blocos de foto prontos (ex.: Banco D): o
    programa monta a grade na hora - blocos mesclados com borda fina e a
    legenda mesclada embaixo, só pros quadros que forem usados."""
    altura_linha_foto: float    # pt, cada uma das linhas do bloco da foto
    altura_espaco: float        # pt, linha entre a foto e a legenda
    altura_legenda: float       # pt
    altura_separador: float     # pt, linha depois da legenda
    limpar: str = ""            # faixa com texto de instrução do modelo a apagar (ex.: "A8:L14")
    colunas_iguais: str = ""    # deixa estas colunas com a mesma largura (as duas fotos do mesmo tamanho)


@dataclass(frozen=True)
class AbaFotos:
    """Aba principal: páginas iguais, com blocos mesclados de foto e o texto
    da descrição logo abaixo de cada bloco."""
    nome: str
    linhas_por_pagina: int
    linhas_foto: tuple        # 1ª linha de cada fileira de fotos, dentro da página
    colunas: tuple            # por foto da fileira: (coluna inicial, coluna final, coluna do texto da descrição)
    altura_bloco: int         # linhas do bloco mesclado da foto
    paginas_no_modelo: int    # passou disso, cria páginas novas copiando a página molde
    pagina_molde: int         # índice (0 = 1ª) da página copiada
    deslocamento_pagina: int  # página k>0 começa na linha k * linhas_por_pagina + deslocamento
    col_fim_impressao: str = "O"
    linha_descricao: int = 0  # descrição na linha 1ª + isto (0 = logo abaixo do bloco)
    grade: Grade = None       # modelo sem blocos prontos: monta a grade (ver Grade)

    @property
    def fotos_por_pagina(self):
        return len(self.linhas_foto) * len(self.colunas)

    @property
    def deslocamento_descricao(self):
        return self.linha_descricao or self.altura_bloco


@dataclass(frozen=True)
class Quadro:
    chave: str
    rotulo: str        # como aparece na janela ("Fachada")
    bloco: str         # bloco mesclado no modelo ("B31:G43")
    titulo: str        # título que o modelo traz (embaixo do quadro ou dentro dele)
    sugestao: str = ""  # regex na legenda da Central de Gestão pro botão "Sugerir"
    limpar_texto: bool = False  # o título fica DENTRO do bloco: apaga quando a foto entra


@dataclass(frozen=True)
class PaginasExtras:
    """Seções novas numa aba de quadros que não tem bloco pra copiar (ex.:
    PAG.1 do Banco D): páginas novas depois da área de impressão do modelo,
    com um título e quadros de 2 em 2 (foto mesclada + nome da seção embaixo)."""
    linha_inicio: int             # 1ª linha da 1ª página nova
    titulo: str
    colunas: tuple                # por quadro da fileira: (coluna inicial, coluna final)
    fileiras: int                 # fileiras por página
    linhas_quadro: int            # linhas do bloco da foto
    altura_linha_foto: float      # pt
    altura_espaco: float          # pt, entre foto e nome
    altura_legenda: float         # pt
    altura_separador: float       # pt, depois do nome
    altura_titulo: float          # pt
    col_fim: str                  # última coluna (título e área de impressão)

    @property
    def linhas_fileira(self):
        return self.linhas_quadro + 3  # foto + espaço + nome + separador

    @property
    def linhas_por_pagina(self):
        return 2 + self.fileiras * self.linhas_fileira  # título + espaço + fileiras

    @property
    def quadros_por_pagina(self):
        return self.fileiras * len(self.colunas)


@dataclass(frozen=True)
class AbaLocais:
    """Aba com quadros fixos (fachada etc.). Seções novas criadas na janela:
    com bloco molde, viram cópias dele (cabeçalho + quadros) uma embaixo da
    outra (Banco C); com paginas_extras, páginas novas montadas na hora
    (Banco D); sem nenhum dos dois, não dá pra criar seção nova."""
    nome: str
    quadros: tuple                # Quadro, na ordem
    bloco_molde: tuple = None     # (linha inicial, linha final) do bloco copiado
    quadros_no_bloco: tuple = ()  # (coluna inicial, coluna final, linha relativa ao início do bloco)
    altura_quadro: int = 13       # o título fica na linha seguinte ao quadro
    col_fim_impressao: str = "O"
    prefixo: str = "Local"        # como o destino aparece na janela: "Local · Fachada"
    paginas_extras: PaginasExtras = None

    @property
    def aceita_secoes_novas(self):
        return bool(self.bloco_molde or self.paginas_extras)


@dataclass(frozen=True)
class Assinatura:
    """Assinatura da engenheira responsável (assets/assinatura_responsavel.png)
    posta centralizada na área mesclada `area` da aba `aba`."""
    aba: str
    area: str


@dataclass(frozen=True)
class PrePreenchimento:
    """Formulário do banco preenchido com os dados da vistoria (ex.:
    FORMULARIO do Banco D). `mapa` = função de formulario.MAPAS que diz o
    valor de cada célula. Duas edições do .exe (comum.PREENCHIMENTO):
    "basico" = só as `basicas` (o que a planilha do banco já chega
    preenchida); "completo" = tudo que a Central de Gestão tiver. `limpar` =
    células que o modelo traz com dados de outro caso e que a Central de Gestão
    não tem (ficam em branco pra pessoa preencher)."""
    aba: str
    mapa: str
    basicas: tuple
    limpar: tuple = ()


@dataclass(frozen=True)
class Perfil:
    banco: str
    pasta_modelos: str            # modelos/<pasta_modelos>/*.xlsm|*.xlsx
    nome_arquivo: str             # sem extensão (vem do modelo); campos: {proposta}, {tipo}
    fotos: AbaFotos
    locais: AbaLocais = None
    nomes_antigos_modelos: dict = field(default_factory=dict)  # pra config salva por versões antigas
    # foto posta num quadro da aba de quadros entra TAMBÉM na aba principal
    # (na posição dela na tela). Banco C: não (o quadro tira a foto da aba
    # principal); Banco D: sim (pedido do Rafael, 09/10/2026).
    quadros_tambem_na_principal: bool = False
    pre_preenchimento: PrePreenchimento = None
    assinatura: Assinatura = None

    def rotulos_locais(self):
        """{chave: rótulo} dos quadros fixos, na ordem."""
        return {q.chave: q.rotulo for q in self.locais.quadros} if self.locais else {}


# ---------------------------------------------------------------- Banco C
# Medido nos modelos Modelo_V3 (comparativo) e V3_2 (evolutivo), que
# têm o mesmo layout nas duas abas (07-08/10/2026).
BANCO_C = Perfil(
    banco="Banco C",
    pasta_modelos="banco_c",
    nome_arquivo="planilha_{proposta}_{tipo}",
    fotos=AbaFotos(
        nome="Fotos internas imóvel",
        linhas_por_pagina=50,
        linhas_foto=(4, 20, 36),
        colunas=(("B", "G", "C"), ("I", "N", "J")),
        altura_bloco=13,
        paginas_no_modelo=11,
        pagina_molde=1,
        deslocamento_pagina=3,
    ),
    locais=AbaLocais(
        nome="Mapa e fotos locais",
        quadros=(
            Quadro("fachada", "Fachada", "B31:G43", "Fachada Imóvel Avaliando", r"fachada"),
            Quadro("identificacao", "Identificação", "I31:N43", "Identificação Imóvel Avaliando",
                   r"n[º°o]\.?\s*residencial|identifica|placa do im[oó]vel|n[uú]mero"),
            # "Rua (baixo/cima)"; não pega "Placa da Rua"
            Quadro("logradouro_esq", "Logradouro à esquerda", "B47:G59", "Logradouro à Esquerda",
                   r"^\s*rua\b|logradouro"),
            Quadro("logradouro_dir", "Logradouro à direita", "I47:N59", "Logradouro à Direita",
                   r"^\s*rua\b|logradouro"),
        ),
        bloco_molde=(29, 61),
        quadros_no_bloco=(("B", "G", 2), ("I", "N", 2), ("B", "G", 18), ("I", "N", 18)),
    ),
    nomes_antigos_modelos={"Modelo_V3": "Comparativo", "Modelo_V3_2": "Evolutivo"},
)

# ---------------------------------------------------------------- Banco D
# Medido no modelo do Banco D (= "Formulario_Banco_D"; cópia
# como os funcionários baixam em referencia/banco_d/planilha_funcionarios_*) e no
# exemplo preenchido (referencia/banco_d), 08/10/2026.
#  - Modelos (13/10/2026): um por tipo de operação, todos com o mesmo layout
#    - Aquisição simples, Aquisição à vista FGTS, Consórcio urbano, Consórcio
#    reforma sem ampliação (+ bloco de reforma na PAG.1) e Garantia imóvel
#    urbano ("Valor de Liquidez" na PAG.1). O tipo de operação (F36) já vem
#    no modelo. Montados por ferramentas/montar_modelos_banco_d.py.
#  - PAG.1: 3 quadros prontos com o texto dentro - "Foto do imóvel" (A6:F21),
#    "Foto Logradouro" (G6:K16) e "Foto nº imóvel" (G17:J21, o pequeno). Sem
#    identificação do prédio (pedido do Rafael). Seções novas (pedido de
#    09/10/2026: "na mesma aba que as que já estão criadas") viram páginas
#    "FOTOS COMPLEMENTARES" na própria PAG.1, depois da linha 68 (fim da área
#    de impressão). A PAG.1 imprime em escala fixa de 66% (não "ajustar à
#    página"), então as quebras manuais valem e a 1ª página não muda.
#  - FOTOS: o modelo só traz cabeçalho (linhas 1-6) e um texto de instrução
#    (B9:B13). A grade segue o exemplo preenchido (2 fotos por fileira,
#    legenda mesclada embaixo, borda fina), mas dimensionada pra A4 com as
#    margens do modelo: 3 fileiras (6 fotos) por página, 13 linhas por
#    fileira (10 de foto + espaço + legenda + separador).
#  - FORMULARIO: pré-preenchida com os dados da proposta (ver pre_preenchimento).
#  - Diferente do Banco C: foto posta num quadro da PAG.1 (fixo ou seção nova)
#    continua TAMBÉM na FOTOS, na posição dela na tela.
BANCO_D = Perfil(
    banco="Banco D",
    pasta_modelos="banco_d",
    nome_arquivo="{proposta}_FORMULARIO",  # pedido de 09/10/2026 (antes: planilha_{proposta})
    fotos=AbaFotos(
        nome="FOTOS",
        linhas_por_pagina=39,
        linhas_foto=(8, 21, 34),
        colunas=(("A", "F", "A"), ("G", "L", "G")),
        altura_bloco=10,
        linha_descricao=11,
        paginas_no_modelo=0,
        pagina_molde=0,
        deslocamento_pagina=8,
        col_fim_impressao="L",
        grade=Grade(altura_linha_foto=19.5, altura_espaco=6, altura_legenda=18, altura_separador=12,
                    limpar="A7:L14", colunas_iguais="A:L"),
    ),
    locais=AbaLocais(
        nome="PAG.1",
        prefixo="PAG.1",
        quadros=(
            Quadro("imovel", "Foto do imóvel", "A6:F21", "Foto do imóvel", r"fachada", limpar_texto=True),
            Quadro("logradouro", "Foto logradouro", "G6:K16", "Foto Logradouro",
                   r"^\s*rua\b|logradouro", limpar_texto=True),
            Quadro("numero", "Foto nº imóvel", "G17:J21", "Foto nº imóvel",
                   r"n[º°o]\.?\s*residencial|n[uú]mero", limpar_texto=True),
        ),
        # páginas de 50 linhas: título + espaço + 3 fileiras de 16 (13 de foto + espaço + nome + separador);
        # 920 pt de altura x 66% = 607 pt, cabe numa folha A4 (728 pt úteis)
        paginas_extras=PaginasExtras(
            linha_inicio=69, titulo="FOTOS COMPLEMENTARES", colunas=(("A", "E"), ("G", "K")), fileiras=3,
            linhas_quadro=13, altura_linha_foto=19.5, altura_espaco=6, altura_legenda=20, altura_separador=14,
            altura_titulo=30, col_fim="K"),
    ),
    quadros_tambem_na_principal=True,  # as fotos da PAG.1 também ficam na FOTOS
    # básicas = o que a planilha que os funcionários recebem já traz
    # preenchido (protocolo, proponente, bairro, cidade, CEP, UF, data de
    # recebimento, matrícula). A "Empresa avaliadora" é fixa e já está no
    # modelo. Junção e nome da junção também vêm preenchidos lá, mas o
    # Central de Gestão não tem: ficam em branco. O tipo de operação vem do modelo
    # escolhido (não é apagado). Edição "completo": todo o resto que o
    # Central de Gestão tiver (formulario.py).
    pre_preenchimento=PrePreenchimento(
        aba="FORMULARIO",
        mapa="banco_d",
        # matrícula em J34 ("Nº Matrícula"), não F34 ("CNS Cartório") - 13/10/2026
        # F5/F6 (junção) vêm da planilha do Banco D na Intranet, quando ela é achada
        basicas=("F5", "F6", "F7", "F11", "F15", "F16", "F17", "F18", "F19", "J34"),
        limpar=("F5", "F6"),
    ),
    # assinatura da engenheira responsável no quadro ao lado de "Engº./Crea" (F67:G68) - o mesmo
    # lugar onde está no laudo pronto de exemplo (pedido de 09/10/2026)
    assinatura=Assinatura(aba="PAG.1", area="F67:G68"),
    # até a 2.1 só havia um modelo, chamado "Banco D" (config salva nos PCs)
    nomes_antigos_modelos={"Banco D": "Aquisição simples"},
)

PERFIS = {p.banco: p for p in (BANCO_C, BANCO_D)}
PADRAO = BANCO_C.banco
