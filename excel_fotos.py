"""
Redimensiona as fotos e insere no modelo Excel (.xlsm) do banco: na aba
principal (todas as escolhidas) e na aba de fotos locais (quadros). O
layout de cada banco vem do perfil (perfis.py).

Usa o próprio Microsoft Excel por automação (pywin32), e não openpyxl: o
openpyxl regrava o arquivo inteiro e descarta o que não entende (validações
de dados estendidas, controles, partes das macros). Pelo Excel, o modelo
sai idêntico, só com as fotos a mais.

Aba principal (ex.: Banco C, "Fotos internas imóvel"):
  - páginas de 50 linhas, 6 fotos por página (3 fileiras x 2 colunas);
  - cada foto ocupa um bloco mesclado de 13 linhas: B:G (esquerda) e I:N
    (direita), começando nas linhas 4, 20 e 36 de cada página
    (4, 20, 36, 54, 70, 86, 104, ...);
  - abaixo de cada foto, "Descrição:" na coluna B/I e o texto em C:G / J:N;
  - o modelo vem com 11 páginas (66 fotos). Se a proposta tiver mais, o
    robô cria páginas novas no fim, copiando a página 2 (layout, mesclagens,
    alturas e "Descrição:"), e põe quebra de página manual em todas as
    páginas usadas (os modelos só trazem quebras até a pág. 4/6).

Aba de fotos locais (ex.: Banco C, "Mapa e fotos locais"): 4 quadrados
mesclados de 13 linhas, com o título fixo do modelo embaixo (B44, I44, B60,
I60) - o título não é trocado pela legenda da foto. Seções novas criadas na
janela viram quadros extras: o bloco "REGISTRO FOTOGRÁFICO LOCALIDADE"
(linhas 29-61: cabeçalho + 4 quadros) é copiado pra baixo, uma página por
bloco de 4 seções, com o nome da seção embaixo de cada quadro.

Logo da empresa: toda célula do modelo escrita só "Logo" (ex.: Banco D, A1:B2
de cada aba) recebe a logo (assets/logo_empresa.png, ou um logo_empresa.png
ao lado do .exe, que tem prioridade) e perde o texto.

Impressão: no fim, impressao.restaurar devolve a configuração de impressão do
modelo (o Excel de automação a estraga ao salvar) e aplica as áreas novas.

Tamanho: as fotos entram em JPEG reduzido pro tamanho do quadro. O padrão
(2 px por ponto, qualidade 78) mantém legível o carimbo de data/GPS do app de
vistoria; robo.gerar aperta a compressão se o arquivo passar de 14 MB.

Banco D: a aba "FOTOS" do modelo não tem blocos prontos - a grade é montada
na hora (perfis.Grade) só pros quadros usados; a "PAG.1" tem 3 quadros com
o texto dentro ("Foto do imóvel"...), apagado quando a foto entra.
"""

import datetime
import math
import os
import shutil

from PIL import Image, ImageOps

import formulario
import impressao
import pasta_publica
from comum import PREENCHIMENTO
from perfis import BANCO_C

XL_QUEBRA_MANUAL = -4135

MARGEM_PT = 3          # folga entre a foto e a borda do bloco
# (pixels por ponto, qualidade JPEG), do melhor pro mais apertado. O 1º é o
# padrão: ~144 dpi deixa o carimbo de data/GPS da foto legível; qualidade 78
# + progressivo + 4:2:0 economiza ~30% em relação ao antigo (85) sem perda visível.
NIVEIS_COMPRESSAO = ((2.0, 78), (1.6, 72), (1.3, 65), (1.0, 58))
PIXELS_POR_PONTO, QUALIDADE_JPEG = NIVEIS_COMPRESSAO[0]

EXTENSOES_IMAGEM = (".jpg", ".jpeg", ".png", ".bmp", ".webp", ".gif", ".tif", ".tiff")

# atalhos do Banco C (perfil padrão) - usados pelos testes
ABA_FOTOS = BANCO_C.fotos.nome
ABA_LOCAIS = BANCO_C.locais.nome
QUADROS_LOCAIS = {q.chave: (q.bloco, q.titulo) for q in BANCO_C.locais.quadros}


def linha_inicio_pagina(pagina, lay=BANCO_C.fotos):
    """1ª linha da página (0 = 1ª página, que começa no título)."""
    return 1 if pagina == 0 else pagina * lay.linhas_por_pagina + lay.deslocamento_pagina


def slots_fotos(paginas, lay=BANCO_C.fotos):
    """Lista (endereco_bloco, endereco_descricao, pagina) na ordem de leitura."""
    slots = []
    for pagina in range(paginas):
        base = pagina * lay.linhas_por_pagina
        for linha in lay.linhas_foto:
            ini = base + linha
            for col_ini, col_fim, col_desc in lay.colunas:
                slots.append((
                    f"{col_ini}{ini}:{col_fim}{ini + lay.altura_bloco - 1}",
                    f"{col_desc}{ini + lay.deslocamento_descricao}",
                    pagina,
                ))
    return slots


def redimensionar(origem, destino, largura_pt, altura_pt, compressao=None):
    """Gira conforme o EXIF (foto de celular), encaixa dentro do bloco sem
    distorcer e salva JPEG. Devolve (largura_pt, altura_pt) finais."""
    ppp, qualidade = compressao or NIVEIS_COMPRESSAO[0]
    with Image.open(origem) as img:
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "L"):
            fundo = Image.new("RGB", img.size, "white")
            fundo.paste(img, mask=img.convert("RGBA").split()[-1])
            img = fundo
        escala = min(largura_pt / img.width, altura_pt / img.height)
        larg_final = img.width * escala
        alt_final = img.height * escala
        px = (max(1, round(larg_final * ppp)), max(1, round(alt_final * ppp)))
        if px[0] < img.width:  # só reduz; foto pequena não é esticada em pixels
            img = img.resize(px, Image.LANCZOS)
        img.convert("RGB").save(destino, "JPEG", quality=qualidade, optimize=True, progressive=True,
                                subsampling=2)
    return larg_final, alt_final


def listar_imagens(pasta):
    return sorted(
        os.path.join(pasta, n) for n in os.listdir(pasta)
        if n.lower().endswith(EXTENSOES_IMAGEM)
    )


def _abrir_excel():
    import win32com.client
    import win32timezone  # noqa: F401 - o pywin32 importa sozinho ao ler datas; explícito pro .exe levar junto

    excel = win32com.client.DispatchEx("Excel.Application")  # instância própria
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable: não roda macro ao abrir
    return excel


def _garantir_paginas(ws, paginas, lay, log):
    """Cria as páginas que faltam (copiando a página molde) e põe quebra
    manual no começo de cada página usada. Devolve as linhas cuja quebra o
    Excel recusou (gravadas por impressao.restaurar)."""
    if paginas > lay.paginas_no_modelo:
        ini = linha_inicio_pagina(lay.pagina_molde, lay)
        molde = ws.Rows(f"{ini}:{ini + lay.linhas_por_pagina - 1}")
        for pagina in range(lay.paginas_no_modelo, paginas):
            molde.Copy(ws.Rows(linha_inicio_pagina(pagina, lay)))
        log(f"[EXCEL] Modelo ampliado: {paginas - lay.paginas_no_modelo} página(s) nova(s) "
            f"(total {paginas} páginas / {paginas * lay.fotos_por_pagina} fotos).")
    return _por_quebras(ws, [linha_inicio_pagina(p, lay) for p in range(1, paginas)], log)


def _construir_grade(ws, lay, quantas, log):
    """Modelo sem blocos prontos (perfis.Grade): monta só os `quantas`
    primeiros quadros - foto mesclada com borda fina e legenda mesclada
    embaixo - e as quebras de página. Devolve as quebras recusadas."""
    g = lay.grade
    if g.limpar:
        ws.Range(g.limpar).ClearContents()
    if g.colunas_iguais:  # as duas fotos da fileira com o mesmo tamanho, mantendo a largura total
        ini, fim = g.colunas_iguais.split(":")
        faixa = ws.Range(f"{ini}1:{fim}1")
        total, n = faixa.Width, faixa.Columns.Count
        largura = ws.Columns(ini).ColumnWidth
        for _ in range(2):
            ws.Columns(g.colunas_iguais).ColumnWidth = largura
            largura *= (total / n) / ws.Range(f"{ini}1").Width
        ws.Columns(g.colunas_iguais).ColumnWidth = round(largura, 2)
    slots = slots_fotos(-(-quantas // lay.fotos_por_pagina), lay)[:quantas]
    fileiras_feitas = set()
    for bloco, celula_desc, _ in slots:
        topo = ws.Range(bloco).Row
        desc = ws.Range(celula_desc).Row
        if topo not in fileiras_feitas:  # alturas das linhas da fileira
            fileiras_feitas.add(topo)
            ws.Rows(f"{topo}:{topo + lay.altura_bloco - 1}").RowHeight = g.altura_linha_foto
            if desc > topo + lay.altura_bloco:
                ws.Rows(f"{topo + lay.altura_bloco}:{desc - 1}").RowHeight = g.altura_espaco
            ws.Rows(desc).RowHeight = g.altura_legenda
            ws.Rows(desc + 1).RowHeight = g.altura_separador
        col_ini, col_fim = bloco.split(":")[0].rstrip("0123456789"), bloco.split(":")[1].rstrip("0123456789")
        for faixa in (ws.Range(bloco), ws.Range(f"{col_ini}{desc}:{col_fim}{desc}")):
            faixa.Merge()
            faixa.HorizontalAlignment = -4108  # centro
            faixa.VerticalAlignment = -4108
            faixa.BorderAround(1, 2)  # linha contínua, fina
        ws.Range(f"{col_ini}{desc}").ShrinkToFit = True  # legenda longa encolhe em vez de cortar
    paginas = -(-quantas // lay.fotos_por_pagina)
    log(f"[EXCEL] {lay.nome}: grade montada para {quantas} foto(s) em {paginas} página(s).")
    return _por_quebras(ws, [linha_inicio_pagina(p, lay) for p in range(1, paginas)], log)


def _por_quebras(ws, linhas, log):
    """Quebra manual de página antes de cada linha. Devolve as que o Excel
    recusou (ficam pra impressao.restaurar)."""
    for n, linha in enumerate(linhas):
        try:
            ws.Rows(linha).PageBreak = XL_QUEBRA_MANUAL
        except Exception:
            # "Não é possível definir a propriedade PageBreak": o Excel só
            # mexe em quebra de página com impressora instalada (visto numa
            # máquina sem impressora, 07/10/2026). Grava direto no arquivo.
            log("[EXCEL] Sem impressora: as quebras de página serão gravadas direto no arquivo.")
            return linhas[n:]
    return []


def _quadros_extras(ws, quantos, loc, log):
    """Cria os blocos de quadros extras na aba de fotos locais. Devolve
    ([(bloco, célula do título)], linhas de quebra recusadas)."""
    ini, fim = loc.bloco_molde
    altura = fim - ini + 1
    blocos = -(-quantos // len(loc.quadros_no_bloco))
    quadros, linhas_quebra = [], []
    for b in range(blocos):
        destino = fim + 1 + b * altura
        ws.Rows(f"{ini}:{fim}").Copy(ws.Rows(destino))
        linhas_quebra.append(destino)
        for col_ini, col_fim, rel in loc.quadros_no_bloco:
            topo = destino + rel
            quadros.append((f"{col_ini}{topo}:{col_fim}{topo + loc.altura_quadro - 1}",
                            f"{col_ini}{topo + loc.altura_quadro}"))
    for _, celula_titulo in quadros[quantos:]:  # quadros que sobraram no último bloco: sem título
        ws.Range(celula_titulo).Value = ""
    log(f"[EXCEL] {loc.nome}: {blocos} página(s) nova(s) para {quantos} seção(ões) extra(s).")
    _estender_area_impressao(ws, fim + blocos * altura, log, loc.col_fim_impressao)
    return quadros[:quantos], _por_quebras(ws, linhas_quebra, log)


def _paginas_extras(ws, quantos, pe, log):
    """Seções novas numa aba sem bloco pra copiar (perfis.PaginasExtras):
    monta páginas novas no fim (título + quadros de 2 em 2, só os usados).
    Devolve ([(bloco, célula do nome)], linhas de quebra recusadas)."""
    paginas = -(-quantos // pe.quadros_por_pagina)
    quadros, inicios = [], []
    for p in range(paginas):
        ini = pe.linha_inicio + p * pe.linhas_por_pagina
        inicios.append(ini)
        titulo = ws.Range(f"A{ini}:{pe.col_fim}{ini}")
        titulo.Merge()
        titulo.Value = pe.titulo
        titulo.Font.Bold = True
        titulo.Font.Size = 14
        titulo.Interior.Color = 0xD5D5D5  # cinza do cabeçalho do modelo
        titulo.HorizontalAlignment = -4108
        titulo.VerticalAlignment = -4108
        titulo.BorderAround(1, 2)
        ws.Rows(ini).RowHeight = pe.altura_titulo
        ws.Rows(ini + 1).RowHeight = 10
    for n in range(quantos):
        p, k = divmod(n, pe.quadros_por_pagina)
        fileira, coluna = divmod(k, len(pe.colunas))
        topo = inicios[p] + 2 + fileira * pe.linhas_fileira
        col_ini, col_fim = pe.colunas[coluna]
        if coluna == 0:  # alturas da fileira
            ws.Rows(f"{topo}:{topo + pe.linhas_quadro - 1}").RowHeight = pe.altura_linha_foto
            ws.Rows(topo + pe.linhas_quadro).RowHeight = pe.altura_espaco
            ws.Rows(topo + pe.linhas_quadro + 1).RowHeight = pe.altura_legenda
            ws.Rows(topo + pe.linhas_quadro + 2).RowHeight = pe.altura_separador
        bloco = f"{col_ini}{topo}:{col_fim}{topo + pe.linhas_quadro - 1}"
        nome = topo + pe.linhas_quadro + 1
        for faixa in (ws.Range(bloco), ws.Range(f"{col_ini}{nome}:{col_fim}{nome}")):
            faixa.Merge()
            faixa.HorizontalAlignment = -4108
            faixa.VerticalAlignment = -4108
            faixa.BorderAround(1, 2)
        ws.Range(f"{col_ini}{nome}").ShrinkToFit = True
        quadros.append((bloco, f"{col_ini}{nome}"))
    fim = inicios[-1] + pe.linhas_por_pagina - 1
    log(f"[EXCEL] {ws.Name}: {paginas} página(s) de “{pe.titulo}” para {quantos} seção(ões) nova(s).")
    _estender_area_impressao(ws, fim, log, pe.col_fim)
    return quadros, _por_quebras(ws, inicios, log)


# áreas de impressão definidas nesta geração: {aba: "A1:O102"}. Aplicadas no
# arquivo no fim (impressao.restaurar) - pelo Excel de automação não funciona
_areas_definidas = {}


def _estender_area_impressao(ws, fim_necessario, log, col_fim="O"):
    """Anota a área de impressão da aba (de A1 até `fim_necessario`). Quem
    grava é impressao.restaurar, direto no XML: o Excel aberto por automação
    grava "Print_Area" como nome comum e o Excel da pessoa não reconhece."""
    _areas_definidas[ws.Name] = f"A1:{col_fim}{fim_necessario}"
    log(f"[EXCEL] {ws.Name}: área de impressão até a linha {fim_necessario}.")


def _caminho_asset(nome):
    """`nome` ao lado do programa (pra trocar sem gerar o .exe) ou o embutido."""
    from comum import PASTA_RECURSOS, PASTA_SCRIPT

    for pasta in (PASTA_SCRIPT, os.path.join(PASTA_RECURSOS, "assets")):
        caminho = os.path.join(pasta, nome)
        if os.path.exists(caminho):
            return caminho
    return None


def caminho_logo():
    return _caminho_asset("logo_empresa.png")


def _colocar_assinatura(wb, ass, log):
    """Assinatura da responsável centralizada na área do perfil (sem esticar)."""
    arquivo = _caminho_asset("assinatura_responsavel.png")
    if not arquivo:
        log("[EXCEL] [AVISO] Assinatura não encontrada (assinatura_responsavel.png): laudo sem assinatura.")
        return
    try:
        ws = wb.Worksheets(ass.aba)
        ws.Activate()  # mesma regra das fotos: posiciona com a aba ativa
        area = ws.Range(ass.area)
        with Image.open(arquivo) as im:
            proporcao = im.width / im.height
        larg, alt = area.Width - 2 * MARGEM_PT, area.Height - 2 * MARGEM_PT
        if larg / alt > proporcao:
            larg = alt * proporcao
        else:
            alt = larg / proporcao
        shape = ws.Shapes.AddPicture(os.path.abspath(arquivo), False, True, area.Left + (area.Width - larg) / 2,
                                     area.Top + (area.Height - alt) / 2, larg, alt)
        shape.Placement = 1
        shape.Name = "Assinatura_responsavel"
        log(f"[EXCEL] Assinatura da responsável em {ass.aba}!{ass.area}.")
    except Exception as e:  # assinatura não derruba a geração: a pessoa assina na mão
        log(f"[EXCEL] [AVISO] Não consegui pôr a assinatura em {ass.aba}!{ass.area}: {e}")
        pasta_publica.erro(f"assinatura não entrou ({ass.aba}!{ass.area}): {e}")  # não derruba, mas avisa o responsável


def _colocar_logos(wb, log):
    """Põe a logo em toda célula escrita só "Logo" (abas visíveis)."""
    logo = caminho_logo()
    if not logo:
        return
    with Image.open(logo) as im:
        proporcao = im.width / im.height
    postas = 0
    for ws in wb.Worksheets:
        if ws.Visible != -1:  # aba oculta: não dá pra ativar (e não imprime)
            continue
        # lê os valores da aba numa chamada só (o Find do Excel herda as opções da
        # última busca feita pela pessoa - "procurar em comentários" etc. - e falha)
        usada = ws.UsedRange
        try:
            valores = usada.Value
        except Exception as e:  # logo é enfeite: nunca derruba a geração do Excel
            log(f"[EXCEL] [AVISO] Não consegui procurar a célula 'Logo' na aba {ws.Name}: {e}")
            continue
        if not isinstance(valores, tuple):
            valores = ((valores,),)
        celulas = [usada.Cells(i + 1, j + 1) for i, linha in enumerate(valores) for j, v in enumerate(linha)
                   if isinstance(v, str) and v.strip().lower() == "logo"]
        for achada in celulas:
            area = achada.MergeArea
            achada.Value = ""
            ws.Activate()  # mesma regra das fotos: posiciona com a aba ativa
            larg, alt = area.Width - 2, area.Height - 2  # espaço da logo é baixo: margem mínima
            if larg / alt > proporcao:
                larg = alt * proporcao
            else:
                alt = larg / proporcao
            shape = ws.Shapes.AddPicture(os.path.abspath(logo), False, True, area.Left + (area.Width - larg) / 2,
                                         area.Top + (area.Height - alt) / 2, larg, alt)
            shape.Placement = 1
            shape.Name = f"Logo_{postas + 1:02d}"
            postas += 1
    if postas:
        log(f"[EXCEL] Logo da empresa colocada em {postas} lugar(es).")


def _colocar_foto(ws, bloco, foto, redimensionada, nome_figura, compressao=None):
    """Redimensiona `foto` pro tamanho do bloco mesclado e põe centralizada."""
    area = ws.Range(bloco)
    if area.Cells(1, 1).MergeArea.Address.replace("$", "") != bloco:
        raise RuntimeError(f"O modelo não tem o quadro de foto esperado em {ws.Name}!{bloco}.")
    larg, alt = redimensionar(foto, redimensionada, area.Width - 2 * MARGEM_PT, area.Height - 2 * MARGEM_PT,
                              compressao)
    shape = ws.Shapes.AddPicture(os.path.abspath(redimensionada), False, True,
                                 area.Left + (area.Width - larg) / 2,
                                 area.Top + (area.Height - alt) / 2, larg, alt)
    shape.Placement = 1  # xlMoveAndSize: acompanha o bloco
    shape.Name = nome_figura


_ROTULOS_BASICAS = {"F5": "junção", "F6": "nome da junção", "F7": "protocolo", "F11": "proponente", "F15": "bairro", "F16": "cidade", "F17": "CEP",
                    "F18": "UF", "F19": "data de recebimento", "J34": "matrícula"}


def _pre_preencher(wb, pp, dados, log, edicao):
    """Escreve os dados da vistoria na aba do perfil (ex.: FORMULARIO) e
    apaga as células com dados de outro caso que a Central de Gestão não tem.
    `edicao` "basico" = só as células básicas; "completo" = tudo."""
    ws = wb.Worksheets(pp.aba)
    valores, feitos = formulario.MAPAS[pp.mapa](dados or {})
    if edicao != "completo":
        valores = {c: x for c, x in valores.items() if c in pp.basicas}
        feitos = [_ROTULOS_BASICAS[c] for c in pp.basicas if c in valores]
    for celula in tuple(pp.limpar) + tuple(pp.basicas):  # não deixa dado do caso antigo do modelo
        ws.Range(celula).MergeArea.ClearContents()
    for celula, valor in valores.items():
        rng = ws.Range(celula)
        if isinstance(valor, formulario.Formula):
            rng.Formula = str(valor)
        elif isinstance(valor, formulario.Data):  # data de verdade (nº de série), no formato da célula
            rng.Value = (valor - datetime.date(1899, 12, 30)).days
        elif isinstance(valor, str) and valor.isdigit():
            rng.Value = "'" + valor  # como na planilha do banco: texto (o Excel converteria em número)
        else:
            rng.Value = valor  # texto, número, Hora (fração do dia)
    faltaram = [r for c, r in _ROTULOS_BASICAS.items() if c in pp.basicas and c not in valores]
    log(f"[EXCEL] {pp.aba} pré-preenchida (edição {edicao}): {', '.join(feitos) or 'nada'}"
        + (f" | sem dado: {', '.join(faltaram)}" if faltaram else ""))


def inserir_fotos(modelo, saida, fotos, pasta_redimensionadas, descricoes=None, log=print, locais=None,
                  extras=None, perfil=BANCO_C, compressao=None, dados=None):
    """Copia `modelo` para `saida` e insere `fotos` (caminhos) na aba
    principal do `perfil`. `descricoes` (opcional) é uma lista paralela de
    textos. `locais` (opcional) = {chave do quadro: caminho} pra aba de fotos
    locais; `extras` = [(título da seção, caminho)] viram quadros novos nessa
    aba. `dados` = central_gestao.dados_proposta() pro pré-preenchimento do
    perfil (se houver). Devolve quantas fotos entraram."""
    lay, loc = perfil.fotos, perfil.locais
    descricoes = descricoes or []
    locais = {k: v for k, v in (locais or {}).items() if v}
    extras = [(t, v) for t, v in (extras or []) if v]
    if (locais or extras) and loc is None:
        raise ValueError(f"O modelo do {perfil.banco} não tem aba de fotos locais.")
    quadros_fixos = {q.chave: q for q in loc.quadros} if loc else {}
    os.makedirs(pasta_redimensionadas, exist_ok=True)
    paginas = max(1, math.ceil(len(fotos) / lay.fotos_por_pagina))
    quebras_pendentes = {}
    _areas_definidas.clear()

    shutil.copyfile(modelo, saida)
    excel = _abrir_excel()
    wb = None
    try:
        wb = excel.Workbooks.Open(os.path.abspath(saida), UpdateLinks=0)
        # o V3 recalcula a planilha a cada foto inserida (6x mais lento);
        # volta pro automático antes de salvar, pra o arquivo abrir normal
        excel.EnableEvents = False
        excel.Calculation = -4135  # xlCalculationManual

        if perfil.pre_preenchimento:
            try:
                _pre_preencher(wb, perfil.pre_preenchimento, dados, log, PREENCHIMENTO)
            except Exception as e:  # dado estranho da Central de Gestão não pode impedir o Excel com as fotos
                log(f"[EXCEL] [AVISO] Não consegui pré-preencher a {perfil.pre_preenchimento.aba} "
                    f"({type(e).__name__}: {e}); o Excel sai com as fotos - confira o formulário.")
                pasta_publica.erro(f"pré-preenchimento da {perfil.pre_preenchimento.aba} falhou: {e}")

        if locais or extras:
            ws_loc = wb.Worksheets(loc.nome)
            ws_loc.Activate()  # mesma regra da aba de fotos: ativa antes de pôr figura
            quadros_extras = []
            if extras:  # antes das fotos: a cópia do bloco levaria as figuras junto
                if loc.bloco_molde:
                    quadros_extras, pend = _quadros_extras(ws_loc, len(extras), loc, log)
                else:
                    quadros_extras, pend = _paginas_extras(ws_loc, len(extras), loc.paginas_extras, log)
                if pend:
                    quebras_pendentes[loc.nome] = pend
            for chave, foto in locais.items():
                q = quadros_fixos[chave]
                if q.limpar_texto:  # "Foto do imóvel" escrito dentro do quadro: sai, entra a foto
                    ws_loc.Range(q.bloco).Cells(1, 1).Value = ""
                _colocar_foto(ws_loc, q.bloco, foto, os.path.join(pasta_redimensionadas, f"local_{chave}.jpg"),
                              f"Local_{chave}", compressao)
                log(f"[EXCEL] {loc.nome}: {q.titulo} -> {q.bloco}")
            for n, ((titulo, foto), (bloco, celula_titulo)) in enumerate(zip(extras, quadros_extras), 1):
                _colocar_foto(ws_loc, bloco, foto, os.path.join(pasta_redimensionadas, f"local_extra_{n:02d}.jpg"),
                              f"Local_extra_{n:02d}", compressao)
                ws_loc.Range(celula_titulo).Value = titulo
                log(f"[EXCEL] {loc.nome}: {titulo} -> {bloco}")
            ws_loc.Range("A1").Select()

        ws = wb.Worksheets(lay.nome)
        # a aba TEM que estar ativa: senão o Excel invisível posiciona as
        # figuras com as medidas de outra aba/janela e elas "sobem" uns 5 pt
        # por página (visto no V3_2, que abre em quebra de página com zoom 90%)
        ws.Activate()
        if lay.grade:
            pend = _construir_grade(ws, lay, len(fotos), log) if fotos else []
        else:
            pend = _garantir_paginas(ws, paginas, lay, log)  # antes das fotos: a cópia levaria as imagens junto
        if pend:
            quebras_pendentes[lay.nome] = pend
        slots = slots_fotos(paginas, lay)
        ultima_pagina = 0

        for i, foto in enumerate(fotos):
            bloco, celula_desc, pagina = slots[i]
            _colocar_foto(ws, bloco, foto, os.path.join(pasta_redimensionadas, f"interna_{i + 1:03d}.jpg"),
                          f"Foto_{i + 1:02d}", compressao)

            if i < len(descricoes) and descricoes[i]:
                ws.Range(celula_desc).Value = descricoes[i]
            ultima_pagina = pagina
            log(f"[EXCEL] Foto {i + 1}/{len(fotos)} -> {bloco}")

        if fotos:
            _estender_area_impressao(ws, linha_inicio_pagina(ultima_pagina + 1, lay) - 1, log,
                                     lay.col_fim_impressao)

        # logo por último: a grade do Banco D mexe na largura das colunas e esticaria a figura
        _colocar_logos(wb, log)
        if perfil.assinatura:
            _colocar_assinatura(wb, perfil.assinatura, log)
        ws.Activate()
        ws.Range("A1").Select()
        excel.Calculation = -4105  # xlCalculationAutomatic
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()
    # o Excel de automação estraga a impressão ao salvar (nomes especiais e
    # escala); volta tudo do modelo direto no arquivo - ver impressao.py
    impressao.restaurar(modelo, saida, areas=dict(_areas_definidas), quebras=quebras_pendentes, log=log)
    return len(fotos)
