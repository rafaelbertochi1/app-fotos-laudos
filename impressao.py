"""
Conserta a configuração de impressão do Excel gerado, direto no arquivo
(.xlsx/.xlsm é um zip de XMLs), depois que o Excel salvou e fechou.

Por quê (achado em 09/10/2026, reclamação: "imprime várias páginas a mais
onde nem tem foto"): o Excel aberto por automação (invisível ou não) não
reconhece os nomes especiais ao abrir o modelo e, ao salvar, grava
"_xlnm.Print_Area" como um nome comum "Print_Area". Resultado no arquivo
entregue: NENHUMA aba tem área de impressão - o Excel imprime a aba
inteira (molduras vazias, colunas de apoio) - e as abas "ajustar a N
páginas" ficam com a escala recalculada pra caber tudo (ex.: 85% -> 26%).
Também afeta o filtro (_FilterDatabase) e o critério de filtro (Criteria).

O conserto não depende do Excel:
  1. os nomes especiais (_xlnm.*) voltam a ser exatamente os do modelo;
  2. a configuração de página (<pageSetup>) de cada aba volta a ser a do
     modelo (escala, ajuste, papel, orientação);
  3. as áreas de impressão que o programa mudou (aba de fotos: até a última
     página com foto; aba de quadros: até a última página de seções novas)
     entram por cima;
  4. quebras de página que o Excel recusou (máquina sem impressora) entram
     no XML.
"""

import os
import re
import zipfile
from xml.sax.saxutils import escape, unescape

_FIM_DA_ABA = (r"<(colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|"
               r"legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|</worksheet>")


def _ler_zip(caminho):
    with zipfile.ZipFile(caminho) as z:
        itens = [(info, z.read(info.filename)) for info in z.infolist()]
    return itens, {info.filename: dados for info, dados in itens}


def _abas(conteudo):
    """[(nome da aba, caminho do XML)] na ordem do arquivo (= localSheetId)."""
    wb = conteudo["xl/workbook.xml"].decode("utf-8")
    rels = conteudo["xl/_rels/workbook.xml.rels"].decode("utf-8")
    alvos = {}
    for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
        rid = re.search(r'\bId="([^"]+)"', rel).group(1)
        alvo = re.search(r'\bTarget="([^"]+)"', rel).group(1).lstrip("/")
        alvos[rid] = alvo if alvo.startswith("xl/") else "xl/" + alvo
    abas = []
    for tag in re.findall(r"<sheet\b[^>]*/>", wb):
        nome = unescape(re.search(r'\bname="([^"]*)"', tag).group(1), {"&quot;": '"', "&apos;": "'"})
        rid = re.search(r'\br:id="([^"]+)"', tag).group(1)
        abas.append((nome, alvos[rid]))
    return abas


def _nomes(wb_xml):
    """[(tag inteira, nome, localSheetId ou None)] dos <definedName>."""
    out = []
    for m in re.finditer(r"<definedName\b([^>]*)>(.*?)</definedName>", wb_xml, re.S):
        attrs = m.group(1)
        nome = re.search(r'\bname="([^"]*)"', attrs).group(1)
        ls = re.search(r'\blocalSheetId="(\d+)"', attrs)
        out.append((m.group(0), nome, int(ls.group(1)) if ls else None))
    return out


def _ref_area(aba, faixa):
    """'Aba'!$A$1:$O$102 (sempre com aspas; aspa no nome vira duas)."""
    ini, fim = faixa.split(":")
    absol = lambda c: re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", c)
    return "'%s'!%s:%s" % (aba.replace("'", "''"), absol(ini), absol(fim))


def restaurar(modelo, saida, areas=None, quebras=None, log=print):
    """`areas` = {aba: "A1:O102"} que o programa definiu; `quebras` = {aba:
    [linhas]} que o Excel recusou. Reescreve `saida` no lugar."""
    areas, quebras = areas or {}, quebras or {}
    _, mod = _ler_zip(modelo)
    itens, ger = _ler_zip(saida)
    abas_mod, abas_ger = _abas(mod), _abas(ger)
    if [a for a, _ in abas_mod] != [a for a, _ in abas_ger]:
        log("[EXCEL] [AVISO] As abas do arquivo gerado não batem com as do modelo; impressão não restaurada.")
        return
    nomes_abas = [a for a, _ in abas_ger]

    # 1 + 3) nomes especiais do modelo de volta (com as áreas novas por cima)
    wb_mod = mod["xl/workbook.xml"].decode("utf-8")
    wb_ger = ger["xl/workbook.xml"].decode("utf-8")
    especiais = [(tag, nome, ls) for tag, nome, ls in _nomes(wb_mod) if nome.startswith("_xlnm.")]
    chaves = {(nome[len("_xlnm."):], ls) for _, nome, ls in especiais} | {(nome, ls) for _, nome, ls in especiais}
    for tag, nome, ls in _nomes(wb_ger):  # tira as versões estragadas ("Print_Area" comum) e as _xlnm atuais
        if (nome, ls) in chaves:
            wb_ger = wb_ger.replace(tag, "", 1)
    novos = []
    feitas = set()
    for tag, nome, ls in especiais:
        aba = nomes_abas[ls] if ls is not None else None
        if nome == "_xlnm.Print_Area" and aba in areas:
            tag = re.sub(r">.*?</definedName>", ">%s</definedName>" % escape(_ref_area(aba, areas[aba])), tag,
                         flags=re.S)
            feitas.add(aba)
        novos.append(tag)
    for aba, faixa in areas.items():  # aba que não tinha área no modelo
        if aba not in feitas:
            novos.append('<definedName name="_xlnm.Print_Area" localSheetId="%d">%s</definedName>'
                         % (nomes_abas.index(aba), escape(_ref_area(aba, faixa))))
    if novos:
        if "<definedNames>" in wb_ger:
            wb_ger = wb_ger.replace("<definedNames>", "<definedNames>" + "".join(novos), 1)
        elif "<definedNames/>" in wb_ger:
            wb_ger = wb_ger.replace("<definedNames/>", "<definedNames>" + "".join(novos) + "</definedNames>", 1)
        else:
            wb_ger = re.sub(r"(</sheets>)", r"\1<definedNames>" + "".join(novos).replace("\\", "\\\\") +
                            "</definedNames>", wb_ger, count=1)
    ger["xl/workbook.xml"] = wb_ger.encode("utf-8")

    # 2 + 4) configuração de página do modelo de volta, e quebras pendentes
    for (aba, xml_mod), (_, xml_ger) in zip(abas_mod, abas_ger):
        x_mod = mod[xml_mod].decode("utf-8")
        x = ger[xml_ger].decode("utf-8")
        ps_mod = re.search(r"<pageSetup\b[^>]*/>", x_mod)
        ps_ger = re.search(r"<pageSetup\b[^>]*/>", x)
        if ps_mod and ps_ger:
            rid = re.search(r'\br:id="[^"]*"', ps_ger.group(0))  # vínculo com as configs da impressora deste arquivo
            novo = re.sub(r'\s+r:id="[^"]*"', "", ps_mod.group(0))
            if rid:
                novo = novo.replace("/>", " %s/>" % rid.group(0), 1)
            x = x[:ps_ger.start()] + novo + x[ps_ger.end():]
        if quebras.get(aba):
            existente = re.search(r"<rowBreaks\b[^>]*/>|<rowBreaks\b.*?</rowBreaks>", x, re.S)
            ids = {int(i) for i in re.findall(r'<brk\b[^>]*\bid="(\d+)"', existente.group(0))} if existente else set()
            ids |= {linha - 1 for linha in quebras[aba]}  # id = última linha antes da quebra
            bloco = '<rowBreaks count="%d" manualBreakCount="%d">%s</rowBreaks>' % (
                len(ids), len(ids), "".join('<brk id="%d" max="16383" man="1"/>' % i for i in sorted(ids)))
            if existente:
                x = x[:existente.start()] + bloco + x[existente.end():]
            else:  # ordem do schema: rowBreaks vem antes destes elementos
                depois = re.search(_FIM_DA_ABA, x)
                x = x[:depois.start()] + bloco + x[depois.start():]
        ger[xml_ger] = x.encode("utf-8")

    tmp = saida + ".tmp"
    with zipfile.ZipFile(tmp, "w") as z:
        for info, _ in itens:
            z.writestr(info, ger[info.filename], compress_type=info.compress_type)
    os.replace(tmp, saida)
    resumo = ", ".join(f"{a} até a linha {f.split(':')[1].lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')}" for a, f in areas.items())
    log(f"[EXCEL] Impressão restaurada do modelo ({len(especiais)} nomes especiais)"
        + (f"; área de impressão: {resumo}" if resumo else "") + ".")


def conferir(modelo, saida):
    """Pra teste: diferenças de impressão entre modelo e gerado (lista vazia = igual)."""
    _, mod = _ler_zip(modelo)
    _, ger = _ler_zip(saida)
    dif = []
    pega = lambda xml: {(n, ls): re.search(r">(.*?)</definedName>", t, re.S).group(1)
                        for t, n, ls in _nomes(xml) if n.startswith("_xlnm.") or n in ("Print_Area", "_FilterDatabase", "Criteria")}
    nm, ng = pega(mod["xl/workbook.xml"].decode("utf-8")), pega(ger["xl/workbook.xml"].decode("utf-8"))
    if set(nm) != set(ng):
        dif.append(("nomes especiais", sorted(set(nm) ^ set(ng))[:6]))
    for (aba, a), (_, b) in zip(_abas(mod), _abas(ger)):
        pa = lambda xml: re.sub(r'\s+r:id="[^"]*"', "", (re.search(r"<pageSetup\b[^>]*/>", xml) or [""])[0] if
                                re.search(r"<pageSetup\b[^>]*/>", xml) else "")
        if pa(mod[a].decode("utf-8")) != pa(ger[b].decode("utf-8")):
            dif.append((aba, "pageSetup"))
    return dif, ng
