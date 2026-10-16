"""
Pré-preenchimento do formulário do banco com os dados da vistoria no
Central de Gestão (central_gestao.dados_proposta). Hoje: Banco D (aba FORMULARIO).

Pedido do Rafael (09/10/2026): "preencher todos que a Central de Gestão tiver". Cada
campo usa o valor da Central de Gestão traduzido pra lista da planilha (as listas
suspensas da FORMULARIO); quando a tradução não é segura, o campo fica em
branco pra pessoa preencher. NÃO vêm da Central de Gestão (ficam em branco):
junção, nome da junção, tipo de operação, nº do laudo, distância dos
serviços (500/1000/>1000 m - a Central de Gestão só diz se existe), localização
(bom/regular/ruim), áreas do IPTU, valores. Padrões fixos (não vêm do
Central de Gestão): engenheira responsável, CREA e local (SÃO PAULO), data do
laudo = HOJE() e áreas sem averbação = estimada - matrícula (fórmulas).

valores_banco_d(dados) -> ({célula: valor}, [o que foi preenchido]).
Valor: str (texto), int/float (número), Data/Hora (vira data/hora do Excel),
Formula ("=..."), "X" (marcação).
"""

import json
import re
import unicodedata
from datetime import date, datetime, timedelta


class Data(date):
    """Data que vai como data de verdade na célula."""


class Hora(float):
    """Hora como fração do dia (formato h:mm da célula)."""


class Formula(str):
    pass


def _sem_acento(texto):
    return "".join(c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn")


def _chave(texto):
    return re.sub(r"[^a-z0-9]+", " ", _sem_acento(texto).lower()).strip()


def _txt(valor):
    return re.sub(r"\s+", " ", str(valor)).strip() if valor not in (None, "") else ""


def _num(valor):
    """'167,10' / '178.12' / '1.234,5' / '19.220' (milhar) / 3 -> número; '' / None / texto -> None."""
    if valor in (None, ""):
        return None
    if isinstance(valor, (int, float)):
        return valor
    s = str(valor).strip().replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        # "19.220" na Central de Gestão é 19 mil (ponto de milhar, jeito brasileiro) - visto
        # na área do terreno do caso 100002; casas decimais lá vêm com 2 dígitos
        s = s.replace(".", "")
    try:
        n = float(s)
    except ValueError:
        return None
    return int(n) if n.is_integer() else n


def _int(valor):
    n = _num(valor)
    return int(round(n)) if n is not None else None


def _informado(n):
    """Medidas/áreas/idade 0 na Central de Gestão = não informado (o vistoriador deixa 0
    quando não mede; a equipe preenche depois): fica em branco."""
    return n if n not in (None, 0) else None


def _lista(campo):
    """Os checkboxes da Central de Gestão vêm como ['[{"key":..,"label":..}, ...]']."""
    itens = []
    for parte in campo or []:
        try:
            itens += json.loads(parte) if isinstance(parte, str) else [parte]
        except ValueError:
            continue
    return [i for i in itens if isinstance(i, dict)]


def _traduz(valor, tabela):
    """Valor da Central de Gestão -> opção da planilha (pela chave sem acento/caixa)."""
    return tabela.get(_chave(valor)) if valor else None


def _quando(texto, menos_horas=0):
    """'2026-10-07T15:50:56.000Z' -> datetime (menos `menos_horas`)."""
    m = re.match(r"(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d)(?::(\d\d))?", texto or "")
    if not m:
        return None
    return datetime(*(int(x or 0) for x in m.groups())) - timedelta(hours=menos_horas)


# ---------------------------------------------------------------- tabelas (Central de Gestão -> lista da planilha)

UTC_BRASILIA = 3  # a Central de Gestão grava em UTC; a data/hora vale no horário de Brasília

# engenheira responsável pelos laudos (padrão da empresa, confirmado pelo
# Rafael em 09/10/2026 - igual ao laudo pronto de exemplo). CREA como número:
# a célula mostra com zeros à esquerda (formato 0000000000000)
RESPONSAVEL_NOME = "ENGENHEIRA RESPONSÁVEL EXEMPLO"
RESPONSAVEL_CREA = 1234567890
RESPONSAVEL_LOCAL = "SÃO PAULO"

TIPOS_LOGRADOURO = ["AEROPORTO", "ALAMEDA", "ÁREA", "AVENIDA", "CAMPO", "CHÁCARA", "Colônia", "Condomínio",
                    "Conjunto", "Distrito", "Esplanada", "Estação", "Estrada", "Favela", "Fazenda", "Feira", "Jardim",
                    "Ladeira", "Lago", "Lagoa", "Largo", "Loteamento", "Morro", "Núcleo", "Parque", "Passarela",
                    "Pátio", "Praça", "Quadra", "Recanto", "Residencial", "Rodovia", "Rua", "Setor", "Sítio",
                    "Travessa", "Trecho", "Trevo", "Vale", "Vereda", "Via", "Viaduto", "Viela", "Vila"]
ABREV_LOGRADOURO = {"r": "Rua", "av": "AVENIDA", "avda": "AVENIDA", "al": "ALAMEDA", "trav": "Travessa",
                    "tv": "Travessa", "pc": "Praça", "pca": "Praça", "rod": "Rodovia", "estr": "Estrada",
                    "rua": "Rua", "avenida": "AVENIDA", "alameda": "ALAMEDA"}

BOM_REGULAR_RUIM = {"boa": "BOM", "bom": "BOM", "regular": "REGULAR", "ruim": "RUIM"}
CONSERVACAO = {"bom": "BOM", "otimo": "BOM", "regular": "REGULAR", "ruim": "RUIM",
               "em implantacao": "EM IMPLANTAÇÃO", "em construcao": "EM IMPLANTAÇÃO"}
OCUPACAO = {"habitado": "OCUPADO", "ocupado": "OCUPADO", "desabitado": "DESOCUPADO", "desocupado": "DESOCUPADO"}
USO = {"residencial": "RESIDENCIAL", "comercial": "COMERCIAL", "misto": "MISTO", "industrial": "INDUSTRIAL"}
PREDOMINANCIA = {"residencial unifamiliar": "RESIDENCIAL", "residencial multifamiliar": "RESIDENCIAL",
                 "residencial": "RESIDENCIAL", "comercial": "COMERCIAL", "industrial": "INDUSTRIAL", "misto": "MISTO"}
# Central de Gestão tem 5 níveis (Baixo..Alto), a planilha 4 (ALTO/NORMAL/BAIXO/MINIMO): médio = normal
PADRAO = {"alto": "ALTO", "medio alto": "NORMAL", "medio": "NORMAL", "normal": "NORMAL", "medio baixo": "BAIXO",
          "baixo": "BAIXO", "minimo": "MINIMO"}
BAIRRO = {"antigo": "ANTIGO", "novo": "NOVO", "em desenvolvimento": "EM DESENVOLVIMENTO",
          "em mudanca de ocupacao": "EM MUDANÇA DE OCUPAÇÃO"}
TRAFEGO = {"reduzido bom": "REDUZIDO", "moderado regular": "MODERADO", "intenso ruim": "INTENSO"}
TOPOGRAFIA = {"plana": "PLANO", "aclive": "ACLIVE", "declive": "DECLIVE"}
FORMATO = {"regular": "REGULAR", "irregular": "IRREGULAR"}
IMPLANTACAO = {"isolado": "ISOLADA", "isolada": "ISOLADA", "condominio": "CONDOMINIO"}
FECHAMENTO = {"alvenaria": "ALVENARIA", "parede de gesso drywall": "DRY WALL", "madeira": "MADEIRA"}
ESTRUTURA = {"alvenaria concreto": "ALVENARIA/ CONCRETO", "madeira": "MADEIRA", "mista": "MISTA",
             "concreto pre moldado": "CONCRETO PRÉ-MOLDADO", "oncreto pre moldado": "CONCRETO PRÉ-MOLDADO"}
TIPO_IMOVEL = {"casa": "CASA", "sobrado": "CASA", "apartamento": "APARTAMENTO", "sala comercial": "SALA",
               "loja": "LOJA", "galpao comercial": "GALPÃO", "galpao": "GALPÃO", "predio comercial": "PRÉDIO COMERCIAL",
               "misto": "MISTO", "industria": "INDÚSTRIA", "hotel": "HOTEL", "pousada": "HOTEL",
               "posto de gasolina": "POSTO DE COMBUSTÍVEL", "hospital": "HOSPITAL", "shopping": "SHOPPING",
               "chacara": "CHÁCARA", "tereno ate 2 000mts": "TERRENO ATÉ 5MIL M²",
               "terreno ate 2 000mts": "TERRENO ATÉ 5MIL M²"}
INFRA = {"agua": "F83", "esgoto sanitario": "F84", "energia eletrica": "F85", "telefone": "F86",
         "pavimentacao": "F87", "arborizacao": "F88", "esgoto pluvial": "F89", "gas canalizado": "F90",
         "iluminacao publica": "F91", "fossa": "F92", "cisterna poco": "F93", "poco": "F93"}
LAZER_PREDIO = {"piscina": "F161", "playground": "F162", "salao de festas": "F163", "quadra esportiva": "F164",
                "quadra": "F164", "churrasqueira": "F165", "gerador": "F166"}
# cômodos: Quartos (118), Suítes (119) e Banheiro (120) têm linha certa no
# formulário; os outros vão de 121 a 130
ULTIMA_LINHA_COMODO = 130
ESQUADRIAS = {"aluminio": "ALUMÍNIO", "ferro": "FERRO", "madeira": "MADEIRA", "pvc": "PVC", "plastico": "PLÁSTICO",
              "vidro": "VIDRO"}


def _linha_fixa(tipo):
    """118/119/120 pros grupos fixos, 0 pra banheiro de suíte (já contado na
    suíte), None pros demais (linhas 121-130)."""
    k = _chave(tipo)
    if k.startswith("banheiro") and "suite" in k:
        return 0
    if k.startswith("suite"):
        return 119
    if k.startswith(("dormitorio", "quarto")):
        return 118
    if k.startswith("banheiro"):
        return 120
    return None


def _logradouro(texto):
    """'Rua W1' -> ('Rua', 'W1'); 'Av. Brasil' -> ('AVENIDA', 'Brasil')."""
    texto = _txt(texto)
    m = re.match(r"([^\s.]+)\.?\s+(.*)", texto)
    if m:
        primeira = _chave(m.group(1))
        tipo = ABREV_LOGRADOURO.get(primeira) or next((t for t in TIPOS_LOGRADOURO if _chave(t) == primeira), None)
        if tipo:
            return tipo, m.group(2).strip()
    return None, texto


# ---------------------------------------------------------------- Banco D

def valores_banco_d(dados):
    """{célula da FORMULARIO: valor} + lista do que foi preenchido."""
    v, feitos = {}, []
    im = dados.get("imovel") or {}
    reg = dados.get("regiao") or {}
    ter = dados.get("terreno") or {}
    cond = dados.get("condominioImovel") or {}
    eng = dados.get("engenheiro") or {}

    def pôr(celula, valor, rotulo=None):
        if valor not in (None, "", []):
            v[celula] = valor
            if rotulo and rotulo not in feitos:
                feitos.append(rotulo)

    # ---- identificação (o que a planilha dos funcionários já trazia + endereço)
    pôr("F7", _txt(dados.get("id") or im.get("id")), "protocolo")
    # junção e nome da junção: da planilha que o Banco D manda (anexada na
    # requisição da Intranet - intranet.dados_do_banco), não da Central de Gestão
    # (pedido do Rafael, 14/10/2026). Junção como número, como vem do banco.
    banco = dados.get("banco") or {}
    juncao = banco.get("juncao")
    pôr("F5", int(juncao) if str(juncao or "").strip().isdigit() else _txt(juncao), "junção")
    pôr("F6", _txt(banco.get("nome_juncao")), "junção")
    pôr("F11", _txt(im.get("proponenteComprador")), "proponente")
    tipo_log, nome_log = _logradouro(im.get("logradouro"))
    pôr("F12", tipo_log, "endereço")
    pôr("G12", nome_log, "endereço")
    pôr("F13", _txt(im.get("numero")), "nº")
    compl = [_txt(im.get("complemento"))] + ([f"Andar {_txt(im.get('andar'))}"] if _txt(im.get("andar")) else [])
    pôr("F14", " - ".join(c for c in compl if c), "complemento")
    pôr("F15", _txt(im.get("bairro")), "bairro")
    pôr("F16", _txt(im.get("city")), "cidade")
    # CEP como número: a célula tem formato 00000-000 (mostra o hífen e o zero da
    # frente), como nos laudos prontos da equipe
    cep = re.sub(r"\D", "", _txt(im.get("cep")))
    pôr("F17", int(cep) if cep else None, "CEP")
    pôr("F18", _txt(im.get("uf")).upper(), "UF")
    criada = _quando(dados.get("_createdAt"), UTC_BRASILIA)
    pôr("F19", Data(criada.year, criada.month, criada.day) if criada else None, "data de recebimento")
    # visita: data/hora da vistoria (UTC); senão a data agendada (UTC)
    visita = (_quando((dados.get("questoesComplementares") or {}).get("dateVistoria"), UTC_BRASILIA)
              or _quando(dados.get("date"), UTC_BRASILIA))
    if visita:
        pôr("F20", Data(visita.year, visita.month, visita.day), "data da visita")
    # matrícula vai em "Nº Matrícula" (J34), com 7 dígitos ("0012345"), como nos
    # 3 laudos prontos da equipe (13/10/2026); F34 é o "CNS Cartório" (6 dígitos,
    # da certidão - a Central de Gestão não tem), Livro/Dígito também ficam pra equipe
    matricula = re.sub(r"\D", "", _txt(im.get("matricula")))
    pôr("J34", matricula.zfill(7) if matricula else None, "matrícula")
    # Livro sempre 2 (Livro nº 2 - Registro Geral, onde ficam as matrículas; os 3
    # laudos prontos da equipe usam 2 - pedido do Rafael, 13/10/2026)
    pôr("H34", "2", "matrícula")

    # ---- vagas e áreas
    descobertas, cobertas = _int(cond.get("vagasDescobertas")), _int(cond.get("vagasCobertas"))
    if descobertas is not None or cobertas is not None:
        pôr("N37", descobertas or 0, "vagas")
        pôr("P37", cobertas or 0, "vagas")
        pôr("F37", Formula("=N37+P37"))
        pôr("F38", "SIM" if (descobertas or 0) + (cobertas or 0) > 0 else "NÃO ")
    area_terreno = _informado(_num(ter.get("area")))
    # construída: a privativa (a equipe usa ela nos apartamentos - casos 100002
    # e 100005; a averbada lá é a total); sem privativa (casa), a averbada
    construida = _informado(_num(cond.get("areaPrivativa"))) or _informado(_num(cond.get("areaAverbada")))
    pôr("F42", area_terreno, "áreas")
    pôr("F43", construida, "áreas")
    pôr("F44", _num(cond.get("areaComum")), "áreas")
    pôr("F47", Formula("=F44+F43"), "áreas")  # total = construída + comum (como nos casos da equipe)
    fracao = _num(ter.get("fracaoTotal"))
    pôr("F48", (fracao / 100 if fracao and fracao > 1 else fracao), "fração ideal")
    pôr("J42", area_terreno)
    pôr("J43", construida)
    # sem averbação = construída estimada - construída da matrícula (pedido do
    # Rafael, 09/10/2026: diferença, não o campo da Central de Gestão); fórmula, pra
    # acompanhar quando a pessoa corrigir a área estimada
    pôr("I48", Formula("=J43-F43"), "áreas")

    # ---- situação
    pôr("F57", _traduz(reg.get("comercializacao"), BOM_REGULAR_RUIM), "comercialização")
    pôr("F59", _traduz(cond.get("indicioOcupacaoImovel"), OCUPACAO), "ocupação")
    pôr("F60", _txt(dados.get("observacoes")), "comentário")
    # data do laudo = hoje (fórmula) e engenheira responsável fixa (padrão da empresa)
    pôr("F61", Formula("=TODAY()"), "data do laudo")
    pôr("F62", RESPONSAVEL_NOME, "engenheira responsável")
    pôr("F63", RESPONSAVEL_CREA, "engenheira responsável")
    pôr("F64", RESPONSAVEL_LOCAL, "engenheira responsável")
    uso = _traduz(cond.get("usoImovel"), USO)
    pôr("F65", {"RESIDENCIAL": "NÃO ", "MISTO": "SIM"}.get(uso, "NÃO SE APLICA" if uso else None))

    # ---- região
    pôr("J71", _traduz(reg.get("bairro"), BAIRRO), "região")
    pôr("J73", _traduz(reg.get("regiao"), PREDOMINANCIA), "região")
    pôr("J75", _traduz(reg.get("padraoregiao"), PADRAO), "região")
    pôr("J77", _traduz(reg.get("trafego"), TRAFEGO), "região")
    for item in _lista(reg.get("Infraestrutura")):
        celula = INFRA.get(_chave(item.get("key") or item.get("label")))
        pôr(celula, "X", "infraestrutura") if celula else None

    # ---- terreno
    pôr("I83", area_terreno)
    pôr("I86", _txt(ter.get("zoneamento") or im.get("zoneamentoCity")))
    pôr("I89", _traduz(ter.get("topografia"), TOPOGRAFIA), "terreno")
    pôr("I92", _traduz(ter.get("formato"), FORMATO), "terreno")
    for celula, campo in (("F96", "testadaFrente"), ("F97", "fundos"), ("F98", "ladoDireito"),
                          ("F99", "ladoEsquerdo")):
        pôr(celula, _informado(_num(ter.get(campo))), "medidas do terreno")
    lat, lon = dados.get("latitude"), dados.get("longitude")
    if lat not in (None, "") and lon not in (None, ""):
        pôr("E101", f"{lat}, {lon}", "coordenada")

    # ---- imóvel
    pôr("E106", _traduz(cond.get("tipoImovelAvaliado") or im.get("type"), TIPO_IMOVEL), "imóvel")
    pôr("G106", _traduz(cond.get("estrutura"), ESTRUTURA), "imóvel")
    pôr("E108", _traduz(ter.get("tipoImplantacao"), IMPLANTACAO), "imóvel")
    pôr("G108", _traduz(cond.get("fechamentoDasParedes"), FECHAMENTO), "imóvel")
    pôr("E110", _traduz(cond.get("padraoAcabamentoImovel"), PADRAO), "imóvel")
    esquadria = cond.get("esquadriasOther") if _chave(cond.get("esquadrias") or "") == "outro" else cond.get("esquadrias")
    # texto livre: em maiúsculas, com acento como a equipe escreve ("ALUMÍNIO";
    # a Central de Gestão grava "Aluminío")
    pôr("G110", ESQUADRIAS.get(_chave(esquadria or "")) or _txt(esquadria).upper(), "imóvel")
    pôr("E112", _traduz(cond.get("estadoConservacaoImovel"), CONSERVACAO), "imóvel")
    pôr("G112", uso, "imóvel")
    pôr("E114", _informado(_int(cond.get("idadeAparenteImovelAvaliado"))), "imóvel")
    pôr("G114", _informado(_int(cond.get("numeroAndaresDoImovelAvaliado"))), "imóvel")

    # ---- cômodos: só as 13 linhas do formulário (118-130, as que a PAG.2
    # mostra); NUNCA cria linha nova - o que passar disso fica de fora e
    # vai pro log (pedido do Rafael, 09/10/2026)
    # Linhas fixas como a equipe faz (laudos 100002 e 100005): "Quartos" soma
    # todos os dormitórios, "Suítes" as suítes e "Banheiro" os banheiros fora os
    # de suíte (a suíte já é quarto + banheiro); esses não se repetem embaixo.
    proxima, fora, fixos = 121, [], {}
    comodos = [c for c in dados.get("comodos") or [] if isinstance(c, dict) and _txt(c.get("divisaoInterna"))]
    for c in comodos:
        tipo = _txt(c.get("divisaoInterna"))
        linha = _linha_fixa(tipo)
        if linha == 0:
            continue  # banheiro da suíte: já está na suíte
        if linha:
            qtd = _int(c.get("quantidade")) or 0
            if linha in fixos:  # mais um do mesmo grupo: soma a quantidade
                fixos[linha] += qtd
                v[f"H{linha}"] = fixos[linha]
                continue
            fixos[linha] = qtd
        else:
            if proxima > ULTIMA_LINHA_COMODO:
                fora.append(tipo)
                continue
            linha, proxima = proxima, proxima + 1
            pôr(f"E{linha}", tipo)
        pôr(f"H{linha}", _int(c.get("quantidade")), "cômodos")
        pôr(f"I{linha}", _txt(c.get("parede")))
        pôr(f"J{linha}", _txt(c.get("piso")))
        pôr(f"K{linha}", _txt(c.get("forroTeto")))
    if fora:
        feitos.append(f"[AVISO] {len(fora)} cômodo(s) não couberam nas 13 linhas do formulário e ficaram de fora: "
                      + ", ".join(fora))
    if not comodos:  # sem a lista de cômodos: pelo menos dormitórios e banheiros
        pôr("H118", _informado(_int(cond.get("qtqDormitorios"))), "cômodos")
        pôr("H120", _informado(_int(cond.get("qtdBanheiros"))), "cômodos")

    # ---- vistoriador
    pôr("E135", re.sub(r"\s+-\s+.*$", "", _txt(eng.get("name"))), "vistoriador")
    pôr("E136", _txt(eng.get("creci")), "vistoriador")
    if visita:
        pôr("E137", Data(visita.year, visita.month, visita.day))
        if (dados.get("questoesComplementares") or {}).get("dateVistoria"):
            pôr("E138", Hora((visita.hour * 60 + visita.minute) / 1440))

    # ---- condições e fatores
    garantia = (dados.get("garantia") or {}).get("condGarantia")
    pôr("F143", {True: "SIM ", False: "NÃO "}.get(garantia), "habitabilidade")
    def marcados(campo, texto_extra=None):
        """Labels marcados; a caixa "Outro" vale pelo texto digitado (OtherDescription)."""
        itens = _lista(campo)
        nomes = [_txt(i.get("label")) for i in itens
                 if _chave(i.get("key") or "") not in ("outro", "otherdescription") and _txt(i.get("label"))]
        outros = [_txt(i.get("label")) for i in itens if _chave(i.get("key") or "") == "otherdescription"]
        outros += [_txt(texto_extra)] if _txt(texto_extra) else []
        return nomes, [o for o in outros if _chave(o) not in ("nenhum", "nenhuma", "nao", "nada", "")]

    nomes, outros = marcados(reg.get("faroresValorizantes"), reg.get("textFatoresValorizantes"))
    pôr("F144", ", ".join(nomes + outros), "fatores")
    nomes, outros = marcados(reg.get("restritivos"), reg.get("textRestritivo"))
    pôr("F145", ", ".join(nomes + outros), "fatores")
    if reg.get("restritivos") is not None:
        pôr("F147", "SIM" if any(_chave(n) == "risco ambiental" for n in nomes) else "NÃO", "fatores")
    estado = _chave(cond.get("estadoConservacaoImovel") or "")
    if estado:
        pôr("F148", "SIM" if ("constru" in estado or "implanta" in estado) else "NÃO")

    # ---- prédio (só apartamento / condomínio)
    em_predio = _chave(cond.get("tipoImovelAvaliado") or im.get("type") or "") == "apartamento"
    if em_predio:
        pôr("F155", _traduz(cond.get("usoCondominio"), USO), "prédio")
        # pavimentos / unidades por andar 0 = não informado (caso 100002: equipe pôs 2 e 1)
        pôr("F156", _informado(_int(cond.get("andaresTipo"))), "prédio")
        pôr("F157", _int(cond.get("qtdSubsolos")), "prédio")
        pôr("F158", _informado(_int(cond.get("unidadesPorAndar"))), "prédio")
        pôr("F159", _int(cond.get("elevaores")), "prédio")
        estado_cond = _traduz(cond.get("estadoConservacapCondominio"), CONSERVACAO)
        pôr("F160", estado_cond, "prédio")
        pôr("F167", estado_cond, "prédio")
        tem = {_chave(i.get("key") or i.get("label")) for i in _lista(cond.get("InfraestruturaDoImovel"))}
        if cond.get("InfraestruturaDoImovel") is not None:
            for chave, celula in LAZER_PREDIO.items():
                if celula not in v:
                    pôr(celula, "SIM" if chave in tem else "NÃO ", "prédio")
            for chave, celula in LAZER_PREDIO.items():  # "quadra" e "quadra esportiva" na mesma célula
                if chave in tem:
                    v[celula] = "SIM"
    return v, feitos


MAPAS = {"banco_d": valores_banco_d}
