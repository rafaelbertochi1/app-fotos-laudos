"""
Propostas FICTÍCIAS usadas pelo servidor de exemplo (servidor_exemplo.py):
o formato é o da API real da Central de Gestão (GET /inspection/<_id>) e da
planilha que o Banco D anexa na Intranet, com os mesmos "defeitos" que o
programa já trata (campo com erro de digitação, número com ponto de milhar,
cômodo vazio na lista, foto antiga sem "ambiente", JPEG com extensão .png).

Nomes, endereços, documentos e números são inventados.
"""

import json

USUARIO = "teste@example.com"
SENHA = "senha-de-exemplo"
TOKEN = "token-de-exemplo"


def _lista(*itens):
    """Checkbox da Central de Gestão: lista com UM texto JSON dentro."""
    return [json.dumps([{"key": k, "label": r} for k, r in itens], ensure_ascii=False)]


def _fotos(proposta, ambientes, observacoes=None, sem_ambiente=()):
    """fotos.fotos[] como a API devolve: um item por ambiente, várias imagens."""
    observacoes = observacoes or {}
    itens = []
    for ambiente, quantas in ambientes:
        nome = "".join(c for c in ambiente.title() if c.isalnum())
        imagens = [f"public/images/{proposta}/{nome}_20261014T10{n:02d}00[{n}].png" for n in range(quantas)]
        itens.append({"ambiente": "" if ambiente in sem_ambiente else ambiente, "images": imagens,
                      "observacao": observacoes.get(ambiente, "")})
    return {"fotos": itens}


def _casa(proposta, **mudancas):
    """Uma casa térrea genérica; `mudancas` troca blocos inteiros."""
    vistoria = {
        "_id": f"exemplo{proposta}",
        "id": int(proposta),
        "_createdAt": "2026-10-13T13:05:00.000Z",
        "date": "2026-10-14T12:00:00.000Z",
        "observacoes": "Imóvel em bom estado. Vistoria realizada com o proprietário (dados de exemplo).",
        "latitude": "-23.000000",
        "longitude": "-46.000000",
        "imovel": {"id": int(proposta), "proponenteComprador": "FULANO DE TAL (EXEMPLO)",
                   "logradouro": "R. das Acácias", "numero": "100", "complemento": "", "andar": "",
                   "bairro": "Jardim Exemplo", "city": "Cidade Exemplo", "cep": "01000-000", "uf": "sp",
                   "matricula": "12.345", "type": "Casa", "zoneamentoCity": "ZM"},
        "regiao": {"comercializacao": "Boa", "bairro": "Antigo", "regiao": "Residencial unifamiliar",
                   "padraoregiao": "Médio", "trafego": "Reduzido/Bom",
                   "Infraestrutura": _lista(("agua", "Água"), ("esgotoSanitario", "Esgoto sanitário"),
                                            ("energiaEletrica", "Energia elétrica"),
                                            ("pavimentacao", "Pavimentação"),
                                            ("iluminacaoPublica", "Iluminação pública")),
                   "faroresValorizantes": _lista(("comercio", "Comércio próximo")),
                   "textFatoresValorizantes": "Nenhum",
                   "restritivos": _lista(), "textRestritivo": ""},
        "terreno": {"area": "250,00", "zoneamento": "", "topografia": "Plana", "formato": "Regular",
                    "testadaFrente": "10", "fundos": "10", "ladoDireito": "25", "ladoEsquerdo": "25",
                    "fracaoTotal": "100", "tipoImplantacao": "Isolado"},
        "condominioImovel": {"vagasDescobertas": "1", "vagasCobertas": "1", "areaPrivativa": "0",
                             "areaAverbada": "120,50", "areaComum": "0", "indicioOcupacaoImovel": "Habitado",
                             "usoImovel": "Residencial", "tipoImovelAvaliado": "Casa",
                             "estrutura": "Alvenaria/Concreto", "fechamentoDasParedes": "Alvenaria",
                             "padraoAcabamentoImovel": "Médio", "esquadrias": "Aluminío", "esquadriasOther": "",
                             "estadoConservacaoImovel": "Bom", "idadeAparenteImovelAvaliado": "15",
                             "numeroAndaresDoImovelAvaliado": "1", "qtqDormitorios": "2", "qtdBanheiros": "1"},
        "comodosAmbientes": {"comodosAmbientes": [
            {"divisaoInterna": "Sala", "quantidade": "1", "parede": "Pintura", "piso": "Cerâmica", "forroTeto": "Laje"},
            {"divisaoInterna": "Cozinha", "quantidade": "1", "parede": "Azulejo", "piso": "Cerâmica",
             "forroTeto": "Laje"},
            {"divisaoInterna": "Dormitório", "quantidade": "2", "parede": "Pintura", "piso": "Laminado",
             "forroTeto": "Laje"},
            {"divisaoInterna": "Banheiro", "quantidade": "1", "parede": "Azulejo", "piso": "Cerâmica",
             "forroTeto": "PVC"},
            {"divisaoInterna": "Área de serviço", "quantidade": "1", "parede": "Azulejo", "piso": "Cerâmica",
             "forroTeto": "Laje"}]},
        "questoesComplementares": {"dateVistoria": "2026-10-14T13:30:00.000Z"},
        "garantia": {"condGarantia": True},
        # o programa só usa nome e registro; o resto (como na API real) nunca sai daqui
        "engenheiro": {"name": "VISTORIADOR DE EXEMPLO - SP", "creci": "CREA 0000000000",
                       "cpf": "000.000.000-00", "conta": "0000-0"},
        "fotos": _fotos(proposta, [("Fachada", 1), ("Número", 1), ("Rua (baixo)", 1), ("Rua (cima)", 1),
                                   ("Sala", 2), ("Cozinha", 2), ("Dormitório", 2), ("Banheiro", 1),
                                   ("Área de serviço", 1), ("Quintal", 1)]),
    }
    vistoria.update(mudancas)
    return vistoria


def _apartamento(proposta, fotos):
    v = _casa(proposta, fotos=fotos)
    v["imovel"].update(logradouro="Av. Brasil Exemplo", numero="2000", complemento="Apto 52", andar="5",
                       type="Apartamento", matricula="98765")
    v["terreno"].update(area="19.220", fracaoTotal="0,8125", tipoImplantacao="Condomínio")
    v["condominioImovel"].update(
        tipoImovelAvaliado="Apartamento", areaPrivativa="68,40", areaAverbada="112,30", areaComum="43,90",
        vagasCobertas="1", vagasDescobertas="0", numeroAndaresDoImovelAvaliado="0", usoCondominio="Residencial",
        andaresTipo="12", qtdSubsolos="1", unidadesPorAndar="4", elevaores="2", estadoConservacapCondominio="Bom",
        InfraestruturaDoImovel=_lista(("piscina", "Piscina"), ("salaoDeFestas", "Salão de festas"),
                                      ("quadra", "Quadra")))
    v["comodosAmbientes"]["comodosAmbientes"] = [
        {"divisaoInterna": "Sala", "quantidade": "1", "parede": "Pintura", "piso": "Porcelanato", "forroTeto": "Gesso"},
        {"divisaoInterna": "Dormitório", "quantidade": "1", "parede": "Pintura", "piso": "Laminado",
         "forroTeto": "Gesso"},
        {"divisaoInterna": "Dormitório", "quantidade": "1", "parede": "Pintura", "piso": "Laminado",
         "forroTeto": "Gesso"},
        {"divisaoInterna": "Suíte", "quantidade": "1", "parede": "Pintura", "piso": "Laminado", "forroTeto": "Gesso"},
        {"divisaoInterna": "Banheiro Suíte", "quantidade": "1", "parede": "Azulejo", "piso": "Cerâmica",
         "forroTeto": "Gesso"},
        {"divisaoInterna": "Banheiro", "quantidade": "1", "parede": "Azulejo", "piso": "Cerâmica",
         "forroTeto": "Gesso"},
        None,  # cômodo apagado no app: a API devolve null no meio da lista
        {"divisaoInterna": "Cozinha", "quantidade": "1", "parede": "Azulejo", "piso": "Porcelanato",
         "forroTeto": "Gesso"},
        {"divisaoInterna": "Varanda", "quantidade": "1", "parede": "Pintura", "piso": "Cerâmica", "forroTeto": "Laje"}]
    return v


def propostas():
    """{nº da proposta: vistoria} - o que GET /inspection/<_id> devolve."""
    p = {}
    p["100001"] = _casa("100001", fotos=_fotos(
        "100001", [("Fachada", 1), ("Número", 1), ("Rua (baixo)", 1), ("Rua (cima)", 1), ("Sala", 2),
                   ("Cozinha", 2), ("Dormitório", 2), ("Banheiro", 1), ("Área de serviço", 1), ("Quintal", 2)],
        observacoes={"Banheiro": "Infiltração no forro (dado de exemplo)."}))
    p["100002"] = _apartamento("100002", _fotos(
        "100002", [("Fachada", 1), ("Número", 1), ("Rua (baixo)", 1), ("Rua (cima)", 1), ("Portaria", 1),
                   ("Sala", 2), ("Cozinha", 2), ("Dormitório", 2), ("Suíte", 1), ("Banheiro Suíte", 1),
                   ("Banheiro", 1), ("Varanda", 1), ("Piscina", 1), ("Salão de festas", 1), ("Garagem", 1)],
        observacoes={"Banheiro Suíte": "Box quebrado.", "Varanda": "Fechamento em vidro não averbado."}))
    # vistoria antiga: parte das fotos sem "ambiente" (a legenda sai do nome do arquivo)
    p["100003"] = _casa("100003", fotos=_fotos(
        "100003", [("Fachada", 1), ("Placa da rua", 1), ("Rua", 2), ("Sala", 1), ("Cozinha", 1),
                   ("Dormitório", 2), ("Banheiro", 1)], sem_ambiente=("Placa da rua", "Rua")))
    p["100003"]["imovel"].update(logradouro="Travessa do Exemplo", numero="7", cep="01000-001")
    p["100004"] = _casa("100004", fotos=_fotos(
        "100004", [("Fachada", 1), ("Número", 1), ("Rua (baixo)", 1), ("Sala", 1), ("Cozinha", 1),
                   ("Dormitório", 2), ("Banheiro", 1), ("Telhado a reformar", 2)],
        observacoes={"Telhado a reformar": "Reforma prevista: troca do telhado."}))
    p["100004"]["condominioImovel"].update(estadoConservacaoImovel="Regular", idadeAparenteImovelAvaliado="0")
    p["100005"] = _casa("100005", fotos=_fotos(
        "100005", [("Fachada", 1), ("Número", 1), ("Rua (baixo)", 1), ("Rua (cima)", 1), ("Sala", 1),
                   ("Cozinha", 1), ("Dormitório", 1), ("Banheiro", 1), ("Córrego nos fundos", 1)]))
    p["100005"]["regiao"].update(restritivos=_lista(("riscoAmbiental", "Risco ambiental")),
                                 textRestritivo="Córrego nos fundos do terreno")
    # Banco C: 74 fotos (70 na aba principal) - passa das 66 que cabem no modelo e obriga a criar páginas novas
    ambientes = [("Fachada", 1), ("Número", 1), ("Rua (baixo)", 1), ("Rua (cima)", 1)] + [
        (f"{nome} {n}", 3) for n in range(1, 4) for nome in ("Sala", "Cozinha", "Dormitório", "Banheiro",
                                                             "Varanda", "Área comum")] + [("Garagem", 16)]
    p["200001"] = _apartamento("200001", _fotos("200001", ambientes))
    return p


# planilha que o Banco D anexa na requisição da Intranet (FORMULARIO!F5, F6, F7 e F36)
PLANILHAS_BANCO_D = {
    "100001": {"juncao": 1234, "nome_juncao": "AGÊNCIA CENTRO (EXEMPLO)", "tipo": "AQUISIÇÃO SIMPLES"},
    "100002": {"juncao": 2345, "nome_juncao": "AGÊNCIA JARDIM (EXEMPLO)", "tipo": "AQUISIÇÃO A VISTA COM FGTS"},
    "100003": {"juncao": 3456, "nome_juncao": "AGÊNCIA NORTE (EXEMPLO)", "tipo": "CONSÓRCIO (IMÓVEL URBANO)"},
    # escrito diferente do modelo (travessão, sem acento): tem que achar o modelo do mesmo jeito
    "100004": {"juncao": 4567, "nome_juncao": "AGÊNCIA SUL (EXEMPLO)",
               "tipo": "CONSORCIO – REFORMA S/AMPLIACAO (IMOVEL URBANO)"},
    "100005": {"juncao": 5678, "nome_juncao": "AGÊNCIA LESTE (EXEMPLO)", "tipo": "GARANTIA (IMÓVEL URBANO)"},
}
