"""
Fluxo de uma proposta, em duas etapas (a janela mostra as fotos entre elas):
  1. buscar(): acha a proposta e baixa TODAS as fotos pra um cache temporário;
  2. gerar(): copia só as fotos escolhidas pra <destino>/Proposta_<nº>/Fotos e
     monta o Excel do banco (nome no perfil, ex.: Banco C ->
     planilha_<nº>_<tipo>.xlsm) - fotos nos quadros da aba de fotos locais e
     as demais na aba principal.

O que muda de banco pra banco vem do perfil (perfis.py).

executar() faz as duas de uma vez (todas as fotos + sugestão automática dos
quadros locais) - usado pelo terminal:
  python robo.py 123456 --destino "C:\\Laudos" --banco "Banco C" --modelo Comparativo
"""

import argparse
import glob
import os
import re
import shutil
import tempfile
import unicodedata

from central_gestao import buscar_fotos, ler_dados, nome_arquivo
from comum import PASTA_RECURSOS, PASTA_SCRIPT, PREENCHIMENTO, VERSAO, fonte_login
from excel_fotos import NIVEIS_COMPRESSAO, inserir_fotos, listar_imagens
from perfis import BANCO_C, PADRAO, PERFIS


def nome_exibido(arquivo):
    """'modelo_comparativo.xlsm' -> 'Comparativo'."""
    nome = os.path.splitext(os.path.basename(arquivo))[0]
    nome = re.sub(r"^modelo[_\s-]*", "", nome, flags=re.I).replace("_", " ").strip()
    return nome[:1].upper() + nome[1:]


def listar_modelos(perfil=BANCO_C):
    """{nome_exibido: caminho} dos .xlsm do banco: pasta modelos/<banco> ao
    lado do programa (tem prioridade - serve pra trocar modelo sem gerar o
    .exe de novo) e os embutidos no .exe."""
    modelos = {}
    for raiz in (PASTA_SCRIPT, PASTA_RECURSOS):
        pasta = os.path.join(raiz, "modelos", perfil.pasta_modelos)
        for caminho in sorted(glob.glob(os.path.join(pasta, "*.xlsm")) + glob.glob(os.path.join(pasta, "*.xlsx"))):
            if not os.path.basename(caminho).startswith("~$"):
                modelos.setdefault(nome_exibido(caminho), caminho)
    return modelos


def _chave_tipo(texto):
    """'CONSÓRCIO – REFORMA S/AMPLIAÇÃO' e 'consorcio - reforma s/ampliacao' dão a mesma chave."""
    t = "".join(c for c in unicodedata.normalize("NFD", str(texto or "")) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", t.lower()).strip()


_tipos_dos_modelos = {}  # {caminho do modelo: chave do tipo de operação (F36)}


def modelo_por_tipo(tipo, perfil):
    """Nome do modelo (como aparece na tela) cujo tipo de operação, escrito
    no próprio modelo (pre_preenchimento.aba!F36), é `tipo`; ou None."""
    if not tipo or not perfil.pre_preenchimento:
        return None
    from intranet import _celulas_xlsx

    alvo = _chave_tipo(tipo)
    for nome, caminho in listar_modelos(perfil).items():
        if caminho not in _tipos_dos_modelos:
            try:
                with open(caminho, "rb") as f:
                    lido = _celulas_xlsx(f.read(), perfil.pre_preenchimento.aba, ["F36"])["F36"]
            except Exception:
                lido = None
            _tipos_dos_modelos[caminho] = _chave_tipo(lido)
        if _tipos_dos_modelos[caminho] and _tipos_dos_modelos[caminho] == alvo:
            return nome
    return None


def descobrir_modelo(proposta, perfil, log=print):
    """Tipo de operação da proposta -> modelo do Excel, pelas fontes em ordem:
      1. planilha diária do Banco D na pasta pública (a fazer - aguardando
         o caminho/formato da planilha);
      2. Intranet: planilha do banco anexada na requisição (intranet.py).
    Devolve {"modelo", "tipo_operacao", "fonte", ...} - "modelo" None se o
    tipo não tiver modelo - ou levanta ValueError com o motivo."""
    if not perfil.pre_preenchimento:
        raise ValueError(f"{perfil.banco}: o modelo não depende do tipo de operação")
    import intranet

    motivos = []
    for fonte, buscar_dados in (("Intranet", intranet.dados_do_banco),):
        try:
            dados = buscar_dados(proposta, log=log)
        except Exception as e:
            motivos.append(f"{fonte}: {e}")
            continue
        dados.update(fonte=fonte, modelo=modelo_por_tipo(dados.get("tipo_operacao"), perfil))
        return dados
    raise ValueError("; ".join(motivos))


def limpar_proposta(texto):
    """Só os dígitos (aceita colar "200.001" ou " 200001 ")."""
    return re.sub(r"\D", "", texto or "")


def caminhos_saida(proposta, destino, modelo, perfil=BANCO_C):
    """(pasta da proposta, pasta das fotos, Excel). O nome do Excel vem do
    perfil (com o tipo do modelo, se o perfil usar) e a extensão do próprio
    modelo: Banco C -> planilha_123_comparativo.xlsm; Banco D -> 123_FORMULARIO.xlsx."""
    tipo = re.sub(r"\W+", "_", nome_exibido(modelo).lower()).strip("_")
    pasta = os.path.join(destino, f"Proposta_{proposta}")
    return pasta, os.path.join(pasta, "Fotos"), os.path.join(
        pasta, perfil.nome_arquivo.format(proposta=proposta, tipo=tipo) + os.path.splitext(modelo)[1].lower())


# o Excel não abre/salva arquivo com caminho maior que ~218 caracteres, e o
# Windows trava em 260 pra qualquer arquivo (visto em 08/10/2026 com pasta funda)
MAX_CAMINHO_EXCEL = 210
MAX_CAMINHO_FOTO = 245

# o Excel gerado não pode passar disso (pedido em 08/10/2026: "14 mega");
# em bytes decimais, que é o menor dos dois jeitos de contar MB
LIMITE_EXCEL_BYTES = 14_000_000


def _caminho_foto(pasta, prefixo, legenda, ext):
    """Nome "prefixo - legenda.ext", encurtando a legenda se o caminho
    completo passar do limite do Windows."""
    legenda = legenda.strip()[:60]  # mesmo corte do nome_arquivo, pra conta do excesso fechar
    caminho = os.path.join(pasta, nome_arquivo(prefixo, legenda) + ext)
    excesso = len(caminho) - MAX_CAMINHO_FOTO
    if excesso > 0:
        legenda = legenda[:max(8, len(legenda) - excesso - 1)].rstrip() + "…"
        caminho = os.path.join(pasta, nome_arquivo(prefixo, legenda) + ext)
    return caminho


def buscar(proposta, log=print, progresso=None):
    """Etapa 1. Devolve (proposta, itens, pasta_cache). Ver central_gestao.buscar_fotos."""
    progresso = progresso or (lambda *_: None)
    proposta = limpar_proposta(proposta)
    if not proposta:
        raise ValueError("Informe o número da proposta (só números).")
    log(f"[INÍCIO] Proposta {proposta} | versão {VERSAO} (preenchimento {PREENCHIMENTO}) | login: {fonte_login()}")
    progresso(0.02, "Buscando a proposta na Central de Gestão...")
    cache = tempfile.mkdtemp(prefix=f"fotos_{proposta}_")
    try:
        itens = buscar_fotos(
            proposta, cache, log=log,
            progresso=lambda feitas, total: progresso(0.05 + 0.95 * feitas / max(total, 1),
                                                       f"Baixando fotos ({feitas}/{total})..."))
    except Exception:
        shutil.rmtree(cache, ignore_errors=True)
        raise
    return proposta, itens, cache


EXTENSOES_FOTO_PC = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff")


def adicionar_fotos_pc(itens, cache, arquivos, guardar=None):
    """Fotos escolhidas no computador entram na busca como as da Central de Gestão:
    copiadas pro cache (a pasta da pessoa nunca é mexida) e acrescentadas a
    `itens` com índices novos; a legenda é o nome do arquivo. Não manda nada
    pra Central de Gestão. `guardar(arquivo)` (opcional) faz a cópia que fica no
    trabalho salvo e devolve o caminho dela - o nome dela vira a chave da foto
    ("pc:<nome>"). Devolve (itens novos, [(arquivo, motivo)] recusados)."""
    from PIL import Image

    proximo = max((it["indice"] for it in itens), default=-1) + 1
    novos, recusados = [], []
    for arquivo in arquivos:
        try:
            with Image.open(arquivo) as img:
                img.verify()  # garante que é imagem de verdade (e não está corrompida)
        except Exception:
            recusados.append((arquivo, "não é uma imagem que o programa consegue abrir"))
            continue
        ext = os.path.splitext(arquivo)[1].lower()
        destino = os.path.join(cache, f"pc_{proximo + 1:03d}{ext if ext in EXTENSOES_FOTO_PC else '.jpg'}")
        shutil.copyfile(arquivo, destino)
        legenda = os.path.splitext(os.path.basename(arquivo))[0].strip() or "Foto do PC"
        copia = guardar(arquivo) if guardar else arquivo
        novo = {"indice": proximo, "legenda": legenda, "caminho": destino, "erro": None, "origem": "pc",
                "chave": "pc:" + os.path.basename(copia)}
        itens.append(novo)
        novos.append(novo)
        proximo += 1
    return novos, recusados


def dados_da_busca(cache):
    """Dados da proposta (protocolo, proponente, ...) gravados pela buscar()."""
    return ler_dados(cache)


def sugerir_locais(itens, perfil=BANCO_C):
    """{chave do quadro: indice} pela legenda; cada foto vai pra um quadro só."""
    usados, sugestao = set(), {}
    for q in (perfil.locais.quadros if perfil.locais else ()):
        if not q.sugestao:
            continue
        for it in itens:
            if it["caminho"] and it["indice"] not in usados and re.search(q.sugestao, it["legenda"], re.I):
                sugestao[q.chave] = it["indice"]
                usados.add(it["indice"])
                break
    return sugestao


def gerar(proposta, destino, modelo, itens, internas, locais, log=print, progresso=None, legendas=None,
          perfil=BANCO_C, dados=None):
    """Etapa 2. `internas` = índices (na ordem) que vão pra aba principal;
    `locais` = {chave do quadro: índice} - chaves dos quadros fixos do perfil
    ou "extra:<nome da seção>" (quadro novo, na ordem do dicionário);
    `legendas` = {índice: texto} pra fotos renomeadas na janela; `dados` =
    dados da proposta (dados_da_busca) pro pré-preenchimento. Devolve o
    caminho do Excel."""
    progresso = progresso or (lambda *_: None)
    legendas = legendas or {}
    rotulos = perfil.rotulos_locais()
    por_indice = {it["indice"]: dict(it, legenda=(legendas.get(it["indice"]) or it["legenda"]).strip())
                  for it in itens}
    if not internas and not locais:
        raise ValueError("Selecione pelo menos uma foto.")
    pasta, pasta_fotos, excel_saida = caminhos_saida(proposta, destino, modelo, perfil)
    if len(os.path.abspath(excel_saida)) > MAX_CAMINHO_EXCEL:
        raise ValueError(f"O caminho da pasta escolhida é longo demais para o Excel "
                         f"({len(os.path.abspath(excel_saida))} caracteres; o máximo é {MAX_CAMINHO_EXCEL}).\n\n"
                         f"Escolha em “Salvar em” uma pasta mais curta (ex.: Documentos\\Fotos Propostas).")

    # pasta da proposta: só as fotos escolhidas (uma nova geração substitui a anterior)
    os.makedirs(pasta_fotos, exist_ok=True)
    for antiga in listar_imagens(pasta_fotos):
        os.remove(antiga)
    fotos_locais, extras = {}, []
    for chave, idx in locais.items():
        it = por_indice[idx]
        secao = rotulos.get(chave) or chave.split(":", 1)[1]
        destino_foto = _caminho_foto(pasta_fotos, f"Local - {secao}", it["legenda"],
                                     os.path.splitext(it["caminho"])[1])
        shutil.copyfile(it["caminho"], destino_foto)
        if chave in rotulos:
            fotos_locais[chave] = destino_foto
        else:
            extras.append((secao, destino_foto))
    fotos_internas, descricoes = [], []
    for n, idx in enumerate(internas, 1):
        it = por_indice[idx]
        destino_foto = _caminho_foto(pasta_fotos, f"{n:02d}", it["legenda"], os.path.splitext(it["caminho"])[1])
        shutil.copyfile(it["caminho"], destino_foto)
        fotos_internas.append(destino_foto)
        descricoes.append(it["legenda"])
    aba_locais = perfil.locais.nome if perfil.locais else "-"
    log(f"[FOTOS] {perfil.banco}: {len(fotos_internas)} na aba {perfil.fotos.nome} + "
        f"{len(fotos_locais) + len(extras)} na aba {aba_locais} ({len(extras)} em seções novas), "
        f"salvas em {pasta_fotos}")

    progresso(0.05, "Montando o Excel...")
    total = len(fotos_internas)
    tmp = tempfile.mkdtemp(prefix="fotos_red_")
    try:
        def log_excel(msg):
            log(msg)
            m = re.match(r"\[EXCEL\] Foto (\d+)/(\d+)", msg)
            if m:
                progresso(0.1 + 0.85 * int(m.group(1)) / total, f"Inserindo no Excel ({m.group(1)}/{total})...")

        for nivel, compressao in enumerate(NIVEIS_COMPRESSAO):
            inserir_fotos(modelo, excel_saida, fotos_internas, tmp, descricoes, log=log_excel, locais=fotos_locais,
                          extras=extras, perfil=perfil, compressao=compressao, dados=dados)
            tamanho = os.path.getsize(excel_saida)
            log(f"[EXCEL] Tamanho: {tamanho / 1_000_000:.1f} MB (limite {LIMITE_EXCEL_BYTES / 1_000_000:g} MB, "
                f"compressão nível {nivel + 1})")
            if tamanho <= LIMITE_EXCEL_BYTES:
                break
            if nivel + 1 < len(NIVEIS_COMPRESSAO):
                log("[EXCEL] Passou do limite: refazendo com as fotos mais comprimidas...")
                progresso(0.05, "Comprimindo mais as fotos para caber no limite...")
        else:
            log(f"[EXCEL] [AVISO] Mesmo na compressão máxima ficou com {tamanho / 1_000_000:.1f} MB.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    progresso(1.0, "Pronto!")
    log(f"[FIM] Excel gerado: {excel_saida}")
    return excel_saida


def falhas_de(itens):
    return [f"foto {it['indice'] + 1} ({it['legenda']}): {it['erro']}" for it in itens if it["erro"]]


def executar(proposta, destino, modelo, log=print, progresso=None, perfil=BANCO_C, dados_banco=None):
    """Tudo de uma vez, sem escolha: todas as fotos que baixaram, com os
    quadros locais sugeridos pela legenda. `dados_banco` = o que veio da
    planilha do banco (descobrir_modelo: junção...). Devolve (Excel, falhas)."""
    proposta, itens, cache = buscar(proposta, log=log)
    try:
        locais = sugerir_locais(itens, perfil)
        internas = [it["indice"] for it in itens if it["caminho"] and
                    (perfil.quadros_tambem_na_principal or it["indice"] not in locais.values())]
        excel = gerar(proposta, destino, modelo, itens, internas, locais, log=log, progresso=progresso,
                      perfil=perfil, dados=dict(dados_da_busca(cache), banco=dados_banco or {}))
    finally:
        shutil.rmtree(cache, ignore_errors=True)
    falhas = falhas_de(itens)
    if falhas:
        log(f"[FIM] [AVISO] {len(falhas)} foto(s) ficaram de fora (ver [ERRO] acima).")
    return excel, falhas


def main():
    ap = argparse.ArgumentParser(description="Baixa as fotos de uma proposta e monta o Excel.")
    ap.add_argument("proposta")
    ap.add_argument("--destino", required=True, help="pasta onde criar Proposta_<nº>")
    ap.add_argument("--banco", default=PADRAO, choices=list(PERFIS))
    ap.add_argument("--modelo", default=None, help="nome do modelo (ex.: Comparativo); padrão: o 1º do banco")
    args = ap.parse_args()
    perfil = PERFIS[args.banco]
    modelos = listar_modelos(perfil)
    _, falhas = executar(args.proposta, args.destino, modelos[args.modelo or next(iter(modelos))], perfil=perfil)
    raise SystemExit(2 if falhas else 0)


if __name__ == "__main__":
    main()
