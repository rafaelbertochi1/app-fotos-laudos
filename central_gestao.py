"""
Download das fotos de uma proposta direto pela API da Central de Gestão - sem
abrir navegador.

Mapeado em 07/10/2026 (ferramentas/mapear_central_gestao.py + código do site, main.*.chunk.js):
  - API: https://api.central-gestao.example.com
  - login: POST /auth {"email", "password"} com o cabeçalho x-api-key (chave
    pública que o próprio site manda) -> {"auth": true, "token": ...}
  - demais chamadas: cabeçalho  authorization: "Barear <token>"  (sic, é
    assim que o site escreve). GET /auth/validToken -> "VALID TOKEN".
  - GET /inspection?id=<proposta> -> {"result": [...], "totalFound": N}
    (busca exata; só aceita número - texto dá erro 500). No resultado,
    "fotos" é só o id; as fotos vêm de GET /inspection/<_id>:
    fotos.fotos[] = {"ambiente": "Cozinha", "images": ["public/images/..."]}
  - as imagens são públicas: GET <API>/public/images/... sem token. Vêm em
    JPEG mesmo com extensão .png.
  - dados da proposta (pré-preenchimento): o mesmo GET /inspection/<_id> ->
    imovel, regiao, terreno, condominioImovel, comodosAmbientes, garantia,
    questoesComplementares.dateVistoria (UTC), _createdAt (UTC). Junção e
    tipo de operação NÃO existem na API. Ver formulario.py.
"""

import io
import json
import os
import re
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import quote

from PIL import Image

from comum import PASTA_DADOS, SESSION_FILE, credenciais_configuradas

# a variável serve pro servidor de exemplo dos testes (testes/servidor_exemplo.py)
API = os.environ.get("CENTRAL_GESTAO_API") or "https://api.central-gestao.example.com"
API_KEY_LOGIN = "00000000-0000-0000-0000-000000000000"  # pública, vem no JS do site
TOKEN_FILE = os.path.join(PASTA_DADOS, "token_central_gestao.json")
DOWNLOADS_PARALELOS = 8
TENTATIVAS = 3


class ErroCentralGestao(Exception):
    """Erro com mensagem já pronta pra mostrar ao usuário."""


class ErroLogin(ErroCentralGestao):
    pass


# ---------------------------------------------------------------- HTTP

def _requisitar(caminho, token=None, dados=None, cabecalhos=None, timeout=60):
    """Faz a chamada e devolve (status, bytes). Erros HTTP não levantam
    exceção (o chamador decide); erros de rede tentam de novo."""
    url = caminho if caminho.startswith("http") else API + caminho
    h = {"Content-Type": "application/json", "User-Agent": "DownloaderFotos/1.0"}
    if token:
        h["authorization"] = "Barear " + token
    h.update(cabecalhos or {})
    corpo = json.dumps(dados).encode() if dados is not None else None
    ultimo_erro = None
    for tentativa in range(TENTATIVAS):
        try:
            req = urllib.request.Request(url, data=corpo, headers=h, method="POST" if corpo else "GET")
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.status, r.read()
        except urllib.error.HTTPError as e:
            if e.code < 500 or tentativa == TENTATIVAS - 1:
                return e.code, e.read()
            ultimo_erro = e
        except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
            ultimo_erro = e
        time.sleep(1.5 * (tentativa + 1))
    raise ErroCentralGestao(f"Sem conexão com a Central de Gestão ({ultimo_erro}). Verifique a internet.")


# ---------------------------------------------------------------- token

def _token_salvo():
    try:
        with open(TOKEN_FILE, encoding="utf-8") as f:
            return json.load(f).get("token")
    except Exception:
        pass
    # sessão salva pelo mapeador (navegador): o site guarda o token no localStorage
    try:
        with open(SESSION_FILE, encoding="utf-8") as f:
            for origem in json.load(f).get("origins", []):
                for item in origem.get("localStorage", []):
                    if item["name"] == "@GLOBAL_T":
                        return json.loads(item["value"])
    except Exception:
        pass
    return None


def _salvar_token(token):
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        json.dump({"token": token}, f)


def apagar_token():
    try:
        os.remove(TOKEN_FILE)
    except FileNotFoundError:
        pass


def _token_valido(token):
    status, corpo = _requisitar("/auth/validToken", token=token, timeout=30)
    return status < 300 and b"VALID TOKEN" in corpo


def obter_token(log=print):
    """Token salvo se ainda valer; senão faz login com o usuário/senha
    gravados no PC e salva o token novo. Nunca registra a senha."""
    token = _token_salvo()
    if token and _token_valido(token):
        return token

    usuario, senha = credenciais_configuradas()
    if not usuario:
        raise ErroLogin("O login da Central de Gestão não está configurado. Clique em 'Configurar login…'.")
    log("[LOGIN] Entrando na Central de Gestão...")
    status, corpo = _requisitar("/auth", dados={"email": usuario, "password": senha},
                                cabecalhos={"x-api-key": API_KEY_LOGIN}, timeout=30)
    try:
        resposta = json.loads(corpo)
    except ValueError:
        resposta = {}
    token = resposta.get("token") if isinstance(resposta, dict) and resposta.get("auth") else None
    if not token:
        raise ErroLogin("A Central de Gestão recusou o usuário/senha gravados. "
                        "Confira em 'Configurar login…'.")
    _salvar_token(token)
    log("[LOGIN] OK.")
    return token


def _get_json(caminho, token):
    status, corpo = _requisitar(caminho, token=token)
    if status == 401:
        apagar_token()
        raise ErroLogin("A sessão da Central de Gestão expirou. Tente de novo.")
    if status >= 400:
        raise ErroCentralGestao(f"A Central de Gestão respondeu erro {status} em {caminho.split('?')[0]}.")
    return json.loads(corpo)


# ---------------------------------------------------------------- fotos

def buscar_vistoria(proposta, token, log=print):
    """Devolve o JSON completo da vistoria da proposta (com as fotos)."""
    if not str(proposta).isdigit():
        raise ErroCentralGestao("O número da proposta deve ter só dígitos.")
    busca = _get_json(f"/inspection?id={proposta}", token)
    achadas = [v for v in busca.get("result", []) if str(v.get("id")) == str(proposta)]
    if not achadas:
        raise ErroCentralGestao(f"A proposta {proposta} não foi encontrada na Central de Gestão.")
    if len(achadas) > 1:
        log(f"[CENTRAL] [AVISO] {len(achadas)} vistorias para a proposta {proposta}; "
            "usando a mais recente com fotos.")

    detalhes = [_get_json(f"/inspection/{v['_id']}", token) for v in achadas]
    com_fotos = [d for d in detalhes if _itens_foto(d)]
    if not com_fotos:
        raise ErroCentralGestao(f"A proposta {proposta} ainda não tem fotos na Central de Gestão.")
    return max(com_fotos, key=lambda d: d.get("_createdAt") or "")


def _itens_foto(vistoria):
    fotos = vistoria.get("fotos")
    return fotos.get("fotos", []) if isinstance(fotos, dict) else []


def listar_fotos(vistoria):
    """[(url, legenda, observação)] na ordem da Central de Gestão. A observação é a
    do ambiente ("Observações:" ao abrir a foto na Central de Gestão - campo
    fotos.fotos[].observacao) e vale pra todas as fotos dele; "" se não tem."""
    lista = []
    for item in _itens_foto(vistoria):
        imagens = item.get("images") or []
        observacao = re.sub(r"\s+", " ", str(item.get("observacao") or "")).strip()
        for n, caminho in enumerate(imagens):
            # vistorias antigas vêm sem "ambiente": usa o começo do nome do
            # arquivo (".../PlacadaRua_2025...[0].png" -> "PlacadaRua")
            legenda = ((item.get("ambiente") or "").strip()
                       or os.path.basename(caminho).split("_")[0].split("[")[0] or "Foto")
            sufixo = f" ({n + 1})" if len(imagens) > 1 else ""
            lista.append((f"{API}/{quote(caminho.lstrip('/'))}", legenda + sufixo, observacao))
    return lista


def _baixar_uma(url, destino_sem_ext):
    status, corpo = _requisitar(url, timeout=120)
    if status == 404:
        raise ErroCentralGestao("o arquivo não existe mais no servidor (404)")
    if status >= 400 or not corpo:
        raise ErroCentralGestao(f"erro {status}")
    try:
        formato = (Image.open(io.BytesIO(corpo)).format or "jpeg").lower()
    except Exception:
        raise ErroCentralGestao("o arquivo baixado não é uma imagem")
    ext = {"jpeg": ".jpg", "png": ".png", "webp": ".webp", "gif": ".gif", "bmp": ".bmp"}.get(formato, ".jpg")
    caminho = destino_sem_ext + ext
    with open(caminho, "wb") as f:
        f.write(corpo)
    return caminho


# ---------------------------------------------------------------- dados da proposta

ARQ_DADOS = "dados_proposta.json"  # na pasta de cache da busca


def dados_proposta(vistoria):
    """A parte da vistoria que serve pro pré-preenchimento do formulário do
    banco (formulario.py traduz). Do engenheiro, só nome e registro - o
    resto (CPF, conta, endereço) não sai daqui."""
    comodos = (vistoria.get("comodosAmbientes") or {})
    eng = vistoria.get("engenheiro") or {}
    return {
        "id": vistoria.get("id"),
        "_createdAt": vistoria.get("_createdAt"),
        "date": vistoria.get("date"),
        "observacoes": vistoria.get("observacoes"),
        "latitude": vistoria.get("latitude"),
        "longitude": vistoria.get("longitude"),
        "imovel": vistoria.get("imovel") if isinstance(vistoria.get("imovel"), dict) else {},
        "regiao": vistoria.get("regiao") if isinstance(vistoria.get("regiao"), dict) else {},
        "terreno": vistoria.get("terreno") if isinstance(vistoria.get("terreno"), dict) else {},
        "condominioImovel": (vistoria.get("condominioImovel")
                             if isinstance(vistoria.get("condominioImovel"), dict) else {}),
        # a lista pode trazer itens vazios (null) - cômodo apagado no app (proposta 100006)
        "comodos": [c for c in (comodos.get("comodosAmbientes") or []) if isinstance(c, dict)]
                   if isinstance(comodos, dict) else [],
        "questoesComplementares": {"dateVistoria": (vistoria.get("questoesComplementares") or {}).get("dateVistoria")
                                   if isinstance(vistoria.get("questoesComplementares"), dict) else None},
        "garantia": {"condGarantia": (vistoria.get("garantia") or {}).get("condGarantia")
                     if isinstance(vistoria.get("garantia"), dict) else None},
        "engenheiro": {"name": eng.get("name"), "creci": eng.get("creci")} if isinstance(eng, dict) else {},
    }


def ler_dados(pasta_cache):
    """dados_proposta() gravado pela busca ({} se não houver)."""
    try:
        with open(os.path.join(pasta_cache, ARQ_DADOS), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError, TypeError):
        return {}


def buscar_fotos(proposta, pasta_cache, log=print, progresso=None):
    """Acha a proposta e baixa TODAS as fotos dela pra `pasta_cache` (pasta
    temporária: a pessoa ainda vai escolher quais usar). Devolve a lista na
    ordem da Central de Gestão: [{"indice", "legenda", "caminho" (ou None), "erro",
    "observacao", "chave"}].
    Grava também os dados da proposta (ARQ_DADOS, ver ler_dados).
    `progresso(feitas, total)` é chamado a cada foto."""
    progresso = progresso or (lambda *_: None)
    token = obter_token(log)
    try:
        vistoria = buscar_vistoria(proposta, token, log)
    except ErroLogin:
        token = obter_token(log)  # token caiu entre a validação e a busca
        vistoria = buscar_vistoria(proposta, token, log)

    fotos = listar_fotos(vistoria)
    total = len(fotos)
    os.makedirs(pasta_cache, exist_ok=True)
    with open(os.path.join(pasta_cache, ARQ_DADOS), "w", encoding="utf-8") as f:
        json.dump(dados_proposta(vistoria), f, ensure_ascii=False)
    log(f"[CENTRAL] {total} foto(s) na proposta {proposta}. Baixando...")
    progresso(0, total)

    # "chave" = endereço da foto: identifica a mesma foto numa busca futura (trabalho.py)
    itens = [{"indice": i, "legenda": legenda, "caminho": None, "erro": None, "observacao": obs, "chave": url}
             for i, (url, legenda, obs) in enumerate(fotos)]
    com_obs = list(dict.fromkeys((it["legenda"].rsplit(" (", 1)[0], it["observacao"])
                                 for it in itens if it["observacao"]))
    log(f"[CENTRAL] Observações nas fotos: {len(com_obs)}"
        + "".join(f"\n    - {amb}: {obs}" for amb, obs in com_obs))
    with ThreadPoolExecutor(DOWNLOADS_PARALELOS) as pool:
        tarefas = {pool.submit(_baixar_uma, url, os.path.join(pasta_cache, f"{i + 1:03d}")): i
                   for i, (url, _, _) in enumerate(fotos)}
        for feitas, tarefa in enumerate(as_completed(tarefas), 1):
            i = tarefas[tarefa]
            try:
                itens[i]["caminho"] = tarefa.result()
            except Exception as e:
                itens[i]["erro"] = str(e)
                log(f"[CENTRAL] [ERRO] Não baixou a foto {i + 1} ({itens[i]['legenda']}): {e}")
            progresso(feitas, total)

    if not any(it["caminho"] for it in itens):
        raise ErroCentralGestao(f"Nenhuma das {total} fotos da proposta {proposta} pôde ser baixada"
                             f" - {itens[0]['erro']}.")
    return itens


def nome_arquivo(prefixo, legenda):
    """Nome final da foto na pasta da proposta (sem extensão)."""
    limpo = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", legenda).strip(" .") or "Foto"
    return f"{prefixo} - {limpo[:60]}"
