"""
Tipo de operação de uma proposta do Banco D pela Intranet da empresa
(https://intranet.example.com) - e, com ele, o modelo do Excel certo.

Cada proposta do Banco D tem uma requisição no fluxo "Laudos BANCO D", e
nos "Arquivos ativos" dela está a planilha que o Banco D manda já
preenchida ("Formulario_Banco_D - <data>.xlsx"), com o tipo
de operação em FORMULARIO!F36 (e junção, protocolo...). Mapeado em
14/10/2026 no Chrome do Rafael:

  - login em 2 etapas: POST /login/username {username} -> página
    com o formulário da senha -> POST dele (cookie de sessão)
  - GET /portal/requisicao/busca?descricao=<proposta> -> link
    /portal/requisicao/entrada/detalhar/<id> (título "<proposta> - ...")
  - GET do detalhe -> link do anexo /portal/requisicao/download/id/<n>
  - GET do download -> o .xlsx

É o "segundo fator" (pedido do Rafael): o primeiro vai ser a planilha
diária do Banco D na pasta pública. Só lê - nada é gravado na Intranet.
Login guardado só neste PC, criptografado pelo Windows (DPAPI), como o da
Central de Gestão; a sessão (cookie) também, pra não entrar a cada proposta.
"""

import json
import os
import re
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from html import unescape
from html.parser import HTMLParser
from http.cookiejar import Cookie, CookieJar
from io import BytesIO

from comum import PASTA_DADOS

# a variável serve pro servidor de exemplo dos testes (testes/servidor_exemplo.py)
BASE = os.environ.get("INTRANET_URL") or "https://intranet.example.com"
ARQ_LOGIN = os.path.join(PASTA_DADOS, "intranet.dat")
ARQ_SESSAO = os.path.join(PASTA_DADOS, "sessao_intranet.dat")
TIMEOUT = 30
NOME_PLANILHA_BANCO = "formulario_banco_d"  # começo do nome do anexo (sem acento/caixa)


class ErroIntranet(Exception):
    """Mensagem já pronta pra mostrar/logar."""


class ErroLoginIntranet(ErroIntranet):
    pass


# ---------------------------------------------------------------- login guardado (DPAPI)

def _proteger(dados):
    import win32crypt

    return win32crypt.CryptProtectData(json.dumps(dados).encode("utf-8"), "DownloaderFotos", None, None, None, 0)


def _abrir(blob):
    import win32crypt

    return json.loads(win32crypt.CryptUnprotectData(blob, None, None, None, 0)[1])


def ler_login():
    """(usuario, senha) gravados pela janela, senão INTRANET_USUARIO/SENHA; ou (None, None)."""
    try:
        with open(ARQ_LOGIN, "rb") as f:
            d = _abrir(f.read())
        if d.get("usuario") and d.get("senha"):
            return d["usuario"], d["senha"]
    except (OSError, ValueError, ImportError):
        pass
    except Exception:  # pywintypes.error: arquivo de outro usuário/PC
        pass
    usuario, senha = os.environ.get("INTRANET_USUARIO", "").strip(), os.environ.get("INTRANET_SENHA", "")
    return (usuario, senha) if usuario and senha else (None, None)


def gravar_login(usuario, senha):
    tmp = ARQ_LOGIN + ".tmp"
    with open(tmp, "wb") as f:
        f.write(_proteger({"usuario": usuario, "senha": senha}))
    os.replace(tmp, ARQ_LOGIN)
    apagar_sessao()
    if ler_login() != (usuario, senha):
        raise RuntimeError("o login foi gravado mas não pôde ser lido de volta")


def apagar_login():
    for caminho in (ARQ_LOGIN, ARQ_SESSAO):
        try:
            os.remove(caminho)
        except FileNotFoundError:
            pass


def apagar_sessao():
    try:
        os.remove(ARQ_SESSAO)
    except FileNotFoundError:
        pass


def usuario_configurado():
    return ler_login()[0]


# ---------------------------------------------------------------- HTML

class _Links(HTMLParser):
    """[(href, texto)] de todos os <a>, e os <form> com seus <input>."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.links, self.forms = [], []
        self._a = None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "a":
            self._a = [a.get("href") or "", ""]
        elif tag == "form":
            self.forms.append({"action": a.get("action") or "", "inputs": []})
        elif tag == "input" and self.forms:
            self.forms[-1]["inputs"].append(a)

    def handle_data(self, data):
        if self._a is not None:
            self._a[1] += data

    def handle_endtag(self, tag):
        if tag == "a" and self._a is not None:
            self.links.append((self._a[0], re.sub(r"\s+", " ", self._a[1]).strip()))
            self._a = None


def _analisar(html):
    p = _Links()
    p.feed(html)
    return p


def _chave(texto):
    t = "".join(c for c in unicodedata.normalize("NFD", str(texto or "")) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", "_", t.lower()).strip("_")


# ---------------------------------------------------------------- sessão HTTP

class Sessao:
    def __init__(self):
        self.jar = CookieJar()
        self._carregar_cookies()
        self.op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.jar))
        self.op.addheaders = [("User-Agent", "Mozilla/5.0 (DownloaderFotos)")]

    # cookies guardados (DPAPI) - evita fazer login a cada proposta
    def _carregar_cookies(self):
        try:
            with open(ARQ_SESSAO, "rb") as f:
                for c in _abrir(f.read()):
                    self.jar.set_cookie(Cookie(0, c["name"], c["value"], None, False, c["domain"], True,
                                               c["domain"].startswith("."), c["path"], True, c["secure"],
                                               c["expires"], False, None, None, {}))
        except Exception:
            pass

    def _gravar_cookies(self):
        try:
            cookies = [{"name": c.name, "value": c.value, "domain": c.domain, "path": c.path, "secure": c.secure,
                        "expires": c.expires} for c in self.jar
                       if (urllib.parse.urlparse(BASE).hostname or "") in c.domain]
            with open(ARQ_SESSAO, "wb") as f:
                f.write(_proteger(cookies))
        except Exception:
            pass

    def _get(self, caminho, dados=None):
        url = caminho if caminho.startswith("http") else BASE + caminho
        corpo = urllib.parse.urlencode(dados).encode() if dados is not None else None
        # consulta (GET) que der erro do servidor (502/503...) ou de rede tenta de novo: o
        # Intranet às vezes dá 502 passageiro (visto no teste de 15/10/2026). Envio de
        # formulário (login) não repete.
        tentativas = 3 if corpo is None else 1
        for n in range(tentativas):
            try:
                with self.op.open(url, data=corpo, timeout=TIMEOUT) as r:
                    return r.geturl(), r.read()
            except urllib.error.HTTPError as e:
                if e.code < 500 or n == tentativas - 1:
                    raise ErroIntranet(f"a Intranet respondeu erro {e.code} em {caminho.split('?')[0]}")
            except (urllib.error.URLError, OSError) as e:
                if n == tentativas - 1:
                    raise ErroIntranet(f"sem conexão com a Intranet ({getattr(e, 'reason', e)})")
            time.sleep(1.5 * (n + 1))

    @staticmethod
    def _na_tela_de_login(url):
        return "/login" in urllib.parse.urlparse(url).path

    def entrar(self, usuario, senha):
        """Login em 2 etapas (usuário, depois senha)."""
        self.jar.clear()
        url, html = self._get("/login")
        form = next((f for f in _analisar(html.decode("utf-8", "replace")).forms if "username" in f["action"]), None)
        if not form:
            raise ErroIntranet("a tela de login da Intranet mudou (formulário do usuário não encontrado)")
        dados = {i["name"]: i.get("value", "") for i in form["inputs"] if i.get("name")}
        dados["username"] = usuario
        url, html = self._get(unescape(form["action"]), dados)
        p = _analisar(html.decode("utf-8", "replace"))
        form = next((f for f in p.forms if any(i.get("type") == "password" for i in f["inputs"])), None)
        if not form:
            raise ErroLoginIntranet("a Intranet não aceitou o usuário (não pediu a senha) - confira o usuário")
        dados = {i["name"]: i.get("value", "") for i in form["inputs"]
                 if i.get("name") and i.get("type") not in ("checkbox", "submit")}
        for i in form["inputs"]:
            if i.get("type") == "password" and i.get("name"):
                dados[i["name"]] = senha
            elif i.get("type") == "checkbox" and i.get("name") and re.search(r"lembr|remember|manter", i["name"], re.I):
                dados[i["name"]] = i.get("value") or "1"  # "manter conectado", se houver
        acao = unescape(form["action"]) or url
        url, _ = self._get(acao, dados)
        if self._na_tela_de_login(url) or not self.logado():
            raise ErroLoginIntranet("a Intranet recusou o usuário/senha")
        self._gravar_cookies()

    def logado(self):
        url, _ = self._get("/portal/requisicao")
        return not self._na_tela_de_login(url)

    def garantir_login(self):
        if self.logado():
            return
        usuario, senha = ler_login()
        if not usuario:
            raise ErroLoginIntranet("login da Intranet não configurado")
        self.entrar(usuario, senha)

    def buscar_html(self, caminho):
        url, html = self._get(caminho)
        if self._na_tela_de_login(url):  # sessão caiu no meio: entra de novo e repete
            self.garantir_login()
            url, html = self._get(caminho)
        return html.decode("utf-8", "replace")

    def baixar(self, caminho):
        url, dados = self._get(caminho)
        if self._na_tela_de_login(url):
            self.garantir_login()
            url, dados = self._get(caminho)
        return dados


# ---------------------------------------------------------------- planilha do banco

def _celulas_xlsx(conteudo, aba, celulas):
    """{célula: valor} de uma aba de um .xlsx, lendo o XML direto (rápido,
    sem Excel nem openpyxl). Texto, número ou None."""
    with zipfile.ZipFile(BytesIO(conteudo)) as z:
        wb = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        m = re.search(r'<sheet\b[^>]*\bname="%s"[^>]*\br:id="([^"]+)"' % re.escape(aba), wb)
        if not m:
            raise ErroIntranet(f"a planilha do banco não tem a aba {aba}")
        alvo = re.search(r'<Relationship\b[^>]*\bId="%s"[^>]*\bTarget="([^"]+)"' % m.group(1), rels) or \
            re.search(r'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="%s"' % m.group(1), rels)
        xml_aba = z.read("xl/" + alvo.group(1).lstrip("/").replace("xl/", "", 1)).decode("utf-8")
        textos = []
        if "xl/sharedStrings.xml" in z.namelist():
            sst = z.read("xl/sharedStrings.xml").decode("utf-8")
            textos = [unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
                      for si in re.findall(r"<si>(.*?)</si>", sst, re.S)]
    saida = {}
    for cel in celulas:
        m = re.search(r'<c r="%s"(?=[\s>/])([^>]*?)(?:/>|>(.*?)</c>)' % cel, xml_aba, re.S)
        if not m or not m.group(2):
            saida[cel] = None
            continue
        attrs, corpo = m.group(1), m.group(2)
        v = re.search(r"<v>(.*?)</v>", corpo, re.S)
        if 't="s"' in attrs and v:
            saida[cel] = textos[int(v.group(1))].strip()
        elif 't="inlineStr"' in attrs:
            saida[cel] = unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", corpo, re.S))).strip()
        elif v:
            texto = unescape(v.group(1)).strip()
            try:
                n = float(texto)
                saida[cel] = int(n) if n.is_integer() else n
            except ValueError:
                saida[cel] = texto
        else:
            saida[cel] = None
    return saida


def dados_do_banco(proposta, log=print):
    """Acha a requisição da proposta na Intranet e lê a planilha do banco.
    Devolve {"tipo_operacao", "juncao", "nome_juncao", "protocolo",
    "arquivo", "requisicao"} ou levanta ErroIntranet (com o motivo)."""
    sessao = Sessao()
    sessao.garantir_login()
    html = sessao.buscar_html("/portal/requisicao/busca?" + urllib.parse.urlencode({"descricao": proposta}))
    candidatos = [(h, t) for h, t in _analisar(html).links
                  if "/requisicao/" in h and "/detalhar/" in h and re.match(r"\s*%s\b" % re.escape(str(proposta)), t)]
    if not candidatos:
        raise ErroIntranet(f"nenhuma requisição da proposta {proposta} na Intranet")
    # a do Banco D primeiro (o título termina com o banco)
    candidatos.sort(key=lambda ht: "banco d" not in ht[1].lower())
    for href, titulo in dict.fromkeys(candidatos):
        detalhe = sessao.buscar_html(unescape(href))
        anexos = [(h, t) for h, t in _analisar(detalhe).links
                  if "/requisicao/download/" in h and _chave(t).startswith(NOME_PLANILHA_BANCO)]
        if not anexos:
            continue
        link, nome = anexos[0]  # o primeiro é o que veio com o pedido
        conteudo = sessao.baixar(unescape(link))
        if conteudo[:2] != b"PK":
            raise ErroIntranet(f"o anexo {nome} não veio como Excel")
        c = _celulas_xlsx(conteudo, "FORMULARIO", ["F36", "F5", "F6", "F7"])
        log(f"[INTRANET] Requisição \"{titulo[:70]}\" -> {nome}: tipo de operação {c['F36']!r}")
        return {"tipo_operacao": c["F36"], "juncao": c["F5"], "nome_juncao": c["F6"], "protocolo": c["F7"],
                "arquivo": nome, "requisicao": titulo}
    raise ErroIntranet(f"a requisição da proposta {proposta} na Intranet não tem a planilha do banco "
                        "(Formulario_Banco_D...xlsx) nos arquivos")


def testar_login(usuario, senha):
    """Pra janela de configuração: entra e confirma (levanta ErroIntranet)."""
    s = Sessao()
    s.entrar(usuario, senha)
