r"""
Helpers comuns do Downloader de Fotos: caminhos, credenciais, sessão salva
e login automático na Central de Gestão.

Login da Central de Gestão (o robô fala direto com a API, ver central_gestao.py):
  1. token salvo (%LOCALAPPDATA%\DownloaderFotos\token_central_gestao.json) -
     enquanto valer, nem passa pelo login;
  2. e-mail/senha gravados pela janela do programa ("Configurar login…") no
     Gerenciador de Credenciais do Windows (criptografado por usuário);
  3. variáveis de ambiente CENTRAL_GESTAO_USUARIO / CENTRAL_GESTAO_SENHA, ou
     CENTRAL_USUARIO / CENTRAL_SENHA (mesmo padrão dos robôs de cadastro).
Nunca imprime a senha. As funções de navegador (Playwright) daqui só são
usadas pelo ferramentas/mapear_central_gestao.py.
"""

import os
import sys

APP_NOME = "DownloaderFotos"
VERSAO = "2.6"  # 15/10/2026: aviso de versão nova, erros/uso na pasta pública, trabalho salvo por proposta (2.5: observações das fotos)
CENTRAL_GESTAO_URL = "https://central-gestao.example.com/home"
CREDENCIAL_ALVO = "DownloaderFotos/central_gestao"

# Pasta do programa: ao lado do .exe (PyInstaller) ou deste .py
if getattr(sys, "frozen", False):
    PASTA_SCRIPT = os.path.dirname(sys.executable)
    PASTA_RECURSOS = getattr(sys, "_MEIPASS", PASTA_SCRIPT)  # arquivos embutidos no .exe
else:
    PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
    PASTA_RECURSOS = PASTA_SCRIPT

# Edição do pré-preenchimento do formulário do banco (perfis.PrePreenchimento):
# "completo" (padrão: aprovado pelo Rafael em 09/10/2026 depois do teste com 2
# .exe) ou "basico". ferramentas/gerar_exe_teste.ps1 ainda gera .exe marcados com um
# preenchimento.txt embutido; rodando pelo Python: variável DOWNLOADER_PREENCHIMENTO.
EDICOES = {"basico": "Básico", "completo": "Completo"}


def _edicao():
    """(edição, se veio marcada no .exe - aí a janela mostra qual é)."""
    try:
        with open(os.path.join(PASTA_RECURSOS, "preenchimento.txt"), encoding="utf-8") as f:
            valor, marcada = f.read().strip().lower(), True
    except OSError:
        valor, marcada = os.environ.get("DOWNLOADER_PREENCHIMENTO", "completo").strip().lower(), False
    return (valor, marcada) if valor in EDICOES else ("completo", marcada)


PREENCHIMENTO, EDICAO_MARCADA = _edicao()
# o que aparece ao lado da versão na janela ("" no .exe normal)
ROTULO_EDICAO = f"TESTE · preenchimento {EDICOES[PREENCHIMENTO]}" if EDICAO_MARCADA else ""

# pasta pública: onde fica o .exe oficial (aviso de versão nova) e pra onde vão
# os erros e o uso de cada PC (pasta_publica.py). A variável serve pros testes.
PASTA_PUBLICA = os.environ.get("DOWNLOADER_PASTA_PUBLICA") or r"\\servidor\pasta_publica"

PASTA_DADOS = os.path.join(os.environ.get("LOCALAPPDATA", PASTA_SCRIPT), APP_NOME)
os.makedirs(PASTA_DADOS, exist_ok=True)
SESSION_FILE = os.path.join(PASTA_DADOS, "sessao_central_gestao.json")
CONFIG_FILE = os.path.join(PASTA_DADOS, "config.json")
PASTA_LOGS = os.path.join(PASTA_DADOS, "logs")


# ---------------------------------------------------------------- credenciais
#
# Onde o login digitado na janela fica guardado: arquivo login.dat em
# %LOCALAPPDATA%\DownloaderFotos, criptografado com a DPAPI do Windows (só o
# mesmo usuário, no mesmo PC, consegue abrir - é o que o Chrome usa pras
# senhas). Antes era o Gerenciador de Credenciais, mas numa máquina da
# empresa ele não guardou o login (07/10/2026 - provável política de rede
# que bloqueia credenciais salvas); a DPAPI não é afetada por essa política.
# O Gerenciador de Credenciais continua sendo lido, por compatibilidade.

ARQ_LOGIN = os.path.join(PASTA_DADOS, "login.dat")
ultimo_erro_login = ""  # motivo da última falha ao ler o login (vai pro log)


def _ler_login_dpapi():
    global ultimo_erro_login
    if not os.path.exists(ARQ_LOGIN):
        return None, None
    try:
        import json
        import win32crypt

        with open(ARQ_LOGIN, "rb") as f:
            dados = json.loads(win32crypt.CryptUnprotectData(f.read(), None, None, None, 0)[1])
        return dados.get("usuario") or None, dados.get("senha") or None
    except Exception as e:
        ultimo_erro_login = f"login.dat ilegível ({type(e).__name__}: {e})"
        return None, None


def _ler_credencial_windows():
    global ultimo_erro_login
    try:
        import win32cred

        cred = win32cred.CredRead(CREDENCIAL_ALVO, win32cred.CRED_TYPE_GENERIC)
        blob = cred["CredentialBlob"] or b""
        senha = blob.decode("utf-16-le") if isinstance(blob, bytes) else blob
        return cred["UserName"] or None, senha or None
    except Exception as e:
        if getattr(e, "winerror", None) != 1168:  # 1168 = não existe: normal
            ultimo_erro_login = f"Gerenciador de Credenciais ({type(e).__name__}: {e})"
        return None, None


def gravar_login(usuario, senha):
    """Grava o login (DPAPI) e confere lendo de volta. Levanta RuntimeError
    com o motivo se não der pra confirmar."""
    import json
    import win32crypt

    blob = win32crypt.CryptProtectData(
        json.dumps({"usuario": usuario, "senha": senha}).encode("utf-8"),
        "DownloaderFotos", None, None, None, 0)
    tmp = ARQ_LOGIN + ".tmp"
    with open(tmp, "wb") as f:
        f.write(blob)
    os.replace(tmp, ARQ_LOGIN)
    if _ler_login_dpapi() != (usuario, senha):
        raise RuntimeError(f"o login foi gravado mas não pôde ser lido de volta. {ultimo_erro_login}")


def apagar_login():
    for caminho in (ARQ_LOGIN, SESSION_FILE):
        try:
            os.remove(caminho)
        except FileNotFoundError:
            pass
    try:
        import win32cred

        win32cred.CredDelete(CREDENCIAL_ALVO, win32cred.CRED_TYPE_GENERIC, 0)
    except Exception:
        pass


def apagar_sessao():
    try:
        os.remove(SESSION_FILE)
    except FileNotFoundError:
        pass


def fonte_login():
    """De onde vem o login em uso (pra log/diagnóstico), nunca a senha."""
    if all(_ler_login_dpapi()):
        return "salvo neste PC"
    if all(_ler_credencial_windows()):
        return "Gerenciador de Credenciais"
    for prefixo in ("CENTRAL_GESTAO", "CENTRAL"):
        if os.environ.get(f"{prefixo}_USUARIO", "").strip() and os.environ.get(f"{prefixo}_SENHA"):
            return f"variáveis {prefixo}_*"
    return f"nenhuma ({ultimo_erro_login or 'nada configurado'}; pasta {PASTA_DADOS})"


def credenciais_configuradas():
    """(usuario, senha) da primeira fonte disponível, ou (None, None).
    O login salvo pela janela do programa vem antes das variáveis de
    ambiente: é o que a pessoa acabou de digitar."""
    for ler in (_ler_login_dpapi, _ler_credencial_windows):
        usuario, senha = ler()
        if usuario and senha:
            return usuario, senha
    for prefixo in ("CENTRAL_GESTAO", "CENTRAL"):
        usuario = os.environ.get(f"{prefixo}_USUARIO", "").strip()
        senha = os.environ.get(f"{prefixo}_SENHA", "")
        if usuario and senha:
            return usuario, senha
    return None, None


def usuario_configurado():
    """Só o usuário (pra mostrar na tela), nunca a senha."""
    return credenciais_configuradas()[0]


# ---------------------------------------------------------------- navegador

def abrir_navegador(p, mostrar=False):
    """Usa o Edge (vem em todo Windows) ou o Chrome instalados - o .exe não
    precisa carregar um navegador próprio. Cai pro Chromium do Playwright
    só em máquina de desenvolvimento."""
    args = ["--start-maximized"] if mostrar else []
    for canal in ("msedge", "chrome", None):
        try:
            return p.chromium.launch(channel=canal, headless=not mostrar, args=args)
        except Exception:
            continue
    raise RuntimeError("Não achei o Microsoft Edge nem o Google Chrome instalados.")


def novo_contexto(browser, mostrar=False):
    return browser.new_context(
        no_viewport=mostrar,
        viewport=None if mostrar else {"width": 1600, "height": 900},
        accept_downloads=True,
        storage_state=SESSION_FILE if os.path.exists(SESSION_FILE) else None,
    )


# ---------------------------------------------------------------- login

def tela_de_login(page):
    try:
        return page.locator("input[type='password']:visible").count() > 0
    except Exception:
        return False


def login_automatico(page, usuario, senha, timeout=30000):
    """Preenche e envia o formulário de login a partir do campo de senha
    visível. Acha o usuário como o primeiro input de texto/e-mail do mesmo
    <form>. True só se conseguiu submeter."""
    try:
        campo_senha = page.locator("input[type='password']:visible").first
        campo_senha.wait_for(timeout=timeout)
    except Exception:
        return False

    form = campo_senha.locator("xpath=ancestor::form[1]")
    escopo = form if form.count() > 0 else page
    campo_usuario = escopo.locator(
        "input[type='text']:visible, input[type='email']:visible, input:not([type]):visible"
    ).first
    try:
        campo_usuario.fill(usuario, timeout=timeout)
        campo_senha.fill(senha, timeout=timeout)
    except Exception:
        return False

    botao = escopo.locator("button[type='submit'], input[type='submit']").first
    try:
        if botao.count() > 0:
            botao.click(timeout=timeout)
        else:
            campo_senha.press("Enter")
    except Exception:
        return False
    return True


def _esperar_sair_do_login(page, timeout):
    try:
        page.locator("input[type='password']:visible").first.wait_for(state="detached", timeout=timeout)
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:
        pass
    return not tela_de_login(page)


def garantir_login(page, permitir_manual=True, timeout=30000, log=print):
    """Abre a Central de Gestão e garante que está logado. True se logou; salva a
    sessão pra próxima vez."""
    page.goto(CENTRAL_GESTAO_URL, wait_until="domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    if not tela_de_login(page):
        return _salvar(page)

    usuario, senha = credenciais_configuradas()
    if usuario:
        log("[LOGIN] Fazendo login na Central de Gestão...")
        if login_automatico(page, usuario, senha, timeout) and _esperar_sair_do_login(page, timeout):
            return _salvar(page)
        log("[LOGIN] [AVISO] O login automático não passou (usuário/senha errados ou página lenta).")
    else:
        log("[LOGIN] Nenhum usuário/senha configurado.")

    if not permitir_manual:
        return False
    log("[LOGIN] Faça o login na janela do navegador (até 5 minutos)...")
    if _esperar_sair_do_login(page, 300000):
        return _salvar(page)
    return False


def _salvar(page):
    try:
        page.context.storage_state(path=SESSION_FILE)
    except Exception:
        pass
    return True
