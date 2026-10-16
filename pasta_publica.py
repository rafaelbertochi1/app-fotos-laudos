r"""
O que o programa conversa com a pasta pública (\\servidor\pasta_publica), pedido
de 15/10/2026:

  1. versão nova: o .exe oficial fica lá como DownloaderFotos_v<versão>.exe;
     se for mais novo que este, a janela avisa (versao_mais_nova);
  2. erros: erro inesperado vira um arquivo curto em _erros\ (versão, quem,
     proposta, o erro e o fim do log) - o responsável vê sem ninguém mandar log;
  3. uso: cada Excel gerado vira uma linha em _uso\<usuário>_<PC>_<aaaa-mm>.csv
     (um arquivo por pessoa/PC/mês: dois PCs nunca escrevem no mesmo arquivo).
     ferramentas\resumo_uso.py resume.

Nada aqui pode travar ou derrubar o programa: tudo roda em segundo plano e
erro de rede é engolido. Se a pasta pública estiver fora do ar, os erros e o
uso esperam em %LOCALAPPDATA%\DownloaderFotos\a_enviar e vão na próxima vez.

Só o .exe manda (ATIVO); rodando pelo Python (desenvolvimento) ou no modo sem
janela (testes), não - senão os testes encheriam a pasta.
"""

import csv
import io
import os
import re
import socket
import sys
import threading
import traceback
from datetime import datetime

from comum import PASTA_DADOS, PASTA_PUBLICA, VERSAO

ATIVO = bool(getattr(sys, "frozen", False)) or os.environ.get("DOWNLOADER_RELATORIO") == "1"
PASTA_ERROS = os.path.join(PASTA_PUBLICA, "_erros")
PASTA_USO = os.path.join(PASTA_PUBLICA, "_uso")
A_ENVIAR = os.path.join(PASTA_DADOS, "a_enviar")
LINHAS_DO_LOG = 80  # fim do log que vai junto do erro

# o que a janela está fazendo (vai no arquivo de erro): proposta, banco, modelo, etapa, arquivo_log
contexto = {}
_trava = threading.Lock()


def usuario():
    return os.environ.get("USERNAME") or "desconhecido"


def computador():
    return os.environ.get("COMPUTERNAME") or socket.gethostname() or "PC"


def _seguro(texto):
    """Pedaço de nome de arquivo."""
    return re.sub(r"[^\w.-]+", "_", str(texto or "")).strip("_")[:40] or "x"


# ---------------------------------------------------------------- versão nova

def _tupla(versao):
    return tuple(int(p) for p in re.findall(r"\d+", versao))


def versao_mais_nova(pasta=PASTA_PUBLICA, atual=VERSAO):
    """(versão, caminho do .exe) se a pasta pública tiver um .exe mais novo
    que `atual`; senão None (inclusive com a pasta fora do ar)."""
    try:
        nomes = os.listdir(pasta)
    except OSError:
        return None
    melhor = None
    for nome in nomes:
        m = re.fullmatch(r"DownloaderFotos_v(\d+(?:\.\d+)*)\.exe", nome, re.I)
        if m and _tupla(m.group(1)) > _tupla(atual) and (not melhor or _tupla(m.group(1)) > _tupla(melhor[0])):
            melhor = (m.group(1), os.path.join(pasta, nome))
    return melhor


# ---------------------------------------------------------------- erros e uso

def _fim_do_log(caminho):
    try:
        with open(caminho, encoding="utf-8", errors="replace") as f:
            return "".join(f.readlines()[-LINHAS_DO_LOG:])
    except (OSError, TypeError):
        return "(sem log)\n"


def erro(titulo, detalhe=None):
    """Registra um erro inesperado (detalhe = traceback; padrão: o da exceção
    que está sendo tratada) e manda pra pasta pública em segundo plano."""
    if not ATIVO:
        return
    try:
        agora = datetime.now()
        ctx = dict(contexto)
        detalhe = detalhe or traceback.format_exc()
        texto = (f"Downloader de Fotos v{VERSAO} - erro em {agora:%d/%m/%Y %H:%M:%S}\n"
                 f"Usuário: {usuario()}   PC: {computador()}\n"
                 f"Programa: {sys.executable if getattr(sys, 'frozen', False) else 'python'}\n"
                 f"Banco: {ctx.get('banco', '-')}   Modelo: {ctx.get('modelo', '-')}   "
                 f"Proposta: {ctx.get('proposta', '-')}   Etapa: {ctx.get('etapa', '-')}\n\n"
                 f"ERRO: {titulo}\n\n{detalhe.strip()}\n\n"
                 f"---- fim do log ({ctx.get('arquivo_log') or '-'}) ----\n{_fim_do_log(ctx.get('arquivo_log'))}")
        nome = f"{agora:%Y-%m-%d_%H%M%S}_{_seguro(usuario())}_{_seguro(ctx.get('proposta') or 'sem-proposta')}.txt"
        pasta = os.path.join(A_ENVIAR, "erros")
        os.makedirs(pasta, exist_ok=True)
        with open(os.path.join(pasta, nome), "w", encoding="utf-8") as f:
            f.write(texto)
    except Exception:
        return
    enviar_em_segundo_plano()


CAMPOS_USO = ["data", "hora", "usuario", "pc", "versao", "banco", "modelo", "proposta", "fotos_central",
              "fotos_no_excel", "fotos_do_pc", "observacoes", "retomado", "segundos"]


def uso(**campos):
    """Uma linha no CSV de uso (cada Excel gerado). Campos: CAMPOS_USO."""
    if not ATIVO:
        return
    try:
        agora = datetime.now()
        linha = dict(campos, data=f"{agora:%d/%m/%Y}", hora=f"{agora:%H:%M:%S}", usuario=usuario(),
                     pc=computador(), versao=VERSAO)
        os.makedirs(A_ENVIAR, exist_ok=True)
        with _trava, open(os.path.join(A_ENVIAR, f"uso_{agora:%Y-%m}.csv"), "a", encoding="utf-8", newline="") as f:
            csv.DictWriter(f, CAMPOS_USO, delimiter=";", extrasaction="ignore").writerow(linha)
    except Exception:
        return
    enviar_em_segundo_plano()


def enviar():
    """Manda o que estiver esperando em a_enviar. Devolve quantos arquivos
    foram (0 se a pasta pública estiver fora do ar)."""
    with _trava:
        if not os.path.isdir(A_ENVIAR) or not os.path.isdir(PASTA_PUBLICA):
            return 0
        enviados = 0
        pasta_erros = os.path.join(A_ENVIAR, "erros")
        for nome in (os.listdir(pasta_erros) if os.path.isdir(pasta_erros) else []):
            try:
                os.makedirs(PASTA_ERROS, exist_ok=True)
                origem = os.path.join(pasta_erros, nome)
                with open(origem, "rb") as f:
                    conteudo = f.read()
                with open(os.path.join(PASTA_ERROS, nome), "wb") as f:
                    f.write(conteudo)
                os.remove(origem)
                enviados += 1
            except OSError:
                pass
        for nome in os.listdir(A_ENVIAR):
            m = re.fullmatch(r"uso_(\d{4}-\d{2})\.csv", nome)
            if not m:
                continue
            origem = os.path.join(A_ENVIAR, nome)
            destino = os.path.join(PASTA_USO, f"{_seguro(usuario())}_{_seguro(computador())}_{m.group(1)}.csv")
            try:
                os.makedirs(PASTA_USO, exist_ok=True)
                with open(origem, encoding="utf-8") as f:
                    linhas = f.read()
                novo = not os.path.exists(destino)
                # utf-8-sig no arquivo novo: o Excel abre os acentos certo
                with open(destino, "a", encoding="utf-8-sig" if novo else "utf-8", newline="") as f:
                    if novo:
                        cab = io.StringIO()
                        csv.writer(cab, delimiter=";").writerow(CAMPOS_USO)
                        f.write(cab.getvalue())
                    f.write(linhas)
                os.remove(origem)
                enviados += 1
            except OSError:
                pass
        return enviados


def enviar_em_segundo_plano():
    threading.Thread(target=lambda: _engolir(enviar), daemon=True).start()


def _engolir(funcao):
    try:
        funcao()
    except Exception:
        pass
