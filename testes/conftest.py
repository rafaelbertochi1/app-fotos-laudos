"""
Configuração do pytest: roda antes de qualquer teste importar o programa.

comum.py cria a pasta de dados em %LOCALAPPDATA%\\DownloaderFotos assim que é
importado (login, token, config, trabalho salvo). Nos testes ela vai pra uma
pasta temporária - nunca mexe na do PC de quem roda nem depende dela (no
Linux do CI a variável nem existe). A pasta pública também.
"""

import os
import sys
import tempfile

_PASTA = tempfile.mkdtemp(prefix="testes_fotos_")
os.environ["LOCALAPPDATA"] = os.path.join(_PASTA, "localappdata")
os.environ["DOWNLOADER_PASTA_PUBLICA"] = os.path.join(_PASTA, "pasta_publica")
for _variavel in ("CENTRAL_GESTAO_USUARIO", "CENTRAL_GESTAO_SENHA", "CENTRAL_USUARIO", "CENTRAL_SENHA",
                  "INTRANET_USUARIO", "INTRANET_SENHA", "CENTRAL_GESTAO_API", "INTRANET_URL"):
    os.environ.pop(_variavel, None)

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
sys.path.insert(0, os.path.join(RAIZ, "testes"))
