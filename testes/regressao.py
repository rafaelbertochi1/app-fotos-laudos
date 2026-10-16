r"""
Teste completo antes de publicar (pedido de 15/10/2026): roda propostas fixas
de ponta a ponta - busca na Central de Gestão, tipo de operação na Intranet,
Excel pelo Excel, pelo modo sem janela do programa - e compara o resultado
com o gabarito guardado em testes\gabaritos\. Pega o que um teste de tela
não pegaria: modelo errado, campo do FORMULARIO que mudou, foto que sumiu do
Excel, impressão desconfigurada, .exe sem algum módulo (já aconteceu com o
win32timezone).

    python testes\regressao.py                       pelo código (Python)
    python testes\regressao.py --exe dist\DownloaderFotos_v2.6.exe   pelo .exe
    python testes\regressao.py --rapido              só 1 proposta por banco
    python testes\regressao.py --gravar              aceita o resultado como gabarito novo
                                                     (depois de conferir que a diferença é esperada)

O publicar.ps1 roda este teste com o .exe novo antes de pôr na pasta pública.
Precisa de Windows com Excel (o Excel é montado pelo próprio Excel).

Versão de portfólio: no projeto real eram propostas de verdade, buscadas na
Central de Gestão e na Intranet com o login salvo no PC. Aqui as propostas
são fictícias e quem responde é o servidor de exemplo (servidor_exemplo.py),
que o teste sobe numa porta livre; o programa roda com uma pasta de dados
(%LOCALAPPDATA%) temporária, sem login/token/config de verdade.

Gabarito por caso: modelo escolhido, nº de fotos em cada aba, configuração de
impressão de cada aba e o conteúdo das abas de fotos/formulário (fórmula
como texto: =TODAY() não muda todo dia).
"""

import argparse
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
import zipfile

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GABARITOS = os.path.join(RAIZ, "testes", "gabaritos")
sys.path.insert(0, RAIZ)
sys.path.insert(0, os.path.join(RAIZ, "testes"))

from perfis import PERFIS  # noqa: E402

# modelo None = o programa descobre sozinho (Banco D, pela Intranet) e o teste confere
CASOS = [
    {"banco": "Banco D", "proposta": "100001", "modelo": None, "esperado": "Aquisição simples"},
    {"banco": "Banco D", "proposta": "100002", "modelo": None, "esperado": "Aquisição à vista FGTS"},
    {"banco": "Banco D", "proposta": "100003", "modelo": None, "esperado": "Consórcio urbano"},
    {"banco": "Banco D", "proposta": "100004", "modelo": None, "esperado": "Consórcio reforma sem ampliação"},
    {"banco": "Banco D", "proposta": "100005", "modelo": None, "esperado": "Garantia imóvel urbano"},
    {"banco": "Banco C", "proposta": "200001", "modelo": "Comparativo", "esperado": "Comparativo"},
    {"banco": "Banco C", "proposta": "200001", "modelo": "Evolutivo", "esperado": "Evolutivo"},
]
RAPIDO = [0, 5]  # 1 Banco D (automático) + 1 Banco C
LIMITE_BYTES = 14_000_000
TEMPO_MAXIMO = 900  # s por proposta


def nome_caso(c):
    """Nome do gabarito, sem acento (arquivo no Git): Banco_D_100001_aquisicao_simples."""
    sem_acento = unicodedata.normalize("NFKD", c["esperado"]).encode("ascii", "ignore").decode()
    banco = re.sub(r"[^A-Za-z0-9]+", "_", c["banco"])
    return f"{banco}_{c['proposta']}_{re.sub(r'[^A-Za-z0-9]+', '_', sem_acento).strip('_').lower()}"


# ---------------------------------------------------------------- leitura do Excel gerado

def _abas(z):
    """[(nome, caminho do xml da aba)] na ordem do arquivo."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    alvos = {m.group(1): m.group(2) for m in re.finditer(r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels)}
    alvos.update({m.group(2): m.group(1) for m in re.finditer(r'<Relationship\b[^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
    abas = []
    for m in re.finditer(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb):
        alvo = alvos[m.group(2)].lstrip("/")
        abas.append((m.group(1).replace("&amp;", "&"), alvo if alvo.startswith("xl/") else "xl/" + alvo))
    return abas, wb


def _rels_da_aba(z, xml_aba):
    pasta, nome = os.path.split(xml_aba)
    caminho = f"{pasta}/_rels/{nome}.rels"
    return z.read(caminho).decode("utf-8") if caminho in z.namelist() else ""


def resumo_excel(caminho, abas_conteudo):
    """O que vai pro gabarito (ver o topo)."""
    saida = {"abas": {}, "celulas": {}}
    with zipfile.ZipFile(caminho) as z:
        abas, wb = _abas(z)
        areas = {}
        for m in re.finditer(r'<definedName\b[^>]*name="_xlnm\.Print_Area"[^>]*localSheetId="(\d+)"[^>]*>(.*?)</definedName>'
                             r'|<definedName\b[^>]*localSheetId="(\d+)"[^>]*name="_xlnm\.Print_Area"[^>]*>(.*?)</definedName>', wb):
            idx, texto = (m.group(1), m.group(2)) if m.group(1) else (m.group(3), m.group(4))
            areas[int(idx)] = re.sub(r"^.*!", "", texto)
        for n, (nome, xml_aba) in enumerate(abas):
            xml = z.read(xml_aba).decode("utf-8")
            imagens = 0
            for d in re.findall(r'Target="([^"]*drawings/drawing\d+\.xml)"', _rels_da_aba(z, xml_aba)):
                alvo = os.path.normpath(os.path.join(os.path.dirname(xml_aba), d)).replace("\\", "/")
                imagens += len(re.findall(r"<xdr:pic>", z.read(alvo).decode("utf-8")))
            setup = re.search(r"<pageSetup\b([^>]*)/?>", xml)
            attrs = dict(re.findall(r'(\w+)="([^"]*)"', setup.group(1))) if setup else {}
            ajustar = re.search(r'<pageSetUpPr\b[^>]*fitToPage="(\w+)"', xml)
            quebras = re.search(r'<rowBreaks\b[^>]*count="(\d+)"', xml)
            saida["abas"][nome] = {
                "imagens": imagens,
                "area_impressao": areas.get(n),
                "impressao": {k: attrs.get(k) for k in ("scale", "fitToWidth", "fitToHeight", "orientation", "paperSize")},
                "ajustar_pagina": ajustar.group(1) if ajustar else None,
                "quebras": int(quebras.group(1)) if quebras else 0,
            }
    import openpyxl  # só no teste (o .exe não leva)
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # validações estendidas dos modelos: o openpyxl reclama e ignora
        wbx = openpyxl.load_workbook(caminho, data_only=False)
    for aba in abas_conteudo:
        if aba not in wbx.sheetnames:
            continue
        celulas = {}
        for linha in wbx[aba].iter_rows():
            for c in linha:
                v = c.value
                if v is None or v == "":
                    continue
                if isinstance(v, (dt.date, dt.datetime, dt.time)):
                    v = v.isoformat()
                elif not isinstance(v, (int, float, str, bool)):
                    v = str(v)
                celulas[c.coordinate] = v
        saida["celulas"][aba] = celulas
    return saida


def abas_de_conteudo(perfil):
    abas = [perfil.fotos.nome] + ([perfil.locais.nome] if perfil.locais else [])
    if perfil.pre_preenchimento:
        abas.append(perfil.pre_preenchimento.aba)
    return abas


# ---------------------------------------------------------------- comparação

def diferencas(antes, agora, caminho=""):
    """Lista de textos 'onde: antes -> agora'."""
    if isinstance(antes, dict) and isinstance(agora, dict):
        saida = []
        for k in sorted(set(antes) | set(agora), key=str):
            if k not in antes:
                saida.append(f"{caminho}{k}: (não tinha) -> {agora[k]!r}")
            elif k not in agora:
                saida.append(f"{caminho}{k}: {antes[k]!r} -> (sumiu)")
            else:
                saida += diferencas(antes[k], agora[k], f"{caminho}{k} / ")
        return saida
    return [] if antes == agora else [f"{caminho.rstrip(' /')}: {antes!r} -> {agora!r}"]


# ---------------------------------------------------------------- execução

def ambiente(url_servidor, pasta):
    """Variáveis do programa testado: servidor de exemplo + pasta de dados própria
    (nada de login, token, config ou trabalho salvo deste PC)."""
    import servidor_exemplo

    env = dict(os.environ, **servidor_exemplo.variaveis(url_servidor))
    env["LOCALAPPDATA"] = os.path.join(pasta, "localappdata")
    env["DOWNLOADER_PASTA_PUBLICA"] = os.path.join(pasta, "pasta_publica")
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def rodar(caso, exe, pasta, env=None):
    """Gera o Excel do caso pelo modo sem janela. Devolve (resultado, falhas, avisos)."""
    log = os.path.join(pasta, nome_caso(caso) + ".log")
    destino = os.path.join(pasta, nome_caso(caso))
    cmd = ([exe] if exe else [sys.executable, os.path.join(RAIZ, "app.py")]) + [
        "--proposta", caso["proposta"], "--destino", destino, "--banco", caso["banco"], "--log", log]
    if caso["modelo"]:
        cmd += ["--modelo", caso["modelo"]]
    falhas, avisos = [], []
    inicio = time.time()
    try:
        proc = subprocess.run(cmd, cwd=RAIZ, timeout=TEMPO_MAXIMO, capture_output=True, text=True,
                              encoding="utf-8", errors="replace", env=env)
        codigo = proc.returncode
    except subprocess.TimeoutExpired:
        return None, [f"passou de {TEMPO_MAXIMO} s"], [], time.time() - inicio
    try:
        with open(log, encoding="utf-8") as f:
            texto = f.read()
    except OSError:
        texto = ""
        falhas.append(f"não gerou log (código {codigo}){': ' + proc.stderr[-500:] if proc.stderr else ''}")
    m = re.search(r"RESULTADO OK (.+?) falhas=(\d+)", texto)
    if not m:
        erro = texto.split("RESULTADO ERRO", 1)[-1].strip()[-1500:] if "RESULTADO ERRO" in texto else texto[-800:]
        falhas.append(f"não gerou o Excel (código {codigo}):\n{erro}")
        return None, falhas, avisos, time.time() - inicio
    excel = m.group(1)
    if int(m.group(2)):
        falhas.append(f"{m.group(2)} foto(s) não baixaram da Central de Gestão")
    escolhido = re.search(r"\[MODELO\] \S+: .*? -> (.+)", texto)
    modelo = caso["modelo"] or (escolhido.group(1).strip() if escolhido else None)
    if modelo != caso["esperado"]:
        falhas.append(f"modelo {modelo!r}, esperado {caso['esperado']!r}")
    tamanho = os.path.getsize(excel)
    if tamanho > LIMITE_BYTES:
        falhas.append(f"Excel com {tamanho / 1e6:.1f} MB (limite 14 MB)")
    for linha in texto.splitlines():
        if "[AVISO]" in linha and "Não consegui" in linha:
            falhas.append(linha.strip())
        elif "[AVISO]" in linha:
            avisos.append(linha.strip()[:200])
    resultado = {"modelo": modelo, **resumo_excel(excel, abas_de_conteudo(PERFIS[caso["banco"]]))}
    return resultado, falhas, avisos, time.time() - inicio


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--exe", help="testar este .exe (padrão: o código, pelo Python)")
    ap.add_argument("--rapido", action="store_true", help="só 1 proposta por banco")
    ap.add_argument("--gravar", action="store_true", help="aceitar o resultado como gabarito")
    ap.add_argument("--manter", action="store_true", help="não apagar os Excel gerados (pasta temporária)")
    a = ap.parse_args()
    exe = os.path.abspath(a.exe) if a.exe else None
    casos = [CASOS[i] for i in RAPIDO] if a.rapido else CASOS
    pasta = tempfile.mkdtemp(prefix="regressao_")
    os.makedirs(GABARITOS, exist_ok=True)
    import servidor_exemplo

    servidor, url = servidor_exemplo.iniciar()
    env = ambiente(url, pasta)
    print(f"Teste completo: {len(casos)} proposta(s) por {'o .exe ' + os.path.basename(exe) if exe else 'Python'}"
          f" | servidor de exemplo em {url} | arquivos em {pasta}", flush=True)
    ruins = 0
    for caso in casos:
        rotulo = f"{caso['banco']} {caso['proposta']} ({caso['esperado']})"
        print(f"\n- {rotulo} ...", flush=True)
        resultado, falhas, avisos, tempo = rodar(caso, exe, pasta, env)
        arq = os.path.join(GABARITOS, nome_caso(caso) + ".json")
        if resultado is not None:
            if a.gravar:
                with open(arq, "w", encoding="utf-8") as f:
                    json.dump(resultado, f, ensure_ascii=False, indent=1, sort_keys=True)
                print("  gabarito gravado")
            elif os.path.exists(arq):
                with open(arq, encoding="utf-8") as f:
                    difs = diferencas(json.load(f), json.loads(json.dumps(resultado)))
                if difs:
                    falhas.append(f"{len(difs)} diferença(s) com o gabarito:\n      "
                                  + "\n      ".join(difs[:40]) + ("\n      ..." if len(difs) > 40 else ""))
            else:
                falhas.append("sem gabarito ainda (rode com --gravar depois de conferir o Excel)")
        for av in avisos[:5]:
            print(f"  aviso: {av}")
        if falhas:
            ruins += 1
            print(f"  FALHOU ({tempo:.0f} s):")
            for f_ in falhas:
                print(f"    * {f_}")
        else:
            print(f"  OK ({tempo:.0f} s)")
    servidor.shutdown()
    print(f"\n{'TUDO OK' if not ruins else f'{ruins} DE {len(casos)} FALHARAM'}", flush=True)
    if not a.manter and not ruins:
        shutil.rmtree(pasta, ignore_errors=True)
    elif ruins:
        print(f"Excel e logs pra conferir: {pasta}")
    return 1 if ruins else 0


if __name__ == "__main__":
    raise SystemExit(main())
