r"""
Trabalho salvo de cada proposta (pedido de 15/10/2026): se o laudo voltar pra
correção dias depois, a pessoa busca a proposta de novo e o programa oferece
voltar do jeito que ficou - quais fotos entram, a ordem, os nomes trocados, a
seção de cada uma e as fotos do computador.

Fica neste PC, em %LOCALAPPDATA%\DownloaderFotos\trabalhos:
  proposta_<nº>.json    estado da galeria (Galeria.estado) + banco e data
  proposta_<nº>_pc\     cópia das fotos do computador (a original pode
                             ter sido apagada/movida até lá)
As fotos da Central de Gestão são identificadas pelo endereço delas (item "chave"),
não pela posição: se a vistoria ganhar/perder fotos, o resto continua certo.
Trabalho sem mexer há mais de DIAS_GUARDADO dias é apagado ao abrir.
"""

import json
import os
import re
import shutil
import time
from datetime import datetime

from comum import PASTA_DADOS

PASTA = os.path.join(PASTA_DADOS, "trabalhos")
DIAS_GUARDADO = 120


def _base(proposta):
    return os.path.join(PASTA, "proposta_" + re.sub(r"[^0-9]", "", str(proposta)))


def pasta_fotos_pc(proposta):
    return _base(proposta) + "_pc"


def ler(proposta):
    """Estado salvo ({} se não houver)."""
    try:
        with open(_base(proposta) + ".json", encoding="utf-8") as f:
            estado = json.load(f)
        return estado if isinstance(estado, dict) and estado.get("ordem") else {}
    except (OSError, ValueError):
        return {}


def salvar(proposta, estado, **extra):
    """Grava o estado (com a data). `extra` = campos a mais (ex.: gerado_em)."""
    os.makedirs(PASTA, exist_ok=True)
    anterior = ler(proposta)
    dados = dict(estado, salvo_em=datetime.now().isoformat(timespec="seconds"))
    for campo in ("gerado_em", "excel"):  # não se perde ao salvar de novo sem gerar
        if campo in anterior:
            dados.setdefault(campo, anterior[campo])
    dados.update(extra)
    caminho = _base(proposta) + ".json"
    with open(caminho + ".tmp", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)
    os.replace(caminho + ".tmp", caminho)
    _limpar_fotos_pc(proposta, dados)


def guardar_foto_pc(proposta, arquivo):
    """Copia a foto do computador pra pasta do trabalho; devolve a cópia (o
    nome dela é a chave da foto)."""
    pasta = pasta_fotos_pc(proposta)
    os.makedirs(pasta, exist_ok=True)
    nome, ext = os.path.splitext(os.path.basename(arquivo))
    destino, n = os.path.join(pasta, nome + ext), 2
    while os.path.exists(destino):  # mesmo nome de outra pasta: não sobrescreve
        destino, n = os.path.join(pasta, f"{nome} ({n}){ext}"), n + 1
    shutil.copyfile(arquivo, destino)
    return destino


def fotos_pc_salvas(proposta, estado):
    """Caminhos das fotos do computador que o estado usa, na ordem dele."""
    pasta = pasta_fotos_pc(proposta)
    caminhos = []
    for chave in estado.get("ordem", []):
        if chave.startswith("pc:"):
            caminho = os.path.join(pasta, chave[3:])
            if os.path.isfile(caminho):
                caminhos.append(caminho)
    return caminhos


def _limpar_fotos_pc(proposta, estado):
    """Apaga as cópias de fotos do computador que o estado não usa mais."""
    pasta = pasta_fotos_pc(proposta)
    usadas = {c[3:] for c in estado.get("ordem", []) if c.startswith("pc:")}
    if not os.path.isdir(pasta):
        return
    for nome in os.listdir(pasta):
        if nome not in usadas:
            try:
                os.remove(os.path.join(pasta, nome))
            except OSError:
                pass


def limpar_antigos(dias=DIAS_GUARDADO):
    """Apaga trabalhos parados há mais de `dias`."""
    if not os.path.isdir(PASTA):
        return
    limite = time.time() - dias * 86400
    for nome in os.listdir(PASTA):
        caminho = os.path.join(PASTA, nome)
        try:
            if nome.endswith(".json") and os.path.getmtime(caminho) < limite:
                os.remove(caminho)
                shutil.rmtree(caminho[:-5] + "_pc", ignore_errors=True)
        except OSError:
            pass


def resumo(estado):
    """Texto curto pra perguntar se retoma: quando e o que tinha."""
    def data(campo):
        try:
            return datetime.fromisoformat(estado[campo]).strftime("%d/%m/%Y às %H:%M")
        except (KeyError, ValueError, TypeError):
            return None
    fora = set(estado.get("fora", []))
    dentro = [c for c in estado.get("ordem", []) if c not in fora]
    partes = [f"{len(dentro)} foto(s) no Excel"]
    if estado.get("nomes"):
        partes.append(f"{len(estado['nomes'])} renomeada(s)")
    if estado.get("destinos"):
        partes.append(f"{len(estado['destinos'])} em quadros/seções")
    pc = sum(1 for c in estado.get("ordem", []) if c.startswith("pc:"))
    if pc:
        partes.append(f"{pc} do computador")
    quando = (f"Excel gerado em {data('gerado_em')}" if data("gerado_em") else
              f"mexida pela última vez em {data('salvo_em') or '?'}")
    return quando, ", ".join(partes)
