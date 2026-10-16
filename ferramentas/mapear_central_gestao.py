"""
Ferramenta de mapeamento da Central de Gestão.

Abre o Chrome visível, já logado (sessão salva ou login automático com
CENTRAL_GESTAO_USUARIO/CENTRAL_GESTAO_SENHA, senão CENTRAL_USUARIO/CENTRAL_SENHA). Você
navega manualmente e, a cada tela importante, volta aqui no terminal e
aperta ENTER. O script salva em exploracao/<numero>_<nome>/:

  - screenshot.png      (foto da tela inteira)
  - frame_N.html        (HTML de cada frame da página)
  - resumo.json         (botões, links, campos e imagens de cada frame)
  - rede.json           (chamadas de API/imagens/downloads feitas desde a
                         captura anterior - é daqui que sai o jeito de
                         baixar todas as fotos de uma vez)

Roteiro sugerido:
  1. ENTER na tela inicial (home)               -> nome: home
  2. filtre por uma proposta, ENTER              -> nome: filtro_proposta
  3. abra a proposta / aba de fotos, ENTER       -> nome: tela_fotos
  4. clique no download de UMA foto, ENTER       -> nome: download_uma_foto
  5. digite 'sair'

A pasta exploracao/ fica aqui no projeto (fora do Git: tem dados de
cliente) e é lida depois, na hora de ajustar central_gestao.py.

Uso (da pasta do projeto): python ferramentas\\mapear_central_gestao.py
Precisa do Playwright (pip install playwright; não vai no .exe).
"""

import json
import os
import re
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # comum.py na pasta acima

from playwright.sync_api import sync_playwright  # noqa: E402

from comum import (CENTRAL_GESTAO_URL, PASTA_SCRIPT, SESSION_FILE, abrir_navegador,  # noqa: E402
                   garantir_login, novo_contexto)

EXPLORACAO_DIR = os.path.join(PASTA_SCRIPT, "exploracao")
TIPOS_REDE = {"xhr", "fetch", "image", "document"}
MAX_CORPO = 200_000  # bytes de cada resposta JSON guardados
PADRAO_SENSIVEL = re.compile(r"/auth|password|senha|\"token\"|accessToken|access_token", re.I)

JS_RESUMO_FRAME = """
() => {
    const pegar = (sel) => Array.from(document.querySelectorAll(sel)).map(el => ({
        texto: (el.innerText || el.value || "").trim().slice(0, 120),
        tag: el.tagName.toLowerCase(),
        id: el.id || null,
        name: el.name || null,
        classe: (typeof el.className === "string" ? el.className : null) || null,
        tipo: el.type || null,
        placeholder: el.placeholder || null,
        href: el.href || null,
        title: el.title || null,
        aria: el.getAttribute("aria-label"),
        download: el.getAttribute("download"),
        onclick: el.getAttribute("onclick"),
    }));
    const imagens = Array.from(document.querySelectorAll("img")).map(el => ({
        src: el.currentSrc || el.src,
        alt: el.alt || null,
        classe: (typeof el.className === "string" ? el.className : null) || null,
        largura: el.naturalWidth, altura: el.naturalHeight,
    }));
    return {
        botoes: pegar("button, [role=button], input[type=submit], input[type=button]"),
        links: pegar("a"),
        campos: pegar("input, select, textarea"),
        imagens,
    };
}
"""


class GravadorRede:
    """Guarda as requisições relevantes até a próxima captura."""

    def __init__(self):
        self.itens = []

    def ao_responder(self, resposta):
        req = resposta.request
        if req.resource_type not in TIPOS_REDE:
            return
        item = {
            "tipo": req.resource_type,
            "metodo": req.method,
            "url": req.url,
            "status": resposta.status,
            "content_type": resposta.headers.get("content-type", ""),
            "content_disposition": resposta.headers.get("content-disposition"),
        }
        # nunca grava login/token: o POST de /auth leva a senha em texto puro
        sensivel = bool(PADRAO_SENSIVEL.search(req.url))
        if req.method != "GET" and req.post_data:
            sensivel = sensivel or bool(PADRAO_SENSIVEL.search(req.post_data))
            item["post_data"] = "[REMOVIDO]" if sensivel else req.post_data[:5000]
        if "json" in item["content_type"] and not sensivel:
            try:
                corpo = resposta.text()[:MAX_CORPO]
                item["corpo"] = "[REMOVIDO]" if PADRAO_SENSIVEL.search(corpo) else corpo
            except Exception:
                pass
        self.itens.append(item)

    def ao_baixar(self, download):
        self.itens.append({
            "tipo": "DOWNLOAD",
            "url": download.url,
            "nome_sugerido": download.suggested_filename,
        })

    def esvaziar(self):
        itens, self.itens = self.itens, []
        return itens


def capturar_etapa(page, numero, nome, gravador):
    pasta = os.path.join(EXPLORACAO_DIR, f"{numero:02d}_{nome}")
    os.makedirs(pasta, exist_ok=True)

    try:
        page.screenshot(path=os.path.join(pasta, "screenshot.png"), full_page=True)
    except Exception as e:
        print(f"      [AVISO] screenshot falhou: {e}")

    resumo_frames = []
    for i, frame in enumerate(page.frames):
        try:
            html = frame.content()
        except Exception:
            html = ""
        with open(os.path.join(pasta, f"frame_{i}.html"), "w", encoding="utf-8") as f:
            f.write(html)
        try:
            dados = frame.evaluate(JS_RESUMO_FRAME)
        except Exception:
            dados = {"botoes": [], "links": [], "campos": [], "imagens": []}
        resumo_frames.append({"indice": i, "url": frame.url, **dados})

    resumo = {
        "etapa": nome,
        "url_pagina": page.url,
        "titulo_pagina": page.title(),
        "capturado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "frames": resumo_frames,
    }
    with open(os.path.join(pasta, "resumo.json"), "w", encoding="utf-8") as f:
        json.dump(resumo, f, ensure_ascii=False, indent=2)
    with open(os.path.join(pasta, "rede.json"), "w", encoding="utf-8") as f:
        json.dump(gravador.esvaziar(), f, ensure_ascii=False, indent=2)

    print(f"      Salvo em: {pasta}")


def main():
    os.makedirs(EXPLORACAO_DIR, exist_ok=True)
    numero = 1 + sum(1 for n in os.listdir(EXPLORACAO_DIR) if n[:2].isdigit())

    with sync_playwright() as p:
        browser = abrir_navegador(p, mostrar=True)
        context = novo_contexto(browser, mostrar=True)
        gravador = GravadorRede()
        context.on("response", gravador.ao_responder)
        context.on("page", lambda pg: pg.on("download", gravador.ao_baixar))
        page = context.new_page()

        print("=" * 60)
        print(" MAPEADOR - CENTRAL DE GESTÃO")
        print("=" * 60)
        if garantir_login(page):
            print("[LOGIN] OK - sessão salva.")
        else:
            print(f"[LOGIN] Não consegui logar. Abra {CENTRAL_GESTAO_URL} e logue na janela.")

        print("Navegue na janela do Chrome até a tela que quer mapear e")
        print("volte aqui pra apertar ENTER (roteiro no topo do arquivo).")

        while True:
            entrada = input(
                "\nENTER = capturar esta tela (ou digite um nome curto antes, "
                "ex.: 'tela_fotos'). 'sair' encerra: "
            ).strip()
            if entrada.lower() in ("sair", "fim", "exit", "encerrar"):
                break
            nome = entrada.lower().replace(" ", "_") if entrada else f"etapa_{numero}"
            # usa a aba ativa (a Central de Gestão pode abrir a proposta em outra aba)
            atual = context.pages[-1] if context.pages else page
            capturar_etapa(atual, numero, nome, gravador)
            numero += 1

        context.storage_state(path=SESSION_FILE)
        browser.close()

    print(f"\nPronto! Mapeamento salvo em {EXPLORACAO_DIR}.")


if __name__ == "__main__":
    main()
