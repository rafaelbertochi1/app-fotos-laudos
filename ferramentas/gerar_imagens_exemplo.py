r"""
Gera as imagens FICTÍCIAS de assets\ desta versão de portfólio (só Pillow):

  - icone.ico                  ícone do programa (câmera num quadrado arredondado)
  - logo_empresa.png           logo genérica, posta nas células "Logo" dos modelos
  - assinatura_responsavel.png "ASSINATURA EXEMPLO" com fundo transparente,
                               posta no quadro da assinatura (perfis.Assinatura)

Na versão real, a logo era a da empresa e a assinatura, a da engenheira
responsável - nenhuma das duas está no repositório.

    python ferramentas\gerar_imagens_exemplo.py
"""

import math
import os

from PIL import Image, ImageDraw, ImageFont

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(REPO, "assets")
AZUL = (30, 42, 56)        # = visual.CABECALHO
AZUL_CLARO = (0, 95, 184)  # = visual.AZUL
BRANCO = (255, 255, 255)


def _fonte(tamanho, negrito=False):
    for nome in (("segoeuib.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf") if negrito
                 else ("segoeui.ttf", "arial.ttf", "DejaVuSans.ttf")):
        try:
            return ImageFont.truetype(nome, tamanho)
        except OSError:
            continue
    return ImageFont.load_default(size=tamanho)


def icone():
    lado = 256
    im = Image.new("RGBA", (lado, lado), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((8, 8, lado - 8, lado - 8), radius=48, fill=AZUL_CLARO)
    d.rounded_rectangle((44, 82, 212, 196), radius=18, fill=BRANCO)    # corpo da câmera
    d.rounded_rectangle((96, 58, 160, 92), radius=10, fill=BRANCO)     # visor
    d.ellipse((88, 98, 168, 178), fill=AZUL_CLARO)                     # lente
    d.ellipse((106, 116, 150, 160), fill=BRANCO)
    d.rectangle((60, 196, 196, 204), fill=BRANCO)                      # "folha" embaixo (planilha)
    caminho = os.path.join(ASSETS, "icone.ico")
    im.save(caminho, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    return caminho


def logo():
    im = Image.new("RGBA", (330, 156), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    # casinha estilizada + nome genérico
    d.polygon([(20, 78), (78, 26), (136, 78)], fill=AZUL_CLARO)
    d.rectangle((36, 76, 120, 130), fill=AZUL_CLARO)
    d.rectangle((66, 96, 90, 130), fill=BRANCO)
    d.text((150, 44), "EMPRESA", font=_fonte(34, True), fill=AZUL)
    d.text((152, 86), "logo de exemplo", font=_fonte(20), fill=AZUL_CLARO)
    caminho = os.path.join(ASSETS, "logo_empresa.png")
    im.save(caminho)
    return caminho


def assinatura():
    im = Image.new("RGBA", (600, 200), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    # rabisco (senoide amortecida) + texto: claramente um exemplo
    pontos = [(40 + x, 90 - 38 * math.sin(x / 26) * math.exp(-x / 520)) for x in range(0, 520, 4)]
    d.line(pontos, fill=(20, 40, 120, 255), width=5, joint="curve")
    d.line((40, 130, 560, 130), fill=(60, 60, 60, 255), width=2)
    d.text((300, 162), "ASSINATURA EXEMPLO", font=_fonte(30, True), fill=(60, 60, 60, 255), anchor="mm")
    caminho = os.path.join(ASSETS, "assinatura_responsavel.png")
    im.save(caminho)
    return caminho


if __name__ == "__main__":
    os.makedirs(ASSETS, exist_ok=True)
    for gerar in (icone, logo, assinatura):
        print("gerado:", os.path.relpath(gerar(), REPO))
