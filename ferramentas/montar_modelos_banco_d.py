"""
Monta os modelos do Banco D (um por tipo de operação) a partir do modelo
base limpo, aplicando só o que muda em cada um. Pedido do Rafael em
13/10/2026, com 4 planilhas de referência (casos/modelos que a equipe já
usa, em referencia/banco_d/modelos_novos/):

  Aquisição simples            -> base (tipo de operação AQUISIÇÃO SIMPLES)
  Aquisição à vista FGTS       -> só o tipo de operação
  Consórcio urbano             -> só o tipo de operação
  Consórcio reforma s/ ampl.   -> tipo + bloco "serviços de reforma / custo /
                                  prazo" (FORMULARIO 170-172 e PAG.1 61-64,
                                  formatação copiada do caso 100007)
  Garantia imóvel urbano       -> tipo + PAG.1 com "Valor de Liquidez"
                                  (como no caso 100005)

Em todos: corrige a PAG.2!A47, que no modelo original do Banco D é 0 em vez
de =FORMULARIO!E126 (9ª linha de cômodos sumia; os casos prontos já tinham
corrigido à mão), e põe as fórmulas que a equipe usa nos casos prontos:
área total (F47) = construída + comum e valor de operação (J139) = valor
final (J133). Garantia: também o aviso "VIDE CARTA..." no topo da PAG.1
quando há item restritivo marcado.

O Excel de automação estraga a impressão ao salvar (ver impressao.py), então
no fim a configuração de impressão é restaurada a partir do modelo base.

Uso (com o Excel fechado):  python ferramentas\\montar_modelos_banco_d.py

Versão de portfólio: o modelo base e as planilhas de referência (referencia/)
têm dados de clientes e não estão no repositório, então este script fica como
registro de como os modelos reais foram montados. Os modelos de
modelos/banco_d/ daqui são fictícios, gerados por
ferramentas/gerar_modelos_exemplo.py com o mesmo resultado final (inclusive a
PAG.2!A47 corrigida e as variações de reforma e garantia).
"""

import os
import shutil
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(AQUI)
sys.path.insert(0, REPO)
import excel_fotos  # noqa: E402
import impressao  # noqa: E402

PASTA = os.path.join(REPO, "modelos", "banco_d")
BASE = os.path.join(REPO, "referencia", "banco_d", "modelo_banco_d_base_limpo.xlsx")
CASO_REFORMA = os.path.join(REPO, "referencia", "banco_d", "modelos_novos",
                            "consorcio_reforma_sem_ampliacao__100007.xlsx")

MODELOS = {
    "modelo_aquisição_simples.xlsx": {"tipo": "AQUISIÇÃO SIMPLES"},
    "modelo_aquisição_à_vista_FGTS.xlsx": {"tipo": "AQUISIÇÃO A VISTA COM FGTS"},
    "modelo_consórcio_urbano.xlsx": {"tipo": "CONSÓRCIO (IMÓVEL URBANO)"},
    "modelo_consórcio_reforma_sem_ampliação.xlsx": {"tipo": "CONSÓRCIO - REFORMA S/AMPLIAÇÃO (IMÓVEL URBANO)",
                                                     "reforma": True},
    "modelo_garantia_imóvel_urbano.xlsx": {"tipo": "GARANTIA (IMÓVEL URBANO)", "liquidez": True},
}
XL_PASTE_FORMATS = -4122
# a PAG.1 do modelo base (1084,5 pt de altura até a linha 68) imprime em
# escala fixa de 66% numa folha; o bloco de reforma deixa a página com 1199 pt
# -> 59% ocupa na folha o mesmo espaço (o caso 100007 usava "ajustar a 1
# página", ~62%, mas isso esmagaria as páginas de FOTOS COMPLEMENTARES)
ESCALA_PAG1_REFORMA = 59


def _escala(arquivo, aba, escala):
    """Troca a escala de impressão (<pageSetup scale>) de uma aba, no XML."""
    import re
    import zipfile

    with zipfile.ZipFile(arquivo) as z:
        itens = [(i, z.read(i.filename)) for i in z.infolist()]
    conteudo = {i.filename: d for i, d in itens}
    xml_aba = dict(impressao._abas(conteudo))[aba]
    x = conteudo[xml_aba].decode("utf-8")
    novo = re.sub(r'(<pageSetup\b[^>]*?\bscale=")\d+(")', r"\g<1>%d\2" % escala, x, count=1)
    if novo == x:
        raise RuntimeError(f"{aba}: pageSetup sem scale")
    conteudo[xml_aba] = novo.encode("utf-8")
    tmp = arquivo + ".tmp"
    with zipfile.ZipFile(tmp, "w") as z:
        for info, _ in itens:
            z.writestr(info, conteudo[info.filename], compress_type=info.compress_type)
    os.replace(tmp, arquivo)


def _bloco_reforma(wb, caso):
    f, p1 = wb.Worksheets("FORMULARIO"), wb.Worksheets("PAG.1")
    cf, cp = caso.Worksheets("FORMULARIO"), caso.Worksheets("PAG.1")
    # FORMULARIO 170-172: rótulos + campos (formatação do caso; valores do caso NÃO)
    cf.Range("C170:P172").Copy()
    f.Range("C170").PasteSpecial(XL_PASTE_FORMATS)
    for r in (170, 171, 172):
        f.Rows(r).RowHeight = cf.Rows(r).RowHeight
        f.Range(f"F{r}:L{r}").MergeCells = True
    f.Range("E170").Value = "Descrição dos serviços de reforma:"
    f.Range("E171").Value = "Custo da reforma:"
    f.Range("E172").Value = "Prazo para conclusão:"
    f.Range("F170:L172").ClearContents()
    # PAG.1: comentário vai até a 60 e as linhas 61-64 mostram a reforma
    p1.Range("A52:L64").UnMerge()
    cp.Range("A52:L64").Copy()
    p1.Range("A52").PasteSpecial(XL_PASTE_FORMATS)
    for r in range(52, 65):
        p1.Rows(r).RowHeight = cp.Rows(r).RowHeight
    for faixa in ("A52:K60", "A61:L61", "A62:L62", "A63:B63", "A64:B64", "C63:D63"):
        p1.Range(faixa).MergeCells = True
    p1.Range("A52").Formula = "=+FORMULARIO!F60"
    for cel, ref in (("A61", "E170"), ("A62", "F170"), ("A63", "E171"), ("C63", "F171"), ("A64", "E172"),
                     ("C64", "F172")):
        p1.Range(cel).Formula = f"=FORMULARIO!{ref}"


def montar():
    if not os.path.exists(BASE):
        raise SystemExit(f"Modelo base não encontrado: {BASE}")
    excel = excel_fotos._abrir_excel()
    caso = None
    try:
        for nome, cfg in MODELOS.items():
            destino = os.path.join(PASTA, nome)
            shutil.copyfile(BASE, destino)
            wb = excel.Workbooks.Open(destino, UpdateLinks=0)
            try:
                f = wb.Worksheets("FORMULARIO")
                f.Range("F36").Value = cfg["tipo"]
                wb.Worksheets("PAG.2").Range("A47").Formula = "=+FORMULARIO!E126"  # 9ª linha de cômodos
                # como nos casos prontos da equipe (pedido do Rafael, 13/10/2026):
                f.Range("F47").Formula = "=F44+F43"  # área total = construída + comum
                f.Range("J139").Formula = "=J133"    # valor de operação = valor final de avaliação
                if cfg.get("liquidez"):
                    p1 = wb.Worksheets("PAG.1")
                    p1.Range("A31").Value = "Valor de Liquidez"
                    p1.Range("A32").Formula = '=+IF(A31="Valor de Mercado",FORMULARIO!J133,FORMULARIO!J135)'
                    # aviso automático no topo (A5:K5, já formatado em vermelho no modelo): aparece quando
                    # algum item restritivo do FORMULARIO (linhas 22-33) está marcado com X - como no caso
                    # 100005 (lá a fórmula olhava N30 em vez de M30; aqui conta a coluna toda)
                    p1.Range("R5").Value = "VIDE CARTA COM PARECER DO BANCO D \n(conforme normativo interno)"
                    p1.Range("A5").Formula = ('=IF(COUNTIF(FORMULARIO!G22:G33,"X")+COUNTIF(FORMULARIO!M22:M33,"X")>0,'
                                              'R5,"")')
                if cfg.get("reforma"):
                    if caso is None:
                        caso = excel.Workbooks.Open(CASO_REFORMA, ReadOnly=True, UpdateLinks=0)
                    _bloco_reforma(wb, caso)
                    excel.CutCopyMode = False
                wb.Activate()  # o caso de reforma aberto fica como pasta ativa
                wb.Worksheets("FORMULARIO").Activate()
                wb.Worksheets("FORMULARIO").Range("A1").Select()
                wb.Save()
            finally:
                wb.Close(SaveChanges=False)
            areas = {"FORMULARIO": "C1:P182"} if cfg.get("reforma") else {}  # área do caso de reforma
            impressao.restaurar(BASE, destino, areas=areas, log=lambda m: None)
            if cfg.get("reforma"):
                _escala(destino, "PAG.1", ESCALA_PAG1_REFORMA)
            print("montado:", nome)
    finally:
        if caso is not None:
            caso.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    montar()
