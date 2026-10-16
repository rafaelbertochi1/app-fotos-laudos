"""
Testes rápidos (segundos, sem internet e sem Excel) das regras do programa.

    pytest                                 (ou: python -m unittest discover -s testes -v)

Rodam sozinhos no gerar_exe.bat e no GitHub Actions (no Linux; a galeria
precisa de tela - no CI, via xvfb). O teste completo (busca + Excel de
verdade, só no Windows) é o testes\regressao.py.
"""

import os
import shutil
import sys
import tempfile
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
sys.path.insert(0, os.path.join(RAIZ, "testes"))
import conftest  # noqa: E402,F401 - pasta de dados temporária também pelo unittest (sem pytest)

import central_gestao  # noqa: E402
import excel_fotos  # noqa: E402
import formulario  # noqa: E402
import impressao  # noqa: E402
import intranet  # noqa: E402
import pasta_publica  # noqa: E402
import regressao  # noqa: E402
import robo  # noqa: E402
import servidor_exemplo  # noqa: E402
import trabalho  # noqa: E402
from perfis import BANCO_C, BANCO_D, PERFIS  # noqa: E402


class Formulario(unittest.TestCase):
    def test_numeros_da_central_gestao(self):
        self.assertEqual(formulario._num("167,10"), 167.1)
        self.assertEqual(formulario._num("178.12"), 178.12)
        self.assertEqual(formulario._num("1.234,5"), 1234.5)
        self.assertEqual(formulario._num("19.220"), 19220)  # milhar (caso 100002)
        self.assertIsNone(formulario._num("abc"))
        self.assertIsNone(formulario._informado(0))  # 0 = não informado: em branco

    def test_padroes_fixos(self):
        v, _ = formulario.valores_banco_d({"banco": {"juncao": "1234", "nome_juncao": "AG X"}})
        self.assertEqual(v["F5"], 1234)
        self.assertEqual(v["F6"], "AG X")
        self.assertEqual(v["H34"], "2")  # livro sempre 2
        self.assertEqual(v["F61"], "=TODAY()")
        self.assertEqual(v["I48"], "=J43-F43")
        self.assertEqual(v["F47"], "=F44+F43")
        self.assertEqual((v["F62"], v["F63"], v["F64"]), ("ENGENHEIRA RESPONSÁVEL EXEMPLO", 1234567890, "SÃO PAULO"))

    def test_comodos_param_nas_13_linhas(self):
        comodos = [{"divisaoInterna": f"Cômodo {n}", "quantidade": 1} for n in range(20)]
        v, feitos = formulario.valores_banco_d({"comodos": comodos})
        self.assertIn("E130", v)
        self.assertNotIn("E131", v)  # nunca cria linha nova (118-120 fixas + 121-130 livres = 13)
        self.assertTrue(any("10 cômodo(s) não couberam" in f for f in feitos))

    def test_suites_e_banheiros_agrupados(self):
        comodos = [{"divisaoInterna": "Dormitório", "quantidade": 1}, {"divisaoInterna": "Dormitório", "quantidade": 2},
                   {"divisaoInterna": "Banheiro", "quantidade": 1}, {"divisaoInterna": "Banheiro Suíte", "quantidade": 1},
                   None]  # item vazio veio na proposta 100006
        v, _ = formulario.valores_banco_d({"comodos": comodos})
        self.assertEqual(v["H118"], 3)
        self.assertEqual(v["H120"], 1)


class Modelos(unittest.TestCase):
    def test_modelos_embutidos(self):
        self.assertEqual(set(robo.listar_modelos(PERFIS["Banco C"])), {"Comparativo", "Evolutivo"})
        self.assertEqual(len(robo.listar_modelos(PERFIS["Banco D"])), 5)

    def test_tipo_de_operacao_acha_o_modelo(self):
        esperado = {
            "AQUISIÇÃO SIMPLES": "Aquisição simples",
            "aquisicao a vista com fgts": "Aquisição à vista FGTS",
            "CONSÓRCIO (IMÓVEL URBANO)": "Consórcio urbano",
            "CONSÓRCIO – REFORMA S/AMPLIAÇÃO (IMÓVEL URBANO)": "Consórcio reforma sem ampliação",
            "GARANTIA (IMÓVEL URBANO)": "Garantia imóvel urbano",
        }
        banco_d = PERFIS["Banco D"]
        for tipo, modelo in esperado.items():
            with self.subTest(tipo=tipo):
                self.assertEqual(robo.modelo_por_tipo(tipo, banco_d), modelo)
        self.assertIsNone(robo.modelo_por_tipo("TIPO QUE NÃO EXISTE", banco_d))

    def test_nome_do_arquivo(self):
        banco_c = robo.listar_modelos(PERFIS["Banco C"])["Comparativo"]
        banco_d = robo.listar_modelos(PERFIS["Banco D"])["Aquisição simples"]
        destino = os.path.join("pasta", "laudos")
        self.assertEqual(robo.caminhos_saida("123", destino, banco_c, PERFIS["Banco C"])[2],
                         os.path.join(destino, "Proposta_123", "planilha_123_comparativo.xlsm"))
        self.assertEqual(robo.caminhos_saida("123", destino, banco_d, PERFIS["Banco D"])[2],
                         os.path.join(destino, "Proposta_123", "123_FORMULARIO.xlsx"))

    def test_quadros_dos_perfis_existem_nos_modelos(self):
        """Cada bloco que o perfil espera está mesclado no modelo (senão
        excel_fotos._colocar_foto recusa: "O modelo não tem o quadro...")."""
        for perfil in PERFIS.values():
            for nome, caminho in robo.listar_modelos(perfil).items():
                with self.subTest(banco=perfil.banco, modelo=nome):
                    wb = _abrir_xlsx(caminho)
                    self.assertTrue({q.bloco for q in perfil.locais.quadros} <= _mescladas(wb[perfil.locais.nome]))
                    if perfil.assinatura:
                        self.assertIn(perfil.assinatura.area, _mescladas(wb[perfil.assinatura.aba]))
                    if not perfil.fotos.grade:  # modelo com os blocos de foto prontos (Banco C)
                        lay = perfil.fotos
                        blocos = {b for b, _, _ in excel_fotos.slots_fotos(lay.paginas_no_modelo, lay)}
                        self.assertTrue(blocos <= _mescladas(wb[lay.nome]))


def _abrir_xlsx(caminho):
    import warnings

    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return openpyxl.load_workbook(caminho)


def _mescladas(ws):
    return {str(m) for m in ws.merged_cells.ranges}


def _quieto(_mensagem):
    pass


class ServidorDeExemplo(unittest.TestCase):
    """Busca de ponta a ponta (sem Excel) contra o servidor de exemplo: login
    na Central de Gestão, fotos, dados da vistoria, Intranet e modelo pelo
    tipo de operação - o mesmo caminho do programa, com dados fictícios."""

    @classmethod
    def setUpClass(cls):
        cls.servidor, url = servidor_exemplo.iniciar()
        cls.antes = (central_gestao.API, intranet.BASE, dict(os.environ))
        central_gestao.API = intranet.BASE = url
        os.environ.update(servidor_exemplo.variaveis(url))
        central_gestao.apagar_token()
        intranet.apagar_sessao()

    @classmethod
    def tearDownClass(cls):
        cls.servidor.shutdown()
        central_gestao.API, intranet.BASE, ambiente = cls.antes
        os.environ.clear()
        os.environ.update(ambiente)

    def test_busca_fotos_e_dados(self):
        proposta, itens, cache = robo.buscar("100.002", log=_quieto)
        try:
            self.assertEqual(proposta, "100002")
            self.assertEqual(len(itens), 18)
            self.assertTrue(all(it["caminho"] and not it["erro"] for it in itens))
            self.assertEqual(itens[0]["legenda"], "Fachada")
            self.assertEqual(itens[5]["legenda"], "Sala (1)")  # ambiente com várias fotos: numeradas
            self.assertEqual({it["observacao"] for it in itens if it["observacao"]},
                             {"Box quebrado.", "Fechamento em vidro não averbado."})
            dados = robo.dados_da_busca(cache)
            self.assertNotIn("cpf", dados["engenheiro"])  # do vistoriador, só nome e registro
            v, _ = formulario.valores_banco_d(dados)
            self.assertEqual(v["F17"], 1000000)            # CEP como número (célula 00000-000)
            self.assertEqual(v["J34"], "0098765")          # matrícula com 7 dígitos
            self.assertEqual(v["F42"], 19220)              # "19.220" = 19 mil
            self.assertEqual(v["F43"], 68.4)               # construída = privativa
            self.assertEqual((v["H118"], v["H119"], v["H120"]), (2, 1, 1))  # banheiro da suíte não conta
            self.assertEqual(v["F164"], "SIM")
            self.assertEqual(robo.sugerir_locais(itens, BANCO_C),
                             {"fachada": 0, "identificacao": 1, "logradouro_esq": 2, "logradouro_dir": 3})
        finally:
            shutil.rmtree(cache, ignore_errors=True)

    def test_foto_antiga_sem_ambiente_usa_o_nome_do_arquivo(self):
        _, itens, cache = robo.buscar("100003", log=_quieto)
        shutil.rmtree(cache, ignore_errors=True)
        self.assertEqual([it["legenda"] for it in itens[:4]], ["Fachada", "PlacaDaRua", "Rua (1)", "Rua (2)"])

    def test_proposta_que_nao_existe(self):
        with self.assertRaises(central_gestao.ErroCentralGestao):
            robo.buscar("999999", log=_quieto)

    def test_modelo_pelo_tipo_de_operacao_na_intranet(self):
        dados = robo.descobrir_modelo("100004", BANCO_D, log=_quieto)
        self.assertEqual(dados["modelo"], "Consórcio reforma sem ampliação")  # escrito com travessão e sem acento
        self.assertEqual((dados["juncao"], dados["fonte"]), (4567, "Intranet"))
        with self.assertRaises(ValueError):
            robo.descobrir_modelo("200001", BANCO_D, log=_quieto)  # sem requisição do Banco D

    def test_senha_errada(self):
        os.environ["CENTRAL_GESTAO_SENHA"] = "errada"
        central_gestao.apagar_token()
        try:
            with self.assertRaises(central_gestao.ErroLogin):
                central_gestao.obter_token(log=_quieto)
        finally:
            os.environ["CENTRAL_GESTAO_SENHA"] = servidor_exemplo.dados_exemplo.SENHA


class Impressao(unittest.TestCase):
    """impressao.restaurar conserta o que o Excel de automação estraga ao salvar."""

    def test_restaura_area_escala_e_quebras(self):
        import re
        import zipfile

        modelo = robo.listar_modelos(BANCO_C)["Comparativo"]
        tmp = tempfile.mkdtemp()
        saida = os.path.join(tmp, "gerado.xlsm")
        try:
            # simula o estrago: "_xlnm.Print_Area" vira nome comum e a escala muda
            with zipfile.ZipFile(modelo) as z, zipfile.ZipFile(saida, "w") as novo:
                for info in z.infolist():
                    dados = z.read(info.filename)
                    if info.filename == "xl/workbook.xml":
                        dados = dados.replace(b'name="_xlnm.Print_Area"', b'name="Print_Area"')
                    elif info.filename.startswith("xl/worksheets/sheet"):
                        dados = re.sub(rb'scale="\d+"', b'scale="26"', dados)
                    novo.writestr(info, dados)
            self.assertTrue(impressao.conferir(modelo, saida)[0])  # estragado: tem diferença
            impressao.restaurar(modelo, saida, areas={BANCO_C.fotos.nome: "A1:O602"},
                                quebras={BANCO_C.fotos.nome: [553, 603]}, log=_quieto)
            abas = regressao.resumo_excel(saida, [])["abas"]
            fotos = abas[BANCO_C.fotos.nome]
            self.assertEqual(fotos["area_impressao"], "$A$1:$O$602")
            self.assertEqual(fotos["impressao"]["scale"], "85")
            self.assertEqual(fotos["quebras"], 6)  # as 4 do modelo + as 2 recusadas
            self.assertEqual(abas[BANCO_C.locais.nome]["area_impressao"], "$A$1:$O$61")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class Gabaritos(unittest.TestCase):
    """Os gabaritos do teste completo batem com os modelos e os casos."""

    def test_um_gabarito_por_caso_com_as_abas_do_modelo(self):
        import json

        for caso in regressao.CASOS:
            with self.subTest(caso=regressao.nome_caso(caso)):
                with open(os.path.join(regressao.GABARITOS, regressao.nome_caso(caso) + ".json"),
                          encoding="utf-8") as f:
                    gabarito = json.load(f)
                perfil = PERFIS[caso["banco"]]
                self.assertEqual(gabarito["modelo"], caso["esperado"])
                modelo = robo.listar_modelos(perfil)[caso["esperado"]]
                self.assertEqual(set(gabarito["abas"]), set(_abrir_xlsx(modelo).sheetnames))
                self.assertEqual(set(gabarito["celulas"]), set(regressao.abas_de_conteudo(perfil)))

    def test_diferencas(self):
        self.assertEqual(regressao.diferencas({"a": 1, "b": {"c": 2}}, {"a": 1, "b": {"c": 2}}), [])
        self.assertEqual(regressao.diferencas({"a": 1, "b": {"c": 2}}, {"b": {"c": 3}, "d": 4}),
                         ["a: 1 -> (sumiu)", "b / c: 2 -> 3", "d: (não tinha) -> 4"])


class PastaPublica(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.antes = {n: getattr(pasta_publica, n) for n in ("PASTA_PUBLICA", "PASTA_ERROS", "PASTA_USO",
                                                              "A_ENVIAR", "ATIVO", "enviar_em_segundo_plano")}
        pasta_publica.PASTA_PUBLICA = os.path.join(self.tmp, "publica")
        pasta_publica.PASTA_ERROS = os.path.join(pasta_publica.PASTA_PUBLICA, "_erros")
        pasta_publica.PASTA_USO = os.path.join(pasta_publica.PASTA_PUBLICA, "_uso")
        pasta_publica.A_ENVIAR = os.path.join(self.tmp, "a_enviar")
        pasta_publica.ATIVO = True
        pasta_publica.enviar_em_segundo_plano = lambda: None  # o teste manda na hora (enviar)

    def tearDown(self):
        for n, v in self.antes.items():
            setattr(pasta_publica, n, v)
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_versao_mais_nova(self):
        pasta = os.path.join(self.tmp, "exes")
        os.makedirs(os.path.join(pasta, "_pendente"))
        for nome in ("DownloaderFotos_v2.5.exe", "DownloaderFotos_v2.10.exe", "LEIA-ME.txt"):
            open(os.path.join(pasta, nome), "w").close()
        self.assertEqual(pasta_publica.versao_mais_nova(pasta, "2.6")[0], "2.10")  # 2.10 > 2.6 (não é texto)
        self.assertIsNone(pasta_publica.versao_mais_nova(pasta, "2.10"))
        self.assertIsNone(pasta_publica.versao_mais_nova(os.path.join(self.tmp, "nao_existe"), "1.0"))

    def test_erro_e_uso_esperam_a_pasta_e_vao_depois(self):
        pasta_publica.contexto.update(proposta="123", banco="Banco D")
        try:
            raise ZeroDivisionError("teste")
        except ZeroDivisionError:
            pasta_publica.erro("deu ruim")
        pasta_publica.uso(banco="Banco D", proposta="123", fotos_no_excel=10)
        self.assertEqual(pasta_publica.enviar(), 0)  # pasta pública fora do ar: espera
        os.makedirs(pasta_publica.PASTA_PUBLICA)
        self.assertEqual(pasta_publica.enviar(), 2)
        erros = os.listdir(pasta_publica.PASTA_ERROS)
        self.assertEqual(len(erros), 1)
        with open(os.path.join(pasta_publica.PASTA_ERROS, erros[0]), encoding="utf-8") as f:
            texto = f.read()
        self.assertIn("ZeroDivisionError", texto)
        self.assertIn("Proposta: 123", texto)
        [csv] = os.listdir(pasta_publica.PASTA_USO)
        with open(os.path.join(pasta_publica.PASTA_USO, csv), encoding="utf-8-sig") as f:
            linhas = f.read().splitlines()
        self.assertEqual(linhas[0].split(";")[:3], ["data", "hora", "usuario"])
        self.assertIn(";Banco D;", linhas[1])
        pasta_publica.uso(banco="Banco C", proposta="456")  # segunda linha: sem repetir o cabeçalho
        pasta_publica.enviar()
        with open(os.path.join(pasta_publica.PASTA_USO, csv), encoding="utf-8-sig") as f:
            self.assertEqual(len(f.read().splitlines()), 3)

    def test_desligado_nao_grava_nada(self):
        pasta_publica.ATIVO = False
        pasta_publica.uso(banco="Banco C")
        self.assertFalse(os.path.exists(pasta_publica.A_ENVIAR))


class Trabalho(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.antes, trabalho.PASTA = trabalho.PASTA, self.tmp

    def tearDown(self):
        trabalho.PASTA = self.antes
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_salva_le_e_limpa_fotos_do_pc(self):
        foto = os.path.join(self.tmp, "fachada.jpg")
        with open(foto, "wb") as f:
            f.write(b"x")
        copia = trabalho.guardar_foto_pc("123", foto)
        copia2 = trabalho.guardar_foto_pc("123", foto)  # mesmo nome: não sobrescreve
        self.assertNotEqual(copia, copia2)
        estado = {"ordem": ["u1", "pc:" + os.path.basename(copia)], "fora": [], "nomes": {}, "destinos": {}}
        trabalho.salvar("123", estado, banco="Banco D")
        self.assertEqual(trabalho.ler("123")["ordem"], estado["ordem"])
        self.assertFalse(os.path.exists(copia2))  # a que o estado não usa sai
        self.assertEqual(trabalho.fotos_pc_salvas("123", estado), [copia])
        trabalho.salvar("123", estado, gerado_em="2026-10-15T10:00:00")
        trabalho.salvar("123", estado)  # salvar de novo não perde a data do Excel
        self.assertEqual(trabalho.ler("123")["gerado_em"], "2026-10-15T10:00:00")
        self.assertIn("15/10/2026", trabalho.resumo(trabalho.ler("123"))[0])
        self.assertEqual(trabalho.ler("999"), {})


class GaleriaEstado(unittest.TestCase):
    """O estado salvo volta certo mesmo com a vistoria mudada na Central de Gestão."""

    def test_ida_e_volta(self):
        import tkinter as tk

        from galeria import Galeria

        try:
            raiz = tk.Tk()
        except tk.TclError:  # sem tela (CI sem xvfb)
            self.skipTest("sem tela para o Tkinter")
        raiz.withdraw()
        try:
            perfil = PERFIS["Banco D"]
            itens = [{"indice": i, "legenda": f"Foto {i}", "caminho": f"x{i}.jpg", "erro": None, "chave": f"u{i}"}
                     for i in range(5)]
            g = Galeria(raiz, perfil=perfil, secoes_extras=["Fundos"])
            g.carregar(itens, {}, {})
            quadro = next(iter(perfil.rotulos_locais()))
            g.mover(3, -3)                                  # u3 vai pro começo
            g.cartoes[1].alternar()                          # u1 fora do Excel
            g.cartoes[2].var_legenda.set("Cozinha nova")     # u2 renomeada
            g.cartoes[4].var_destino.set(g.destino_da_chave(quadro))
            g.destino_mudou(4)                               # u4 num quadro da PAG.1
            g.cartoes[0].var_destino.set(g.destino_da_chave("extra:Fundos"))
            g.destino_mudou(0)                               # u0 numa seção nova
            estado = g.estado()
            self.assertEqual(estado["ordem"], ["u3", "u0", "u1", "u2", "u4"])

            # outra busca: a Central de Gestão perdeu a u2 e ganhou a u9; a seção "Fundos" foi apagada neste PC
            novos = [it for it in itens if it["chave"] != "u2"] + [
                {"indice": 5, "legenda": "Nova", "caminho": "x9.jpg", "erro": None, "chave": "u9"}]
            g2 = Galeria(raiz, perfil=perfil, secoes_extras=[])
            g2.carregar(novos, {}, {})
            retomadas, sumidas, novas = g2.aplicar_estado(estado)
            self.assertEqual((retomadas, sumidas, novas), (4, 1, 1))
            self.assertEqual([g2._chave(i) for i in g2.ordem], ["u3", "u0", "u1", "u4", "u9"])
            self.assertFalse(g2.cartoes[1].var_incluir.get())
            self.assertEqual(g2.cartoes[4].chave_local, quadro)
            self.assertEqual(g2.cartoes[0].chave_local, "extra:Fundos")  # seção criada de novo
            internas, locais = g2.selecao()
            self.assertIn(4, internas)  # Banco D: foto do quadro também na FOTOS
            self.assertEqual(locais[quadro], 4)
        finally:
            raiz.destroy()


if __name__ == "__main__":
    unittest.main()
