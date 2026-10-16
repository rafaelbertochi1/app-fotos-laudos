# app-fotos-laudos

[![Testes](https://github.com/rafaelbertochi1/app-fotos-laudos/actions/workflows/testes.yml/badge.svg)](https://github.com/rafaelbertochi1/app-fotos-laudos/actions/workflows/testes.yml)

> **Sobre este repositório.** Versão de portfólio de um projeto real, desenvolvido em
> ambiente corporativo em 2026. Os nomes da empresa, dos sistemas e dos bancos, as URLs
> e os dados de exemplo foram trocados por nomes genéricos ("Central de Gestão",
> "Intranet", "Banco C", "Banco D"), então o código não roda contra os endereços de
> exemplo. Os modelos de Excel dos bancos, a logo e a assinatura foram substituídos por
> equivalentes fictícios, e um servidor local de exemplo faz o papel dos sistemas da
> empresa nos testes. O histórico de commits é novo e resumido; o caminho real do
> projeto está em [HISTORICO.md](HISTORICO.md).

Programa de Windows (**Python + Tkinter**, distribuído como `.exe`) usado pela equipe
de laudos de avaliação de imóveis: busca as fotos de uma vistoria na Central de
Gestão (o sistema interno da empresa), deixa a pessoa **escolher, ordenar e renomear**
as fotos numa galeria e **monta o Excel no modelo do banco** — fotos nos quadros
certos, legendas, páginas a mais quando precisa, impressão configurada e, no Banco D,
o formulário do banco já pré-preenchido com os dados da vistoria e a assinatura da
engenheira responsável.

**Por quê.** Antes, cada laudo começava com a pessoa baixando as fotos uma a uma,
redimensionando, colando foto por foto no modelo do banco, digitando a legenda
embaixo de cada uma e copiando os dados da vistoria para o formulário. Com dezenas de
fotos por proposta, era o trecho mais repetitivo do laudo. O programa faz isso em
segundos e a pessoa só decide o que entra e em que ordem.

**Na prática:** foram 6 versões publicadas (2.1 a 2.6) em cerca de uma semana
(09/10 a 15/10/2026), cada uma a partir do que a equipe pedia ao usar a anterior. O
`.exe` era distribuído numa pasta pública da rede; o próprio programa avisa quando há
versão nova lá, e registra os erros inesperados e o uso (uma linha por Excel gerado)
na mesma pasta, para o responsável acompanhar sem ninguém precisar mandar log.

## Funcionalidades

- **Busca direto na API** da Central de Gestão (sem navegador): login com token
  reaproveitado, download das fotos em paralelo, ordem e legendas da vistoria.
- **Galeria** para escolher (incluir/tirar), arrastar para reordenar, renomear e
  mandar cada foto para a aba principal ou para um quadro fixo (fachada, logradouro,
  número…); clique abre a foto grande com navegação pelo teclado.
- **Seções novas** criadas pela pessoa ("+ Nova seção…"): viram quadros a mais no
  Excel, em páginas copiadas do modelo (Banco C) ou montadas na hora (Banco D).
- **Fotos do computador** além das da vistoria (não vão para a Central de Gestão).
- **Observações das fotos** da vistoria: aviso ao buscar, selo "OBS" no cartão e
  faixa na foto grande.
- **Excel montado pelo próprio Excel** (pywin32): o modelo do banco sai idêntico, só
  com as fotos a mais — macros, validações e listas preservadas.
- **Páginas novas automáticas**: passou de 66 fotos, copia a página modelo; no
  Banco D, monta a grade de fotos do zero, só para as fotos usadas.
- **Impressão consertada direto no arquivo** (`impressao.py`): o Excel aberto por
  automação estraga a área de impressão e a escala ao salvar; o programa devolve a
  configuração do modelo e grava as áreas e quebras novas no XML.
- **Tamanho controlado**: fotos em JPEG redimensionadas para o quadro (o carimbo de
  data/GPS continua legível); se o Excel passar de 14 MB, refaz com mais compressão.
- **Banco D**: formulário pré-preenchido com o que a vistoria tem (endereço, áreas,
  vagas, região, terreno, cômodos, vistoriador…), traduzido para as listas do banco;
  **modelo escolhido sozinho** pelo tipo de operação, lido na planilha do banco
  anexada à requisição na Intranet; assinatura da responsável no quadro certo.
- **Trabalho salvo por proposta**: se o laudo volta para correção dias depois, o
  programa oferece retomar de onde parou (ordem, nomes, quadros, fotos do PC), mesmo
  que a vistoria tenha ganhado ou perdido fotos.
- **Logo da empresa** em toda célula do modelo escrita "Logo".
- **Modo sem janela** (`--proposta … --destino …`), usado pelo teste completo e no
  suporte.

## Um perfil por banco (`perfis.py`)

Um programa só, com o campo "Banco". Tudo o que muda de um banco para o outro fica
num `Perfil` — o resto (busca, galeria, login, Excel) é igual para todos:

| No perfil | Banco C | Banco D |
|---|---|---|
| modelos (`modelos/<pasta>/`) | Comparativo e Evolutivo (`.xlsm`) | um por tipo de operação (5, `.xlsx`) |
| aba principal (`AbaFotos`) | blocos prontos: 3 fileiras × 2 fotos, páginas de 50 linhas | sem blocos: grade montada na hora (`Grade`) |
| aba de quadros (`AbaLocais`) | "Mapa e fotos locais": 4 quadros; seção nova copia o bloco das linhas 29-61 | "PAG.1": 3 quadros com o texto dentro; seção nova vira página "FOTOS COMPLEMENTARES" (`PaginasExtras`) |
| foto do quadro também na aba principal | não | sim (`quadros_tambem_na_principal`) |
| pré-preenchimento (`PrePreenchimento`) | — | aba FORMULARIO (`formulario.py`) |
| assinatura (`Assinatura`) | — | PAG.1!F67:G68 |
| nome do arquivo | `planilha_<nº>_<tipo>.xlsm` | `<nº>_FORMULARIO.xlsx` |

**Banco novo** = uma pasta `modelos/<banco>/` com os modelos + um `Perfil` registrado
em `PERFIS`, com as linhas/colunas medidas no próprio modelo. Modelo novo de um banco
que já existe nem precisa de código: basta pôr o arquivo na pasta (o nome exibido sai
do arquivo, e no Banco D o tipo de operação é lido do próprio modelo). Para trocar um
modelo sem gerar o `.exe` de novo, uma pasta `modelos\<banco>\` ao lado do `.exe` tem
prioridade sobre os modelos embutidos.

## Arquitetura

```
app.py (janela) ─ galeria.py (cartões desenhados num Canvas)
   │
   └─ robo.py  buscar(): proposta → cache temporário com todas as fotos
   │           gerar():  fotos escolhidas → pasta da proposta + Excel
   ├─ central_gestao.py  API da Central de Gestão (login, vistoria, fotos, dados)
   ├─ intranet.py        tipo de operação / junção na planilha do Banco D
   ├─ excel_fotos.py     Excel por automação: fotos, grade, páginas, logo, assinatura
   │    ├─ formulario.py   dados da vistoria → células do formulário do banco
   │    └─ impressao.py    configuração de impressão restaurada direto no XML
   └─ perfis.py          o que muda por banco
```

| Arquivo | Para quê |
|---|---|
| `app.py`, `galeria.py`, `visual.py` | janela (Tkinter + tema Sun Valley). A galeria é **um Canvas** com os cartões desenhados, não widgets: com ~1.000 widgets a rolagem chegava a 128 ms por quadro; desenhado fica em ~9 ms. O nome e a seção são editados num campo/menu que só aparece no clique. |
| `robo.py` | fluxo de uma proposta (buscar → gerar), fotos do computador, sugestão dos quadros pela legenda, limite de 14 MB, modelo pelo tipo de operação |
| `central_gestao.py` | API da Central de Gestão, sem navegador (mapeada com `ferramentas/mapear_central_gestao.py`) |
| `intranet.py` | login em 2 etapas, busca da requisição e leitura da planilha anexada (lida direto do XML do `.xlsx`, sem Excel) |
| `excel_fotos.py` | inserção pelo Excel (pywin32); por que não openpyxl e os problemas já resolvidos estão comentados no topo |
| `formulario.py` | tradução dos dados da vistoria para o formulário do Banco D (listas, números "19.220", CEP, matrícula, cômodos agrupados…) |
| `impressao.py` | conserta área de impressão, escala e quebras no XML depois que o Excel salva |
| `perfis.py` | perfis dos bancos |
| `pasta_publica.py` | versão nova, erros e uso na pasta pública (fila local se a rede estiver fora) |
| `trabalho.py` | trabalho salvo por proposta (`%LOCALAPPDATA%\DownloaderFotos\trabalhos`) |
| `comum.py` | caminhos, versão, login (DPAPI do Windows) |
| `modelos/` | modelos de Excel embutidos no `.exe` (**fictícios** nesta versão) |
| `assets/` | ícone, logo e assinatura (**fictícios** nesta versão) |
| `ferramentas/` | geração dos modelos e imagens de exemplo, mapeamento da Central de Gestão, montagem dos modelos reais do Banco D, `.exe` de teste, resumo de uso |
| `testes/` | testes rápidos, teste completo com gabaritos e o servidor de exemplo |
| `dist/LEIA-ME.txt` | guia de quem usa (vai junto do `.exe` para a pasta pública) |

## Como rodar

Precisa de **Windows com Excel** (o Excel é montado pelo próprio Excel) e Python 3.12.

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

No uso real, a pessoa clica em "Configurar login…" (e "Configurar Intranet") e o
login fica guardado só naquele PC, criptografado pelo Windows (DPAPI). Como aqui os
endereços são de exemplo, para ver o programa funcionando suba o **servidor de
exemplo**, que responde como a Central de Gestão e a Intranet com propostas
fictícias (100001 a 100005 no Banco D, 200001 no Banco C):

```powershell
python testes\servidor_exemplo.py        # mostra as variáveis de ambiente a definir
# em outro terminal, depois de definir as variáveis mostradas:
python app.py
```

Sem janela (todas as fotos, quadros sugeridos pela legenda, log em
`%LOCALAPPDATA%\DownloaderFotos\logs`):

```powershell
python app.py --proposta 100002 --destino C:\temp --banco "Banco D"
```

O passo a passo de quem usa está em [`dist/LEIA-ME.txt`](dist/LEIA-ME.txt).

## Testes

```powershell
pip install -r requirements.txt pytest
pytest
```

`testes/test_unidade.py` cobre as regras que já deram problema de verdade: a
tradução dos dados da vistoria para o formulário (número com ponto de milhar,
cômodos agrupados, 13 linhas no máximo), o modelo achado pelo tipo de operação
escrito de outro jeito, os nomes de arquivo, a versão mais nova na pasta pública
("2.10" > "2.6"), erros e uso esperando a rede voltar, o trabalho salvo voltando
certo mesmo com a vistoria mudada, os quadros de cada perfil existindo nos modelos, a
impressão restaurada depois do "estrago" do Excel, e uma **busca de ponta a ponta
contra o servidor de exemplo** (login, fotos, observações, dados, Intranet). Rodam em
segundos, sem internet nem Excel; o GitHub Actions roda a mesma suíte a cada push no
Linux (selo no topo; a galeria Tkinter roda sob `xvfb`).

**Teste completo** (`testes/regressao.py`, Windows com Excel): gera o Excel de 7
propostas fixas pelo modo sem janela — os 5 modelos do Banco D, com o modelo
descoberto pela Intranet, e os 2 do Banco C com 74 fotos (70 na aba principal, o que força páginas novas) —
e compara com os gabaritos em `testes/gabaritos/` (modelo escolhido, fotos em cada
aba, configuração de impressão de cada aba, conteúdo das abas de fotos e do
formulário). No projeto real eram propostas de verdade; aqui o teste sobe o servidor
de exemplo e roda com uma pasta de dados temporária.

```powershell
python testes\regressao.py             # ~45 s
python testes\regressao.py --gravar    # aceita o resultado como gabarito novo (depois de conferir)
```

O `gerar_exe.bat` roda os testes rápidos e não gera o `.exe` se falharem; o
`publicar.ps1` roda o teste completo com o `.exe` novo e não publica se falhar.

## Gerar e publicar o .exe

- `gerar_exe.bat`: testes rápidos → PyInstaller (um arquivo só, com `modelos/` e
  `assets/` embutidos) → `dist\DownloaderFotos_v<versão>.exe`. A versão vem de
  `VERSAO` em `comum.py` e aparece na janela.
- `publicar.ps1`: teste completo com o `.exe` novo → copia para a pasta pública com a
  versão no nome e apaga as anteriores. Se alguém estiver com a versão antiga aberta
  (o Windows não deixa apagar `.exe` em uso), uma tarefa agendada termina a troca de
  5 em 5 minutos e se remove sozinha. Com `-GitHub`, cria/atualiza a Release
  `v<versão>` com o `.exe` e o LEIA-ME.
- `ferramentas/gerar_exe_teste.ps1`: gera duas edições marcadas do `.exe` (foi usado
  para a equipe comparar o pré-preenchimento "básico" com o "completo"; o completo
  foi o aprovado).

## O que é fictício nesta versão

| Arquivo | Como foi feito |
|---|---|
| `modelos/banco_c/*.xlsm`, `modelos/banco_d/*.xlsx` | `python ferramentas\gerar_modelos_exemplo.py` (openpyxl): mesmas abas, blocos mesclados, quadros, células do formulário, áreas/escalas de impressão e quebras que os perfis esperam, com textos genéricos. |
| `assets/*` | `python ferramentas\gerar_imagens_exemplo.py` (Pillow): ícone, logo "EMPRESA" e "ASSINATURA EXEMPLO". |
| `testes/dados_exemplo.py` + `servidor_exemplo.py` | propostas e fotos inventadas, servidas no formato da API real. |
| `testes/gabaritos/*.json` | gravados com `regressao.py --gravar` a partir dos itens acima. |

Os modelos reais dos bancos eram montados a partir das planilhas da equipe por
`ferramentas/montar_modelos_banco_d.py`, mantido como registro (as planilhas de
referência têm dados de clientes e ficaram fora).

## O que nunca commitar

O `.gitignore` já bloqueia: `referencia/` (planilhas reais da equipe, com dados de
clientes), `exploracao/` (mapeamento da Central de Gestão), `apagar/` (saídas de
teste), sessões/tokens/login (`sessao_*`, `token_*`, `login.dat`, `intranet.dat` — na
prática ficam em `%LOCALAPPDATA%`), `build/` e o `.exe`.

## Projetos relacionados

- [robo-cadastro-banco-a](https://github.com/rafaelbertochi1/robo-cadastro-banco-a) e
  [robo-cadastro-banco-b](https://github.com/rafaelbertochi1/robo-cadastro-banco-b):
  robôs que levam o status das inspeções para a Central de Gestão.
- [pipeline-laudos-banco-a](https://github.com/rafaelbertochi1/pipeline-laudos-banco-a) e
  [pipeline-laudos-banco-b](https://github.com/rafaelbertochi1/pipeline-laudos-banco-b):
  pipelines de download e extração dos laudos em PDF.
