# Histórico do projeto

Este repositório é a versão de portfólio de um projeto real, desenvolvido em
ambiente corporativo entre 09/10/2026 e 15/10/2026. O repositório original é
privado e teve 19 commits de trabalho nesse período, com 6 versões publicadas
(Releases 2.1 a 2.6). O programa já chegou ao primeiro commit na versão 1.9: a
busca das fotos, a galeria e o Excel do primeiro banco foram feitos antes de o
projeto ir para o Git.

O histórico de commits que você vê aqui é **novo**: o código foi publicado em
poucos commits, depois de anonimizado. Este arquivo conta como o projeto de fato
evoluiu.

## O que foi anonimizado

- Nome da empresa e do seu sistema interno: aparece como "Central de Gestão".
- Intranet da empresa (onde ficam as requisições com a planilha do banco):
  aparece como "Intranet".
- Bancos clientes: "Banco C" e "Banco D".
- URLs, endpoints e a pasta da rede: trocados por endereços `example.com` e
  `\\servidor\pasta_publica`.
- Números de proposta e de casos citados em comentários: trocados por valores
  fictícios (100001…, 200001).
- Nome, registro e assinatura da engenheira responsável, e a logo da empresa:
  trocados por textos e imagens de exemplo.
- Modelos de Excel dos bancos e planilhas da equipe: substituídos por modelos
  fictícios com o mesmo layout (`ferramentas/gerar_modelos_exemplo.py`); os
  gabaritos do teste completo foram regravados com propostas fictícias, servidas
  por um servidor local de exemplo (`testes/servidor_exemplo.py`).

Por causa disso o código não roda contra os endereços de exemplo. A lógica, a
estrutura e a documentação são as do projeto real; o servidor de exemplo, os
geradores de modelos/imagens e os testes que usam o servidor foram acrescentados
nesta versão para o projeto rodar sem os sistemas da empresa.

## Como o projeto evoluiu

### 1. Ponto de partida: v1.9 (09/10)

O primeiro commit já trazia o programa usável:

- fotos buscadas direto na API da Central de Gestão, sem navegador, com o login
  guardado no PC (DPAPI do Windows) e o token reaproveitado;
- galeria para escolher, ordenar e renomear, com a aba de quadros de cada banco e
  seções novas criadas pela pessoa;
- um perfil por banco (`perfis.py`): o Banco C com os modelos Comparativo e
  Evolutivo, o Banco D com a grade de fotos montada na hora;
- logo da empresa, fotos comprimidas (Excel de no máximo 14 MB) e a configuração
  de impressão restaurada direto no arquivo (`impressao.py`), depois de descobrir
  que o Excel aberto por automação apagava a área de impressão ao salvar;
- `publicar.ps1`, que publica o `.exe` na pasta pública mesmo com a versão antiga
  aberta (troca agendada).

### 2. Formulário do Banco D e assinatura: v2.0 e v2.1 (09/10)

- formulário do Banco D pré-preenchido com os dados da vistoria. Foram geradas
  duas edições de teste (só os campos que a planilha do banco já trazia × tudo
  que a vistoria tem); a equipe aprovou a completa, que virou o padrão;
- limite das 13 linhas de cômodos do formulário, com aviso do que ficou de fora;
- assinatura da engenheira responsável no quadro da PAG.1;
- fotos do computador na galeria, além das da vistoria;
- o Excel do Banco D passou a se chamar `<nº>_FORMULARIO.xlsx`;
- Release no GitHub com o `.exe` (`publicar.ps1 -GitHub`).

### 3. Um modelo por tipo de operação: v2.2 (13/10)

- cinco modelos do Banco D (aquisição simples, à vista com FGTS, consórcio,
  consórcio com reforma, garantia), montados a partir do modelo base por
  `ferramentas/montar_modelos_banco_d.py`, que aplica só o que muda em cada um;
- correção de uma fórmula do modelo original que escondia a 9ª linha de cômodos,
  e as fórmulas que a equipe usava nos laudos prontos;
- pré-preenchimento conferido campo a campo contra laudos que a equipe tinha feito
  à mão: matrícula no campo certo, "19.220" como 19 mil, área construída pela
  privativa, CEP como número, quartos e banheiros agrupados;
- correção de um erro na geração quando a vistoria trazia cômodos vazios: o
  pré-preenchimento passou a nunca impedir o Excel com as fotos;
- publicação só com a versão no nome do `.exe`.

### 4. Modelo escolhido sozinho: v2.3 e v2.4 (14/10)

- o tipo de operação (e, com ele, o modelo) passou a ser lido na planilha que o
  banco anexa à requisição na Intranet, em paralelo com o download das fotos;
- da mesma planilha vêm a junção e o nome da junção do formulário.

### 5. Observações das fotos: v2.5 (14/10)

- as observações que o vistoriador escreve nas fotos aparecem ao buscar, com selo
  na galeria e faixa na foto grande.

### 6. Operação e testes: v2.6 (15/10)

- aviso de versão nova ao abrir (e a cada 30 minutos), lendo a pasta pública;
- erros inesperados e uso registrados na pasta pública, com fila local quando a
  rede está fora, e `ferramentas/resumo_uso.py` para resumir por dia, pessoa e
  banco;
- trabalho salvo por proposta, para retomar um laudo que volta para correção;
- testes rápidos (rodam antes de gerar o `.exe`) e teste completo com gabaritos
  (roda antes de publicar).

## Projetos relacionados

- [robo-cadastro-banco-a](https://github.com/rafaelbertochi1/robo-cadastro-banco-a) e
  [robo-cadastro-banco-b](https://github.com/rafaelbertochi1/robo-cadastro-banco-b):
  robôs que levam o status das inspeções para a Central de Gestão.
- [pipeline-laudos-banco-a](https://github.com/rafaelbertochi1/pipeline-laudos-banco-a) e
  [pipeline-laudos-banco-b](https://github.com/rafaelbertochi1/pipeline-laudos-banco-b):
  pipelines de download e extração dos laudos em PDF.
