APLICACAO HIBRIDA - ACOMPANHAMENTO DE OPERACOES
================================================

OBJETIVO
Excel/VBA continua sendo o motor de dados. O visual passa a ser HTML/CSS/JS local.
Isto elimina a perda de qualidade causada por CopyPicture/Chart.Export do Excel.

ESTRUTURA
Dashboard_Operacoes_Central.xlsm   <- seu arquivo macro habilitado
modOperacoes.bas                   <- modulo antigo de consolidacao
modPainelWeb.bas                   <- novo modulo deste pacote
web\index.html
web\app.js
web\dados.js                       <- gerado/atualizado pelo VBA
web\selecao.js                     <- gerado pelo VBA
EXPORTADOS\                        <- PNGs de alta qualidade

INSTALACAO
1. Coloque a pasta inteira em um local LOCAL do computador administrativo.
   Os arquivos de origem continuam em:
   H:\Producao BE\PRODUCAO\ACOMPANHAMENTO OPERACOES
   (ou no caminho configurado na aba CONFIG).

2. Use o seu Dashboard_Operacoes_Central.xlsm atual.
   Se estiver usando o .xlsx deste pacote, abra-o e salve como .xlsm.

3. ALT+F11 > Arquivo > Importar arquivo...
   Importe primeiro modOperacoes.bas (caso ainda nao exista).
   Remova o modulo parcial do painel e importe modPainelWeb_Completo.bas.

4. Execute uma vez a macro:
   InstalarPainelWeb

USO DIARIO
- ABRIR PAINEL WEB:
  Atualiza os arquivos encontrados na pasta configurada, gera web\dados.js e abre index.html no navegador.
  O painel permite escolher AGV, CED, PIR, etc., buscar e ordenar operacoes.

- EXPORTAR PNG HQ:
  Usa a unidade selecionada em DASHBOARD!B5.
  O VBA gera dados.js e usa o Microsoft Edge em modo headless para salvar
  uma imagem em alta resolucao diretamente em EXPORTADOS.

- PNG TODAS UNIDADES:
  Gera uma imagem separada para cada unidade encontrada na BASE_CONSOLIDADA.

QUALIDADE
A exportacao web usa renderizacao do Chromium/Edge com escala 2x.
Nao passa por CopyPicture do Excel.

IMPORTANTE
- Mantenha a pasta web ao lado do .xlsm.
- O arquivo web\dados.js entregue no pacote reflete a BASE_CONSOLIDADA atual.
  Nos usos seguintes, a macro GerarDadosWeb o substituirá pelos dados recém-atualizados.
- O Edge precisa estar instalado (padrao no Windows 10/11 corporativo).
- O painel nao precisa de internet, Node, Vercel, servidor ou bibliotecas externas.
- Para maior confiabilidade, deixe esta aplicacao em disco local do PC. Apenas os
  arquivos OPERAÇÕES ficam no H: da empresa.
