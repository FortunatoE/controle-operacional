# Painel de Operações Agrícolas — Excel + HTML

Painel operacional para acompanhamento de atividades agrícolas, criado para cenários em que o Excel já está disponível no ambiente da fazenda, mas uma solução como Power BI Desktop não é prática para o usuário final.

O projeto usa o Excel como ponto único de manutenção dos dados e entrega um painel HTML local, com exportação de imagem PNG para compartilhamento rápido.

> **Privacidade:** este repositório não versiona planilhas, dados operacionais, arquivos gerados ou exportações. Os nomes e valores demonstrativos são fictícios.

## Problema resolvido

Administradores de fazenda precisam acompanhar o avanço das operações sem depender de instalação de uma nova ferramenta ou de chamados de TI.

O Power BI Desktop pode ser tecnicamente adequado, mas nem sempre é operacionalmente acessível para usuários que trabalham no campo ou em estações compartilhadas. A alternativa adotada foi uma aplicação híbrida:

- o usuário atualiza um único arquivo XLSM;
- o Excel consolida os indicadores;
- uma macro gera o arquivo web/dados.js;
- o navegador abre o dashboard HTML local;
- o painel pode ser exportado em PNG para WhatsApp, e-mail ou apresentações.

## Arquitetura

~~~
Excel (.xlsm)
RESUMO · PREMISSA · BASE_CONSOLIDADA · DASHBOARD
                       │
                       │ VBA
                       ▼
web/dados.js
Contrato de dados JavaScript gerado localmente
                       │
                       ▼
index.html + app.js
Dashboard HTML/CSS/JavaScript aberto no navegador
                       │
                       ├── Visualização local
                       └── Exportação PNG
~~~

## Funcionalidades

- Acompanhamento por operação: planejado, realizado e pendente;
- cálculo do percentual executado;
- contagem de talhões concluídos e pendentes;
- identificação automática da fazenda pelo sufixo do nome do arquivo;
- cadastro de nomes de fazenda na aba PREMISSA;
- dashboard responsivo executado localmente;
- data/hora da última geração;
- geração de PNG diretamente pelo navegador;
- geração de PNG via Edge em modo headless pelo VBA;
- tratamento de caracteres especiais em dados gerados pelo Excel.

## Estrutura do projeto

~~~
AGV/
├── DADOS/
│   └── OPERACOES - AGV.xlsm       # Local; ignorado pelo Git
├── MACROS/
│   ├── modFazendaUnica.bas        # Fluxo para uma fazenda
│   ├── modPainelWeb.bas           # Painel e exportação por Edge
│   └── modOperacoes.bas           # Rotinas complementares
├── web/
│   ├── index.html                 # Estrutura e estilos do painel
│   ├── app.js                     # Renderização, métricas e exportação PNG
│   ├── dados.example.js           # Exemplo do contrato de dados
│   ├── dados.js                   # Gerado pelo Excel; ignorado pelo Git
│   ├── selecao.js                 # Gerado para captura; ignorado pelo Git
│   ├── logo-embed.js              # Logo incorporada para exportação sem CORS
│   └── logoboaesperanca-branca.png
├── .gitignore
└── README.md
~~~

## Requisitos

### Estação de atualização

- Microsoft Excel para Windows com macros habilitadas;
- Power Query, caso a planilha utilize consultas;
- Microsoft Edge para exportação de PNG pelo VBA;
- acesso de leitura/gravação na pasta do projeto.

### Estação de consulta

- navegador moderno;
- acesso ao diretório compartilhado que contém a pasta web.

Não há necessidade de Power BI Desktop ou servidor web.

## Configuração do Excel

### Nome do arquivo e código da fazenda

O código é obtido pelos três últimos caracteres do nome do arquivo:

~~~
OPERACOES - AGV.xlsm  →  AGV
~~~

### Cadastro de fazendas

Na aba PREMISSA, preencha a tabela abaixo:

| Célula | Valor |
| --- | --- |
| R2 | CODIGO FAZENDA |
| S2 | NOME FAZENDA |
| R3 | AGV |
| S3 | Fazenda Agrovera |

O VBA envia codigoFazenda e fazenda ao arquivo dados.js. O cabeçalho do HTML e do PNG lê o mesmo dado.

### Importação dos módulos VBA

No editor VBA (Alt + F11):

1. Importe o módulo adequado da pasta MACROS.
2. Use modFazendaUnica.bas quando cada fazenda mantiver o seu próprio arquivo.
3. Use apenas um módulo que exponha GerarDadosWeb, pois modFazendaUnica.bas e modPainelWeb.bas possuem rotinas com esse mesmo nome.
4. Salve como XLSM e habilite macros ao abrir o arquivo.

## Fluxo de atualização

1. O operador atualiza os dados na aba RESUMO.
2. As consultas podem ser atualizadas somente quando houver alteração, evitando travamento durante a navegação.
3. A macro AtualizarPainelFazenda recalcula a pasta de trabalho, atualiza a BASE_CONSOLIDADA, gera web/dados.js e abre o painel.
4. app.js lê window.DASHBOARD_DATA, agrega os dados por operação e renderiza os cartões.
5. O usuário abre o painel no navegador ou baixa a imagem PNG.

## Contrato de dados

O arquivo web/dados.js é uma atribuição JavaScript, sem dependência de API ou servidor:

~~~
window.DASHBOARD_DATA = {
  geradoEm: "01/01/2026 08:00",
  origem: "OPERACOES - XXX.xlsm",
  codigoFazenda: "XXX",
  fazenda: "Fazenda Exemplo",
  registros: [
    {
      unidade: "XXX",
      operacao: "OPERACAO EXEMPLO",
      planejado: 1000,
      realizado: 650,
      pendente: 350,
      talhoesConcluidos: 8,
      talhoesPendentes: 4
    }
  ]
};
~~~

| Campo | Tipo | Descrição |
| --- | --- | --- |
| planejado | número | Área total prevista para a operação |
| realizado | número | Área já executada |
| pendente | número | Área ainda não executada |
| talhoesConcluidos | número | Talhões finalizados |
| talhoesPendentes | número | Talhões pendentes |

O percentual executado é calculado por realizado / planejado. O valor é limitado entre 0 e 1 para evitar dados visuais inválidos.

## Modo demonstração

Para apresentar o painel sem expor a realidade operacional, abra o arquivo index.html com o parâmetro demo=1:

~~~
file:///.../web/index.html?demo=1
~~~

Esse modo usa web/dados.demo.js, com fazenda, operações, hectares e talhões fictícios. O cabeçalho exibe MODO DEMONSTRAÇÃO e o fluxo operacional padrão continua usando web/dados.js.

## Exportação PNG

Existem dois modos:

1. **No navegador:** app.js desenha o layout em HTMLCanvasElement, converte para Blob e disponibiliza o download.
2. **No Excel:** modPainelWeb.bas usa o Edge em modo headless para gerar um screenshot do index.html.

A logo usada no Canvas é incorporada em logo-embed.js. Isso evita o erro de canvas contaminado (tainted canvas) causado por imagens locais incompatíveis com a exportação.

## Codificação de caracteres

O comando VBA Open For Output pode gravar arquivos em Windows-1252. Por compatibilidade, index.html declara:

~~~
<script charset="windows-1252" src="./dados.js"></script>
~~~

Além disso, a função Json do módulo VBA serializa caracteres não ASCII como sequências Unicode. Essa combinação preserva textos como GRADE 1° MÃO.

## Segurança e versionamento

Os seguintes itens são excluídos pelo .gitignore:

- planilhas XLSM e XLSX;
- web/dados.js, que pode conter dados de operação;
- web/selecao.js;
- diretório de exportações PNG;
- arquivos temporários do Office.

Antes de tornar um repositório público, revise também logos, nomes de clientes, caminhos de rede e capturas de tela.

## Limitações atuais

- O painel é local (file:///) e não possui autenticação;
- o navegador precisa acessar a pasta compartilhada;
- a rotina de exportação via VBA depende do Edge instalado;
- a edição do layout é feita diretamente em HTML, CSS e JavaScript;
- para uma central multi-fazendas, recomenda-se substituir o arquivo local por uma API, banco de dados ou repositório de dados controlado.

## Próximos passos sugeridos

- Adicionar validação de dados e mensagens de erro no Excel;
- produzir uma versão web hospedada em intranet;
- armazenar histórico de atualizações;
- implementar filtros por período e responsável;
- adicionar testes automatizados para o contrato de dados;
- disponibilizar um arquivo Excel modelo anonimizado.

## Licença

Uso interno. Defina uma licença antes de distribuir ou aceitar contribuições externas.
