# Sistema de Gestão de Patrimônio com Google Apps Script

## Visão Geral

Este projeto é uma aplicação web completa para gerenciamento de patrimônio (ativos), construída inteiramente sobre a plataforma Google Workspace. Ele utiliza **Google Apps Script** para o backend e lógica de negócios, e **Google Sheets** como banco de dados e fonte de configuração, oferecendo uma interface web interativa hospedada diretamente pelo Google.

O sistema foi desenvolvido para centralizar o controle de ativos, rastrear o histórico de movimentações e edições, fornecer visualizações de dados através de um dashboard e gerenciar o acesso de usuários, tudo dentro do ambiente familiar do Google Workspace.

## Funcionalidades Principais 

* **Autenticação Segura:** Login baseado em credenciais armazenadas em uma planilha, com diferentes níveis de acesso (Administrador, Editor, Leitor) que definem as permissões de CRUD.
* **Interface Web Intuitiva:** Permite adicionar, visualizar, editar e excluir itens do patrimônio através de modais e uma tabela dinâmica.
* **Geração Automática de ID:** IDs únicos são gerados automaticamente para novos itens com base em sua categoria.
* **Pesquisa e Filtragem Completa:** Funcionalidade de busca textual em todo o inventário e filtros avançados por ID, data de entrada, categoria, setor, status, responsável e valor.
* **Dashboard Interativo:** Gráficos (usando Chart.js) que exibem a distribuição dos ativos por setor, categoria, status e os nomes mais frequentes, com filtros específicos para o dashboard.
* **Rastreamento de Histórico:** Cada item possui um histórico detalhado de manutenções, movimentações, etc. Edições nos dados do item também são registradas automaticamente, mostrando o campo alterado, o valor antigo e o novo.
* **Relatórios em PDF:** Capacidade de gerar relatórios em PDF customizáveis, contendo os itens atualmente filtrados na tabela principal.
* **Controle de Concorrência e Sincronização:**
    * **Bloqueio de Edição:** Impede que dois usuários editem o *mesmo item* simultaneamente, utilizando uma aba na planilha para registrar o bloqueio.
    * **Indicador de Atualização (Polling):** Um indicador visual (🟢/🔴) no cabeçalho alerta o usuário se os dados exibidos estão desatualizados em relação ao servidor (planilha). Um botão "Atualizar" permite sincronizar manualmente. O sistema inteligentemente não mostra o alerta para o próprio usuário que acabou de fazer a alteração.
* **Configuração Dinâmica de Menus:** As opções disponíveis nos menus suspensos (Categoria, Setor, Status do Produto) são lidas diretamente de colunas dedicadas na planilha "Status", permitindo fácil personalização sem alterar o código.

## Tecnologias Utilizadas 

* **Google Apps Script:** Lógica de backend, manipulação da planilha, hospedagem da web app.
* **Google Sheets:** Utilizado como banco de dados principal e para configurações dinâmicas.
* **HTML:** Estrutura da interface web.
* **CSS (Tailwind CSS via CDN):** Estilização da interface web.
* **JavaScript (Client-side):** Interatividade da interface, comunicação com o backend Apps Script, renderização de gráficos com Chart.js.
* **Chart.js:** Biblioteca para criação dos gráficos do dashboard.

## Pré-requisitos 

* Uma Conta Google (para acesso ao Google Drive, Sheets e Apps Script).
* Familiaridade básica com Google Sheets.

## Instruções de Configuração 

Siga estes passos para configurar e implantar sua própria instância do sistema:

**1. Configuração da Planilha Google:**

* Crie uma nova Planilha Google no seu Google Drive.
* **Renomeie/Crie as seguintes abas (planilhas) exatamente como listado:**
    * `Patrimonio`:
        * Crie as colunas na primeira linha: `ID`, `Categoria`, `Nome`, `Responsavel`, `Setor`, `Status do Produto`, `Patrimonio`, `Data de Entrada`, `Valor Unitario`.
    * `Historico`:
        * Crie as colunas na primeira linha: `ID`, `Data`, `Tipo de Evento`, `Descricao`, `Custo`, `Responsavel`.
    * `Usuarios`:
        * Crie as colunas na primeira linha: `email`, `senha`, `nivel`.
        * Adicione pelo menos um usuário nas linhas seguintes, preenchendo o email, senha e nível (ex: `Administrador`, `Editor` ou `Leitor`).
    * `Status`:
        * Crie as colunas na primeira linha: `Status`, `Editor`, `Opcoes_Categoria`, `Opcoes_Setor`, `Opcoes_StatusProduto`, `ItemID_Lock`, `Editor_Lock`.
        * Na célula `A2`, coloque o valor `0`.
        * Preencha as opções desejadas para os menus suspensos nas colunas `Opcoes_Categoria` (a partir de C2), `Opcoes_Setor` (a partir de D2) e `Opcoes_StatusProduto` (a partir de E2). Deixe as colunas `ItemID_Lock` e `Editor_Lock` vazias a partir da linha 2.

**2. Configuração do Projeto Apps Script:**

* Abra a Planilha Google que você acabou de configurar.
* Vá em `Extensões` > `Apps Script`. Isso abrirá o editor de scripts.
* No editor, haverá um arquivo chamado `Code.gs` por padrão. Apague o conteúdo dele e cole **todo** o conteúdo do arquivo `code(Sincronizacao).js` deste repositório.
* Crie um novo arquivo HTML: Clique em `Arquivo` > `Novo` > `Arquivo HTML`. Dê o nome `index` (exatamente `index`, sem `.html`).
* Apague o conteúdo padrão do arquivo `index.html` recém-criado e cole **todo** o conteúdo do arquivo `index(Sincronizacao).html` deste repositório.
* Clique no ícone de disquete (Salvar projeto) e dê um nome ao seu projeto Apps Script (ex: "Gestão de Patrimônio App").

**3. Implantação como Aplicativo Web:**

* No editor do Apps Script, clique em `Implantar` (canto superior direito) > `Nova implantação`.
* Clique no ícone de engrenagem ao lado de "Selecione o tipo" e escolha `Aplicativo da Web`.
* Preencha as configurações
* Clique em `Implantar`.
* Após a implantação bem-sucedida, será exibida uma **URL do aplicativo da Web**. Copie esta URL. Ela é o link para acessar seu sistema!

---
