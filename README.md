# EstágioFlow

Interface web para consulta, edição e cadastro dos dados do programa de estágio **MS Multiplica**, do setor de sustentabilidade da Secretaria de Estado de Administração. Construído com Google Apps Script e publicado como Google Web App, o EstágioFlow funciona sobre a planilha do Google Sheets que a equipe já utiliza. Os servidores pesquisam, filtram e atualizam os registros de estagiários e mentores por uma tela simples, sem manipular a planilha diretamente e sem migrar os dados para outro local.

Projeto desenvolvido no âmbito do Programa de Extensão UFMS Digital.

## Instalação e publicação

### Arquivos

| Arquivo | Tipo no Apps Script | Função |
|---|---|---|
| `Code.gs` | Script | Leitura, edição e cadastro na planilha; configuração (`CONFIG`) |
| `Index.html` | HTML | Estrutura da página |
| `Estilos.html` | HTML | Visual (CSS) |
| `App.html` | HTML | Pesquisa, filtros, tabela e formulário (JavaScript) |
| `appsscript.json` | Manifesto | Fuso horário, permissões e modo de execução do Web App |

### 1. Instalar os scripts na planilha

1. Abra a planilha no Google Sheets e acesse **Extensões > Apps Script**.
2. Renomeie o projeto (canto superior esquerdo) para **EstágioFlow**.
3. Apague o conteúdo de `Código.gs` (ou `Code.gs`) e cole o conteúdo de `Code.gs`.
4. Clique em **+ > HTML** e crie três arquivos com os nomes exatos `Index`, `Estilos` e `App` (sem `.html`; o editor adiciona sozinho). Cole o conteúdo correspondente em cada um.
5. Em **Configurações do projeto** (ícone de engrenagem), marque **Mostrar arquivo de manifesto "appsscript.json" no editor**. Volte ao editor, abra `appsscript.json` e substitua o conteúdo.
6. Salve tudo (Ctrl+S).
7. No editor, selecione a função **`diagnostico`** e clique em **Executar**. Autorize o acesso quando solicitado. O **Registro de execução** deve listar as abas, a quantidade de registros e os campos encontrados. Se algo estiver errado aqui, estará errado no site também.

> Os nomes das abas precisam bater com `CONFIG` em `Code.gs`. Ajuste, se necessário

### 2. Publicar como Web App

1. No editor, clique em **Implantar > Nova implantação**.
2. Em **Selecionar tipo** (engrenagem), escolha **App da Web**.
3. Preencha:
   - **Descrição:** `Versão 1`
   - **Executar como:** **Usuário com acesso ao app da Web**
   - **Quem pode acessar:** **Qualquer pessoa com uma Conta do Google**
4. Clique em **Implantar** e copie a **URL do app da Web** (termina em `/exec`). Esse é o endereço a divulgar para os servidores.

#### Quem consegue ver e editar

Como o app executa com a conta de quem acessa, **o compartilhamento da própria planilha define as permissões**:

| Compartilhamento da planilha | No EstágioFlow |
|---|---|
| Editor | Consulta, edita e cadastra |
| Leitor | Só consulta (salvar retorna erro de permissão) |
| Sem acesso | Não consegue abrir os dados |

No primeiro acesso, cada usuário precisa autorizar o app. Se a conta dona do script não for de um domínio Workspace verificado, o Google mostra o aviso "O Google não verificou este app". Nesse caso, o usuário deve clicar em **Avançado > Acessar EstágioFlow (não seguro)**. O aviso é esperado para apps internos.

### 3. Atualizar depois de alterar o código

- **Para testar:** use **Implantar > Testar implantações**. Essa URL termina em `/dev`, sempre usa o código salvo mais recente e só funciona para quem edita o script.
- **Para publicar:** use **Implantar > Gerenciar implantações**, clique no lápis, escolha **Versão: Nova versão** e depois **Implantar**. A URL `/exec` continua a mesma.

### 4. Ajustes comuns (em `CONFIG`, no `Code.gs`)

| Chave | Para que serve |
|---|---|
| `tabela` | Colunas exibidas na lista (pelo nome do cabeçalho) |
| `filtros` | Campos que aparecem como filtros na lateral |
| `sugestoes` | Campos de texto que sugerem valores já usados |
| `obrigatorios` / `unico` | Campos exigidos no cadastro e campo que não pode se repetir (CPF) |
| `rotulos` | Nome exibido para uma coluna, no lugar do cabeçalho da planilha (ex.: `AA: 'CPF do supervisor'`) |
| `grupos` | Seções do formulário por intervalo de colunas. Se a linha 1 da planilha tiver o nome do grupo, ele é usado no lugar |
| `vinculo` | Liga o mentor aos estagiários pelo campo **Mentor** |
 
As opções das listas suspensas vêm da **validação de dados da primeira linha de dados** de cada aba. Se você alterar uma validação na planilha, o site acompanha automaticamente.

### 5. Problemas comuns

| Mensagem | Solução |
|---|---|
| `Aba "..." não encontrada` | Corrija `aba:` no `CONFIG` |
| `Planilha não encontrada` | O script não está vinculado à planilha. Preencha `ID_PLANILHA` e troque o escopo `spreadsheets.currentonly` por `spreadsheets` no manifesto |
| Erro de permissão ao salvar | O usuário não é **Editor** da planilha |
| "Registro foi alterado por outra pessoa" | Outra pessoa editou ou reordenou a planilha. Clique em **Atualizar dados** e repita |
| Datas com um dia de diferença | Confira o fuso em **Arquivo > Configurações** da planilha |

## Limitações atuais

- Não há exclusão de registros (por segurança, isso continua sendo feito na planilha).
