# Inventário TI - Correção Graph 404 em tabelas vazias

Esta versão mantém o caminho do SharePoint:

- Site: `flashcouriercombr.sharepoint.com:/sites/Suporte_Tcnico`
- Arquivo: `ESTOQUE TI/Estoque TI.xlsx`

## O que foi ajustado

1. O app agora lista as tabelas do workbook antes de buscar linhas.
2. Se a tabela `Itens` ou `Movimentacoes` existir, mas estiver vazia, o erro `Graph 404 ItemNotFound` é tratado como lista vazia.
3. A inclusão de linhas passou a tentar o endpoint recomendado `/rows/add` e mantém fallback para `/rows`.
4. As mensagens de erro ficaram mais claras quando a tabela real do Excel não existe.

## Estrutura obrigatória da planilha

A planilha precisa ter tabelas reais do Excel, não apenas abas:

- Tabela: `Itens`
- Tabela: `Movimentacoes`

No Excel: clique dentro da tabela > Design da Tabela > Nome da Tabela.

## Permissões Azure

Manter permissões delegadas do Microsoft Graph:

- `Files.ReadWrite.All`
- `Sites.ReadWrite.All`
- `User.Read`

Com admin consent concedido.

## Redirect URI

Cadastrar em Authentication > Single-page application:

- `https://inventario-ti-ten.vercel.app/redirect.html`
- `http://localhost:5173/redirect.html`
