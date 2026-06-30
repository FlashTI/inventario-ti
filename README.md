# Correção da planilha do Inventário TI

Este pacote corrige o erro:

`tabela "Movimentacoes" não encontrada na planilha. Tabelas encontradas: ListasApp.`

## O que foi ajustado

- `Estoque TI.xlsx` foi recriada com tabelas reais do Excel:
  - `Itens`
  - `Movimentacoes`
  - `ListasApp`
- As tabelas `Itens` e `Movimentacoes` têm uma linha técnica com ID `__MODELO__` para o Microsoft Graph reconhecer a tabela como ativa.
- `src/App.jsx` foi ajustado para ignorar linhas com ID `__MODELO__`.

## Como usar

1. Substitua no SharePoint o arquivo atual por este arquivo:
   `Estoque TI.xlsx`
2. O caminho deve continuar sendo:
   `Documentos Partilhados / ESTOQUE TI / Estoque TI.xlsx`
3. Substitua o `src/App.jsx` no projeto pelo deste pacote.
4. Faça novo deploy no Vercel.

Não renomeie as tabelas do Excel. O app procura exatamente por `Itens` e `Movimentacoes`.
