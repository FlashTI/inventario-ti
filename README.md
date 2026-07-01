# Ajuste — movimentação por quantidade

Arquivos alterados para o app de Inventário TI.

## O que mudou

- A tela **Movimentar** agora tem o campo **Quantidade a movimentar**.
- O sistema valida se a quantidade informada é maior que a quantidade disponível no registro.
- Se a quantidade movimentada for igual ao total, o comportamento continua parecido com o antigo: o próprio item é atualizado.
- Se a quantidade movimentada for menor que o total, o sistema divide o registro:
  - mantém o item original com o saldo restante;
  - cria um novo item com a quantidade movimentada, novo status e responsável;
  - registra a movimentação no histórico.
- A quantidade movimentada é gravada na observação da tabela `Movimentacoes`, sem exigir alteração na estrutura da planilha.

## Arquivos

Copie estes arquivos por cima do projeto atual:

```txt
src/App.jsx
redirect.html
vite.config.js
README.md
```

Depois faça novo deploy no Vercel.

## Observação

A planilha existente pode continuar a mesma, desde que já tenha as tabelas reais do Excel:

```txt
Itens
Movimentacoes
ListasApp
```
