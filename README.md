# Ajuste — telas e fluxo de movimentação

Arquivos alterados para o app de Inventário TI.

## O que mudou

O sistema agora tem as telas:

- Itens disponíveis
- Retirados para uso
- Em manutenção
- Quantidade por categoria
- Histórico

## Regras de movimentação

### Itens disponíveis
Ações disponíveis:

- Retirada: move a quantidade informada para a tela **Retirados para uso**.
- Envio para manutenção: move a quantidade informada para a tela **Em manutenção**.
- Item quebrado/danificado: remove a quantidade informada do estoque e mantém o registro apenas no histórico.

### Retirados para uso
Ações disponíveis:

- Devolução: retorna a quantidade informada para **Itens disponíveis**.
- Item quebrado/danificado: remove a quantidade informada do estoque e mantém o registro apenas no histórico.

### Em manutenção
Ações disponíveis:

- Retorno da manutenção: retorna a quantidade informada para **Itens disponíveis**.
- Item quebrado/danificado: remove a quantidade informada do estoque e mantém o registro apenas no histórico.

## Agrupamento

As telas principais agrupam itens semelhantes e somam a quantidade para evitar vários cards repetidos.
Quando a movimentação é parcial, o sistema mantém o saldo na tela atual e cria/atualiza o item no destino.

## Histórico

O histórico mostra todas as movimentações, incluindo quantidade movimentada, origem, destino, responsável, usuário que registrou e observação.

## Arquivos

Copie estes arquivos por cima do projeto atual:

```txt
src/App.jsx
redirect.html
vite.config.js
README.md
```

Depois faça novo deploy no Vercel.

## Planilha

Não precisa alterar a estrutura da planilha. Ela precisa continuar com as tabelas reais do Excel:

```txt
Itens
Movimentacoes
ListasApp
```

Se a aba/lista de validação do Excel tiver status fixos, adicione também:

```txt
Retirado para uso
```

A ação **Item quebrado/danificado** não fica salva como status do item, porque o item é removido do estoque e fica apenas no histórico.
