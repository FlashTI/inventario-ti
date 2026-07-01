# Alterações - Fluxo de estoque parametrizado

Arquivos alterados para o app de Inventário TI.

## Mudanças aplicadas

- Removida a tela **Retirados para uso**.
- A ação **Retirada** foi substituída por **Retirado do Estoque**.
- **Retirado do Estoque** remove a quantidade escolhida do estoque e mantém o registro apenas no histórico.
- Removido o campo **Responsável** das telas e modais. O histórico usa o usuário autenticado em **Registrado por**.
- Nos cards de item, o botão **Movimentar** foi reduzido e foi criado o botão **Adicionar** ao lado.
- O botão **Adicionar** abre o cadastro já preenchendo Nome, Categoria e Localização com base no card clicado.
- O campo **Nome do item** virou uma lista suspensa com os nomes parametrizados solicitados.
- A tela **Quantidade por categoria** agora exibe Disponível, Manutenção e Total.
- O Histórico removeu a coluna Responsável e permanece mostrando quem registrou a ação.

## Arquivos do pacote

- `src/App.jsx`
- `redirect.html`
- `vite.config.js`
- `README.md`

Substitua esses arquivos no projeto e faça novo deploy no Vercel.
