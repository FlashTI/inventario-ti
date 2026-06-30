# Arquivos corrigidos — MSAL v5 + Graph 403 EditModeAccessDenied

Este pacote mantém a correção do login com `redirect.html` e ajusta as permissões do Microsoft Graph para tentativa de edição da planilha no SharePoint.

## Arquivos no pacote

- `src/App.jsx`
- `redirect.html`
- `vite.config.js`
- `README.md`

## Alteração principal

No arquivo `src/App.jsx`, os escopos foram alterados de:

```js
const GRAPH_SCOPES = ['Files.ReadWrite', 'Sites.Read.All', 'User.Read'];
```

para:

```js
const GRAPH_SCOPES = [
  'Files.ReadWrite.All',
  'Sites.ReadWrite.All',
  'User.Read',
];
```

## O que configurar no Azure / Microsoft Entra ID

No App Registration `37ff5e3e-1558-4add-b4e9-8e5c97e21943`, vá em:

```txt
API permissions → Add a permission → Microsoft Graph → Delegated permissions
```

Adicione:

```txt
Files.ReadWrite.All
Sites.ReadWrite.All
User.Read
```

Depois clique em:

```txt
Grant admin consent
```

Se você não for administrador do tenant, um administrador precisa aprovar essas permissões.

## Redirect URI

Mantenha estes Redirect URIs cadastrados em **Authentication → Single-page application**:

```txt
https://inventario-ti-ten.vercel.app/redirect.html
http://localhost:5173/redirect.html
```

## Permissão real no SharePoint

Mesmo com os escopos corretos, o usuário logado precisa ter permissão de edição no arquivo:

```txt
/ESTOQUE TI/Estoque TI.xlsx
```

Se o usuário abrir a planilha manualmente e ela estiver como somente leitura, o Graph continuará retornando erro 403.
