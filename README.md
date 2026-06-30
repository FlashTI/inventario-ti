# Arquivos corrigidos — SharePoint link direto + MSAL v5 + Graph

Este pacote mantém a correção do login com `redirect.html`, mantém os escopos de escrita do Microsoft Graph e altera a localização da planilha para usar o link real do SharePoint informado.

## Arquivos no pacote

- `src/App.jsx`
- `redirect.html`
- `vite.config.js`
- `README.md`

## Link da planilha usado no código

No `src/App.jsx`, foi adicionado:

```js
const EXCEL_FILE_WEB_URL = 'https://flashcouriercombr.sharepoint.com/:x:/r/sites/Suporte_Tcnico/Documentos%20Partilhados/ESTOQUE%20TI/Estoque%20TI.xlsx?d=w8665de340bf445ac826b93bcaae1bb16&csf=1&web=1&e=FONu2c';
```

Antes o app tentava localizar a planilha por site/path:

```txt
/sites/Suporte_Tcnico
/ESTOQUE TI/Estoque TI.xlsx
```

Agora ele resolve o arquivo pelo próprio link compartilhado usando:

```txt
/shares/{shareId}/driveItem
```

Depois usa o `driveId` e `itemId` retornados para acessar o workbook:

```txt
/drives/{driveId}/items/{itemId}/workbook
```

Isso evita erro por diferença de biblioteca/caminho, como `Documentos Partilhados`, `Shared Documents` ou biblioteca padrão do site.

## Permissões necessárias no Azure / Microsoft Entra ID

No App Registration `37ff5e3e-1558-4add-b4e9-8e5c97e21943`, mantenha em:

```txt
API permissions → Microsoft Graph → Delegated permissions
```

As permissões:

```txt
Files.ReadWrite.All
Sites.ReadWrite.All
User.Read
```

Depois clique em:

```txt
Grant admin consent
```

## Redirect URI

Mantenha estes Redirect URIs cadastrados em **Authentication → Single-page application**:

```txt
https://inventario-ti-ten.vercel.app/redirect.html
http://localhost:5173/redirect.html
```

## Importante sobre o erro 403

Mesmo com o link correto e os escopos corretos, o usuário logado precisa conseguir editar a planilha manualmente no SharePoint/Excel Online.

Teste com o mesmo usuário que faz login no app:

1. Abra a planilha pelo link.
2. Tente editar uma célula qualquer.
3. Salve.

Se abrir somente leitura ou pedir para solicitar acesso ao proprietário, o Graph continuará retornando:

```txt
EditModeAccessDenied
```

Nesse caso, o dono da planilha/site precisa dar permissão de **edição** para esse usuário ou grupo.
