# Correção — Graph 404 ItemNotFound no SharePoint

Este pacote remove a dependência do endpoint `/shares/{shareId}/driveItem`, que retornou:

```txt
Graph 404: ItemNotFound
```

Agora o app resolve a planilha pelo caminho real do SharePoint:

```txt
Host: flashcouriercombr.sharepoint.com
Site: /sites/Suporte_Tcnico
Arquivo: ESTOQUE TI/Estoque TI.xlsx
```

A lógica nova faz:

1. Localiza o site:

```txt
/sites/flashcouriercombr.sharepoint.com:/sites/Suporte_Tcnico
```

2. Tenta encontrar o arquivo na biblioteca padrão:

```txt
/sites/{siteId}/drive/root:/ESTOQUE%20TI/Estoque%20TI.xlsx
```

3. Se não achar, lista todas as bibliotecas do site:

```txt
/sites/{siteId}/drives
```

4. Tenta localizar o arquivo em cada biblioteca:

```txt
/drives/{driveId}/root:/ESTOQUE%20TI/Estoque%20TI.xlsx
```

5. Depois usa o `driveId` e `itemId` encontrados para acessar o workbook:

```txt
/drives/{driveId}/items/{itemId}/workbook
```

## Arquivos no pacote

- `src/App.jsx`
- `redirect.html`
- `vite.config.js`
- `README.md`

## Permissões necessárias no Azure / Microsoft Entra ID

No App Registration `37ff5e3e-1558-4add-b4e9-8e5c97e21943`, mantenha as permissões delegadas:

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

Mantenha estes Redirect URIs em **Authentication → Single-page application**:

```txt
https://inventario-ti-ten.vercel.app/redirect.html
http://localhost:5173/redirect.html
```

## Atenção

Se depois desta correção voltar o erro `Graph 403 EditModeAccessDenied`, o caminho foi encontrado, mas o usuário logado não possui permissão de edição na planilha.

O usuário precisa conseguir abrir e editar manualmente:

```txt
https://flashcouriercombr.sharepoint.com/:x:/r/sites/Suporte_Tcnico/Documentos%20Partilhados/ESTOQUE%20TI/Estoque%20TI.xlsx
```
