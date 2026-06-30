# Inventário TI — Flash Courier

App standalone (React + Vite) para controle de inventário da sala de TI, conectado à planilha "Estoque TI.xlsx" no SharePoint via Microsoft Graph.

## Correção MSAL v5

Este projeto usa `@azure/msal-browser` v5. Por isso, o retorno de autenticação deve usar uma página dedicada com `redirect-bridge`, e não uma página totalmente em branco.

Arquivos importantes:

- `redirect.html` — página dedicada de retorno do login Microsoft.
- `vite.config.js` — inclui `redirect.html` como segunda entrada no build do Vite.
- `src/App.jsx` — usa `https://SEU-DOMINIO/redirect.html` como `redirectUri`.

Se existir `public/blank.html` de uma tentativa anterior, pode excluir esse arquivo.

## Passo obrigatório no Azure / Microsoft Entra ID

No portal Microsoft Entra ID:

1. Vá em **App registrations**.
2. Abra o app com Client ID `37ff5e3e-1558-4add-b4e9-8e5c97e21943`.
3. Vá em **Authentication**.
4. Em **Single-page application**, cadastre exatamente:

```txt
https://inventario-ti-ten.vercel.app/redirect.html
```

Para teste local, cadastre também:

```txt
http://localhost:5173/redirect.html
```

Depois salve e faça um novo deploy no Vercel.

## Rodando localmente

```bash
npm install
npm run dev
```

A aplicação abre em:

```txt
http://localhost:5173
```

A página `redirect.html` só é usada pelo fluxo de autenticação. Se você abrir `/redirect.html` manualmente, verá apenas uma página simples de processamento.
