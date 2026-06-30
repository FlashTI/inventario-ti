# Inventário TI — Flash Courier

App standalone (React + Vite) para controle de inventário da sala de TI, conectado à planilha "Estoque TI.xlsx" no SharePoint via Microsoft Graph.

## Deploy no Vercel (recomendado, gratuito)

1. Crie uma conta em https://vercel.com (pode entrar com GitHub, GitLab ou e-mail).
2. Suba esta pasta para um repositório no GitHub (ou GitLab/Bitbucket):
   ```
   git init
   git add .
   git commit -m "Inventário TI inicial"
   git remote add origin <URL_DO_SEU_REPOSITORIO>
   git push -u origin main
   ```
3. No painel do Vercel, clique em **Add New → Project**, selecione o repositório.
4. Framework Preset: o Vercel detecta "Vite" automaticamente. Não precisa mudar nada (Build Command: `npm run build`, Output Directory: `dist`).
5. Clique em **Deploy**. Em ~1 minuto você terá uma URL tipo `https://inventario-ti-flash-courier.vercel.app`.

## Passo OBRIGATÓRIO depois do deploy: atualizar o Azure AD

O login da Microsoft só funciona em domínios cadastrados como "Redirect URI". Depois do deploy:

1. Copie a URL final do Vercel (ex: `https://inventario-ti-flash-courier.vercel.app`).
2. Vá em https://entra.microsoft.com → **Registros de aplicativo** → seu app → **Autenticação**.
3. Em **Redirecionamentos da Web — Single-page application**, clique em **Adicionar URI** e cadastre também a URL de autenticação em branco:
   ```
   https://SEU-DOMINIO-VERCEL/blank.html
   ```
   Exemplo:
   ```
   https://inventario-ti-flash-courier.vercel.app/blank.html
   ```
4. Para teste local, cadastre também:
   ```
   http://localhost:5173/blank.html
   ```
5. Salve.

Sem esse passo, o login pode falhar com `redirect_uri mismatch`, CORS ou erro de pop-up aninhado no MSAL.

## Domínio próprio (opcional)

No painel do projeto no Vercel, em **Settings → Domains**, você pode apontar um domínio próprio da empresa (ex: `inventario.flashcourier.com.br`), criando o CNAME indicado pelo Vercel no DNS do domínio. Depois, repita o passo acima de adicionar essa URL nova como Redirect URI no Azure AD.

## Rodando localmente (opcional, para testar antes de publicar)

```
npm install
npm run dev
```
Abre em `http://localhost:5173`. Lembre de adicionar `http://localhost:5173/blank.html` como Redirect URI no Azure AD se quiser testar o login localmente.

## Estrutura

- `src/App.jsx` — toda a aplicação (UI + integração com Microsoft Graph/Excel)
- `index.html` — carrega o Tailwind via CDN
- A biblioteca MSAL (login Microsoft) fica nas dependências do projeto (`@azure/msal-browser`).
