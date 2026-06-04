# Dash de Carga no Cloudflare Worker

Este pacote foi feito para o projeto que aparece como:

```txt
dash-de-carga.hhhthiti.workers.dev
Deploy command: npx wrangler deploy
```

Ele usa Worker com static assets:

```txt
src/index.js              Worker + APIs
public/cco.html           Dashboard
public/cco_script...js    Script do dashboard
public/style.css          Estilos
wrangler.toml             Config Cloudflare
```

## Build/Deploy no Cloudflare

```txt
Build command: npm install
Deploy command: npx wrangler deploy
Root directory: /
```

Se o Cloudflare executar apenas Deploy command, o `npx wrangler deploy` baixa o wrangler temporariamente.

## Secrets obrigatorios

Use Secret para:

```txt
BREVO_API_KEY
SUPABASE_SERVICE_ROLE_KEY
CRON_SECRET
```

Use Variable normal para:

```txt
SUPABASE_URL=https://pwjatxqtkvvcmzmjjvbi.supabase.co
REPORT_FROM_EMAIL=jonn3224@gmail.com
REPORT_FROM_NAME=dash de carga
REPORT_DASHBOARD_URL=https://dash-de-carga.hhhthiti.workers.dev/cco.html
REPORT_TIME_ZONE=America/Sao_Paulo
```

Importante: `SUPABASE_SERVICE_ROLE_KEY` nao e o token base64 da Brevo/MCP. A service role real do Supabase costuma comecar com `eyJhbGciOiJIUzI1Ni`.

## Rotas

```txt
/api/send-report
/api/report-config
/api/cron-report-final
```

## Cron

O `wrangler.toml` agenda:

```txt
59 14 * * *
```

Isso equivale a `11:59` em `America/Sao_Paulo`.
