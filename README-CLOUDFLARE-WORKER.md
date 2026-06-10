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
SUPABASE_SERVICE_ROLE_KEY
CRON_SECRET
ZOHO_SMTP_USER
ZOHO_SMTP_PASS
```

Use Variable normal para:

```txt
SUPABASE_URL=https://pwjatxqtkvwcmzmjjvbi.supabase.co
REPORT_FROM_EMAIL=relatorio@zohomail.com
REPORT_FROM_NAME=dash de carga
REPORT_DASHBOARD_URL=https://dash-de-carga-cloud.hhhthiti.workers.dev/cco.html
REPORT_TIME_ZONE=America/Sao_Paulo
ZOHO_SMTP_HOST=smtp.zoho.com
ZOHO_SMTP_PORT=465
```

Importante: `SUPABASE_SERVICE_ROLE_KEY` nao e token de provedor de email. A service role real do Supabase costuma comecar com `eyJhbGciOiJIUzI1Ni`.

No Zoho, gere e use uma senha de aplicativo em `Security > App Passwords`. Mantenha `REPORT_FROM_EMAIL` igual ao usuario autenticado no SMTP para evitar `relaying disallowed`.

## Rotas

```txt
/api/send-report
/api/report-config
/api/cron-report-final
```

## Cron

O `wrangler.toml` agenda:

```txt
0 17 * * *  -> 14h em America/Sao_Paulo, fechamento T1
0 1 * * *   -> 22h em America/Sao_Paulo, fechamento T2
0 9 * * *   -> 06h em America/Sao_Paulo, fechamento T3
```

Cada envio leva o resumo visual no corpo do e-mail e um CSV anexo com as cargas do reporte sem materiais.
