# Configurar no Cloudflare Pages

Use **Pages**, nao Worker puro.

Se o projeto estiver como Worker com `Deploy command: npx wrangler deploy`, `/api/send-report` vai dar 404 porque este pacote usa Cloudflare Pages Functions.

## Build

Para HTML/JS estatico:

```txt
Framework preset: None
Build command: deixar vazio
Build output directory: /
Root directory: /
```

Se o painel nao aceitar `/`, use:

```txt
Build command: echo "static deploy"
Build output directory: .
```

## Rotas criadas

Cloudflare Pages Functions usa a pasta `functions`:

```txt
functions/api/send-report.js       -> /api/send-report
functions/api/report-config.js     -> /api/report-config
functions/api/cron-report-final.js -> /api/cron-report-final
```

## Variaveis e secrets

Use **Secret** para:

```txt
BREVO_API_KEY
SUPABASE_SERVICE_ROLE_KEY
CRON_SECRET
```

Use variável normal para:

```txt
REPORT_DASHBOARD_URL
REPORT_FROM_EMAIL
REPORT_FROM_NAME
REPORT_TIME_ZONE
SUPABASE_URL
```

`SUPABASE_SERVICE_ROLE_KEY` precisa ser a service role real do Supabase. Ela normalmente comeca com:

```txt
eyJhbGciOiJIUzI1Ni...
```

Nao use o token base64/MCP da Brevo neste campo.

## Logs / Observability

Pode ativar logs pelo painel Cloudflare em:

```txt
Workers & Pages > seu Pages project > Settings > Observability
```

O JSON sugerido pelo painel e valido para projetos Workers/wrangler, mas nao resolve 404. Primeiro a rota precisa existir via `functions/api/send-report.js`.

## Email Service do Cloudflare

Da para usar, mas ele esta em beta e exige configurar Email Sending/binding. Para este projeto, Brevo continua mais simples e previsivel.
