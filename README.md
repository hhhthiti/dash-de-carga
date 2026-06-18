# Dashboard de Carga

Base preservada da branch `dash-resend`. O dashboard continua concentrado em
`public/cco.html` e `public/cco_script_fim_agenda.js`; a automacao apenas alimenta
as tabelas que ele ja utiliza.

## Responsabilidades

### Cloudflare Pages/Assets

- `public/index.html`
- `public/cco.html`
- `public/style.css`
- `public/cco_script_fim_agenda.js`

Exibe tabela, lideres, planejamento, materiais, mapa de separacao, logs e
reporte. Os uploads de grade e ZLES002 continuam disponiveis como fallback.

### Cloudflare Worker

- `src/index.js`
- recebe dados normalizados da Oracle VM;
- preserva campos operacionais ao atualizar a grade;
- grava `reporte_carga`, `reporte_materiais` e `dt_logs`;
- expoe o status da ultima execucao;
- mantem as rotas atuais de configuracao e e-mail Zoho.

Rotas da VM:

```text
POST /api/oracle/sync-grade
POST /api/oracle/sync-zles002
POST /api/oracle/job-status
POST /api/oracle/report-image
GET  /api/oracle/last-status
```

As rotas POST exigem `X-Oracle-Job-Secret`. A rota GET retorna somente o resumo
operacional usado pelo badge do dashboard.

### Oracle VM

Os scripts ficam em `automacoes/oracle/`. Consulte
`automacoes/oracle/README.md` para teste local e systemd.

### Supabase

O dashboard atual usa:

- `reporte_carga`: grade e estado operacional por `dt + data_ref`;
- `reporte_materiais`: materiais/quantidades da DT;
- `dt_logs` e `reporte_logs`: historico;
- tabelas de snapshots e configuracao de reporte.

Execute primeiro o schema existente e depois a migration:

```text
supabase/schema-clean.sql
supabase/migrations/20260617_oracle_automation.sql
```

A migration adiciona apenas colunas ja esperadas pelo frontend e duas estruturas
isoladas para status da automacao e imagem do reporte.

## Variaveis do Worker

Secrets:

```text
SUPABASE_SERVICE_ROLE_KEY
ORACLE_JOB_SECRET
CRON_SECRET
ADMIN_SECRET
ZOHO_SMTP_USER
ZOHO_SMTP_PASS
```

Variables:

```text
SUPABASE_URL=https://pwjatxqtkvwcmzmjjvbi.supabase.co
REPORT_TIME_ZONE=America/Sao_Paulo
REPORT_FROM_EMAIL=...
REPORT_FROM_NAME=...
REPORT_DASHBOARD_URL=...
ZOHO_SMTP_HOST=smtp.zoho.com
ZOHO_SMTP_PORT=465
```

## Desenvolvimento e deploy

```bash
npm install
cp .dev.vars.example .dev.vars
npx wrangler dev
npx wrangler deploy
```

Os crons do Worker permanecem em `wrangler.toml` para os reportes de 14h, 22h e
06h. A coleta de grade/ZLES002 e executada na VM a cada 20 minutos.
