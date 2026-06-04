# Configurar envio Brevo na Vercel

Este projeto envia reporte direto pela Vercel, sem `.bat` e sem expor a chave da Brevo no navegador.

## 1. Variaveis de ambiente da Vercel

No projeto `dash-de-carga` na Vercel, configure em `Settings > Environment Variables`:

```env
BREVO_API_KEY=sua_chave_brevo
REPORT_FROM_EMAIL=jonn3224@gmail.com
REPORT_FROM_NAME=dash de carga
REPORT_DASHBOARD_URL=https://dash-de-carga.vercel.app/cco.html
SUPABASE_URL=https://pwjatxqtkvvcmzmjjvbi.supabase.co
SUPABASE_SERVICE_ROLE_KEY=sua_service_role_key
CRON_SECRET=uma_senha_grande_aleatoria
REPORT_TIME_ZONE=America/Sao_Paulo
```

O `REPORT_FROM_EMAIL` precisa estar validado como remetente dentro da Brevo.

## 2. Banco Supabase

Rode a migration:

```txt
supabase/migrations/20260603_report_delivery_config.sql
```

Ela cria:

- `report_delivery_config`: guarda destinatarios, horario e automatico
- `report_delivery_runs`: evita disparo duplicado

## 3. Como o dashboard funciona

No dashboard, abra `Configurações`:

- Serviço de e-mail: `Brevo`
- Remetente: `jonn3224@gmail.com`
- URL da automação: pode deixar vazio na Vercel
- Destinatários/CC: escolha quem recebe
- E-mail automático: `Ligado`
- Intervalo: `60`
- Reporte inicial: `00:00`
- Reporte final: `11:59`

Ao salvar, o dashboard grava essa configuração em `/api/report-config`.

## 4. Rotas criadas

- `POST /api/send-report`: envio manual pelo botão
- `GET /api/report-config`: carrega configuração salva
- `POST /api/report-config`: salva destinatarios/horarios
- `GET /api/cron-report`: envio automatico de hora em hora
- `GET /api/cron-report-final`: disparo especial das 11:59

## 5. Agendamento

No plano Hobby, a Vercel nao aceita cron de hora em hora. Por isso o agendamento desta versao fica no GitHub Actions, em:

```txt
.github/workflows/send-report-cron.yml
```

Configure no GitHub em `Settings > Secrets and variables > Actions`:

- Secret `CRON_SECRET`: o mesmo valor configurado na Vercel
- Variable `DASHBOARD_BASE_URL`: `https://dash-de-carga.vercel.app`

O workflow roda:

- `0 * * * *`: chama `https://dash-de-carga.vercel.app/api/cron-report`
- `59 14 * * *`: chama `https://dash-de-carga.vercel.app/api/cron-report-final`

O horario do GitHub Actions tambem usa UTC. `59 14 * * *` equivale a `11:59` em `America/Sao_Paulo`.

No disparo de `11:59`, a API envia:

- reporte do dia atual
- reporte do dia seguinte

## 6. WhatsApp

WhatsApp ficou fora desta versão.

Automação por WhatsApp Web precisa de um servidor sempre ligado mantendo sessão de navegador. A Vercel não é ideal para isso, porque Functions são temporarias e não mantêm sessão aberta.

## Segurança

Não coloque `BREVO_API_KEY` nem `SUPABASE_SERVICE_ROLE_KEY` no HTML/JS do dashboard.

Como a chave Brevo e a service role foram compartilhadas em texto puro durante a configuração, o ideal é gerar novas chaves depois de validar que tudo está funcionando.
