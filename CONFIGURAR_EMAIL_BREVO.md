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
- `GET /api/cron-report-final`: disparo diario das 11:59

## 5. Agendamento

No plano Hobby, a Vercel aceita cron nativo somente uma vez por dia. Esta versao usa apenas o fechamento diario:

```json
{
  "crons": [
    {
      "path": "/api/cron-report-final",
      "schedule": "59 14 * * *"
    }
  ]
}
```

O horario do cron da Vercel usa UTC. `59 14 * * *` equivale a `11:59` em `America/Sao_Paulo`.

Observacao: no plano Hobby, a Vercel informa precisao de agendamento em janela de ate uma hora. Entao o disparo pode ocorrer pouco depois de 11:59. Para horario exato, use Vercel Pro ou GitHub Actions.

No disparo de `11:59`, a API envia:

- reporte do dia atual
- reporte do dia seguinte

## 6. WhatsApp

WhatsApp ficou fora desta versão.

Automação por WhatsApp Web precisa de um servidor sempre ligado mantendo sessão de navegador. A Vercel não é ideal para isso, porque Functions são temporarias e não mantêm sessão aberta.

## Segurança

Não coloque `BREVO_API_KEY` nem `SUPABASE_SERVICE_ROLE_KEY` no HTML/JS do dashboard.

Como a chave Brevo e a service role foram compartilhadas em texto puro durante a configuração, o ideal é gerar novas chaves depois de validar que tudo está funcionando.
