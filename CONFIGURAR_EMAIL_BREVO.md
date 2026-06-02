# Configurar envio de reporte pela Brevo

Este projeto agora tem um servidor local para enviar o e-mail sem expor a chave Brevo no HTML.

## 1. Criar o arquivo local de ambiente

Crie um arquivo chamado `.env.local` nesta pasta, usando `.env.local.example` como modelo.

Campos principais:

```env
BREVO_API_KEY=sua_chave_brevo
BREVO_EVENT_FALLBACK_EMAIL=contato@empresa.com
REPORT_FROM_EMAIL=email_verificado_na_brevo@suaempresa.com
REPORT_FROM_NAME=Reporte Operacional
REPORT_DEFAULT_TO=destinatario@suaempresa.com
REPORT_DEFAULT_CC=
REPORT_SERVER_PORT=8787
REPORT_PUBLIC_BASE_URL=https://seu-dashboard.vercel.app
REPORT_IMAGE_TITLE=Dashboard de Carga
```

O `REPORT_FROM_EMAIL` precisa ser um remetente validado na Brevo.

## 2. Iniciar o servidor

Abra `iniciar-servidor-email.bat`.

Ele sobe:

```txt
http://localhost:8787
```

Endpoint usado pelo dashboard:

```txt
http://localhost:8787/send-report
```

## 3. Configurar no dashboard

Em Configurações do reporte:

- Serviço de e-mail: `Brevo`
- Remetente: o mesmo e-mail validado na Brevo
- URL da automação: `http://localhost:8787/send-report`
- Destinatários/CC: conforme necessário

Se o serviço for `Brevo` e a URL ficar vazia, o dashboard usa automaticamente `http://localhost:8787/send-report`.

## 4. Eventos customizados

O dashboard agora consegue registrar eventos no Brevo de dois jeitos:

- tracker no navegador, usando o `client_key` do Brevo
- endpoint do servidor em `POST /brevo-event`, usando a `BREVO_API_KEY`

No painel de configurações do dashboard, preencha:

- `BREVO CLIENT KEY` se você quiser usar o tracker do navegador
- `E-MAIL DO CONTATO BREVO` para associar o evento a um contato
- `URL DO EVENTO BREVO` para apontar para seu servidor local ou produção

Eventos enviados pelo app:

- `dashboard_opened`
- `snapshot_saved`
- `planned_suzano_updated`
- `report_sent`

## 5. Imagem do relatório

O servidor agora expõe uma rota pública para imagem:

```txt
http://localhost:8787/api/relatorio?ts=TIMESTAMP
```

Em produção, use:

```txt
https://seu-dashboard.vercel.app/api/relatorio?ts=TIMESTAMP
```

O parâmetro `ts` ajuda a quebrar cache em Gmail, Brevo e outros clientes.

## Segurança

Não coloque a chave Brevo dentro do `cco.html` nem no `cco_script_fim_agenda.js`.

Como a chave foi compartilhada em texto puro durante a configuração, o ideal é gerar uma nova chave na Brevo e desativar a antiga.
