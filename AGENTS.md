# Contexto do projeto

Este repositório é a base do **Dash de Carga** e da sua integração operacional
com a VM Oracle. Não trate o dashboard, a VM e o WMS como projetos isolados.

## Limites obrigatórios

- Nunca incluir `.env`, chaves Supabase, secrets do Worker, senhas ou cookies no Git.
- Nunca colocar `service_role` no frontend.
- SAP/Citrix, OTM, Geflux, Excel Online e WhatsApp dependem de sessões autorizadas;
  não automatizar digitação de senha, MFA ou aprovação de autenticação.
- Não sobrescrever `estoque_area` do WMS com dados MB52. MB52 é saldo SAP para
  consulta e conferência.
- Antes de publicar no SharePoint, gerar e validar o XLSX online-safe.
- Antes de mudar um scheduler, conferir logs, lock e abas Chromium/CDP existentes;
  automações devem reutilizar uma aba do serviço, não criar abas em loop.

## Referências rápidas

- Dashboard web: `public/cco.html` e `public/cco_script_fim_agenda.js`
- Worker: `src/index.js`
- Automação na VM: `automacoes/oracle/`
- Schema/migrations: `supabase/`
- Documentação de congelamento: `docs/`

Leia `docs/ARQUIVO_DO_PROJETO.md` antes de alterar arquitetura ou infraestrutura.
