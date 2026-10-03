# Arquivo do Projeto — Dash de Carga

Este documento congela o desenho conhecido do projeto sem desligá-lo ou apagar
arquivos. Ele serve para que o sistema possa ser entendido e reconstruído mesmo
se a VM, uma sessão corporativa ou o link remoto deixarem de funcionar.

## Objetivo

O Dash de Carga foi criado para consolidar a agenda logística, o andamento
operacional e os dados de expedição em um painel web e em uma planilha online.
O sistema usa dados de OTM, SAP/ZLES002 e Geflux; MB51 e MB52 também alimentam
o módulo WMS/estoque.

## O que o sistema resolve

- Mostra DTs e horários da Grade OTM mesmo antes da ZLES002 chegar.
- Enriquece DTs com materiais, peso, remessa e cliente quando a ZLES002 existe.
- Mantém uma camada separada de status Geflux para refletir o andamento real.
- Publica um Excel operacional compatível com Excel Online/SharePoint.
- Registra status da automação para o dashboard e para diagnóstico.
- Usa MB52 e MB51 como referência SAP para o WMS, sem apagar estoque manual.

## Estado de congelamento

O código e os artefatos locais estão preservados; isso não significa que todas
as dependências externas estejam disponíveis. OTM, Citrix/SAP, Geflux, WhatsApp,
SharePoint e links TryCloudflare podem expirar, reiniciar ou requerer sessão
autorizada.

## Mapa da documentação

- [Arquitetura](arquitetura.md)
- [Fluxo de dados](fluxo-de-dados.md)
- [Banco de dados](banco-de-dados.md)
- [Regras de negócio](regras-de-negocio.md)
- [Infraestrutura e recuperação](infraestrutura.md)
- [Histórico e aprendizados](historico.md)

## Como retomar com segurança

1. Leia `docs/infraestrutura.md` e confirme acesso SSH à VM.
2. Confirme serviços, memória, espaço, CDP e últimos logs antes de rodar ciclos.
3. Verifique sessões autorizadas nos perfis Chromium separados.
4. Rode um ciclo manual ou `--dry-run` antes de reativar timers.
5. Só publique Excel se a geração e a versão online-safe forem válidas.
