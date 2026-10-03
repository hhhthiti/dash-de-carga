# Arquitetura

```text
OTM ─────┐
SAP/ZLES002 ─┼─> Oracle VM: coleta, validação e normalização
Geflux ───┤                         │
MB51/MB52 ─┘                         ├─> Cloudflare Worker -> Supabase -> Dash web
                                      ├─> Excel operacional -> SharePoint
                                      └─> tabelas WMS -> aplicativo WMSS
```

## Componentes

| Componente | Responsabilidade |
|---|---|
| Oracle VM Ubuntu | Perfis Chromium, download, parse, validação, timers e publicação |
| Cloudflare Worker | API protegida da VM, persistência e status para o dashboard |
| Supabase | Dados de carga, logs, status, snapshots e tabelas de integração WMS |
| Dashboard web | Consulta e operação visual de cargas |
| Excel Online | Saída operacional publicada no arquivo oficial `Grade teste.xlsx` |
| WMSS | Módulo separado de estoque, contagem e conferência |

## Princípios

- Grade OTM cria a carga; ZLES002 complementa, não bloqueia sua visualização.
- Geflux é uma fonte de status operacional, não substitui automaticamente toda
  informação manual.
- A VM é coletora central; Excel e dashboard são saídas.
- Toda automação de navegador deve usar perfil e CDP próprios, reutilizando aba.
