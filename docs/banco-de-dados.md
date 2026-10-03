# Banco de Dados

## Dash de Carga

| Tabela | Uso |
|---|---|
| `reporte_carga` | Grade, campos operacionais, status e referência por DT/data |
| `reporte_materiais` | Materiais e quantidades relacionados à DT |
| `dt_logs`, `reporte_logs` | Histórico e auditoria |
| `oracle_job_status` | Último ciclo da VM e diagnóstico de health |
| `oracle_report_images` | Imagens de reporte |
| `carga_geflux_status` | Snapshot de status operacional do Geflux |
| `import_logs` | Resultado de imports |

## WMS

| Tabela | Uso |
|---|---|
| `wms_mb52_estoque_atual` | Saldo SAP por material/centro/depósito/lote |
| `wms_mb51_movimentacoes` | Histórico SAP de movimentações |
| `wms_demanda_zles002` | Demanda por DT/material |
| `estoque_area` | Controle operacional do WMS; não sobrescrever automaticamente |

Schemas e migrations estão em `supabase/`. Escritas administrativas usam
service role somente na VM ou Worker; o frontend usa chave pública sob RLS.
