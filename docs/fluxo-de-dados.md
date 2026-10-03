# Fluxo de Dados

## Carga

1. OTM exporta a Grade.
2. O watcher valida arquivo, detecta formato e o padroniza como `grade.xls`.
3. A VM normaliza DTs e envia a grade ao Worker.
4. O dashboard mostra DT nova como `PLANEJADO` e `AGUARDANDO_ZLES002` quando
   ainda não há detalhes SAP.
5. A ZLES002 é exportada em sessão SAP autorizada, validada e cruzada por DT.
6. Geflux é lido separadamente; seus campos de portaria, chegada e carregamento
   enriquecem o status operacional sem apagar a agenda.

## Excel operacional

`update_operational_dashboard.py` gera o Excel local. Em seguida,
`repair_excel_for_online.py` produz uma versão compatível com Excel Online e
o publicador atualiza o arquivo oficial no SharePoint quando a sessão/permissão
está disponível.

## WMS

MB52 entra em `wms_mb52_estoque_atual`; MB51 entra em
`wms_mb51_movimentacoes`; demanda ZLES002 entra em `wms_demanda_zles002`.
Essas fontes são comparadas ao WMS, mas não devem atualizar `estoque_area`
automaticamente.
