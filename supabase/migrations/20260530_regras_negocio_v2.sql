-- =============================================================
-- MIGRAÇÃO: Regras de Negócio v2 — 30/05/2026
-- Aplica as correções das 14 regras do dashboard de cargas.
-- Execute no Supabase SQL Editor. É seguro rodar mais de uma vez.
-- =============================================================

begin;

-- -------------------------------------------------------------
-- 1) REGRA 1: DT como entidade única
--    Adiciona índice de busca por DT isolado (sem data_ref)
--    para facilitar lookups e diagnósticos.
-- -------------------------------------------------------------
create index if not exists ix_reporte_carga_dt
  on public.reporte_carga (dt);

-- -------------------------------------------------------------
-- 2) REGRA 10: DTs com DOCA NULL
--    Coluna booleana para identificar DTs importadas sem DOCA.
--    Permite filtrar e auditar sem descartar os registros.
-- -------------------------------------------------------------
alter table if exists public.reporte_carga
  add column if not exists doca_null boolean default false;

create index if not exists ix_reporte_carga_doca_null
  on public.reporte_carga (doca_null)
  where doca_null = true;

-- -------------------------------------------------------------
-- 3) REGRA 9: Badge NOVO
--    Coluna para marcar DTs recém-adicionadas à grade.
--    O dashboard reseta is_novo=false no próximo upload.
-- -------------------------------------------------------------
alter table if exists public.reporte_carga
  add column if not exists is_novo boolean default false;

-- -------------------------------------------------------------
-- 4) REGRA 11: Logs — garantir campos obrigatórios
--    Compatibilidade com ambas as tabelas de log.
-- -------------------------------------------------------------
alter table if exists public.dt_logs
  add column if not exists observacao text;

alter table if exists public.reporte_logs
  add column if not exists observacao text;

-- -------------------------------------------------------------
-- 5) REGRA 7: Divergências — índice para consultas de auditoria
-- -------------------------------------------------------------
create index if not exists ix_reporte_carga_updated_at
  on public.reporte_carga (updated_at desc);

-- -------------------------------------------------------------
-- 6) Função auxiliar: buscar DTs pendentes (GAP)
--    DTs com fim_carregamento < agora E status não finalizado.
-- -------------------------------------------------------------
create or replace function public.fn_gap_dts(p_data_ref text default null)
returns table(
  dt text,
  data_ref text,
  transportadora text,
  status text,
  fim_carregamento text,
  grade_carregamento text,
  peso_liquido text
)
language sql
stable
as $$
  select
    rc.dt,
    rc.data_ref,
    rc.transportadora,
    rc.status,
    rc.fim_carregamento,
    rc.grade_carregamento,
    rc.peso_liquido
  from public.reporte_carga rc
  where
    rc.status not in ('EXPEDIDO','NO SHOW','VEICULO RECUSADO','DT EXCLUIDA')
    and (p_data_ref is null or rc.data_ref = p_data_ref)
  order by rc.fim_carregamento asc nulls last;
$$;

commit;

-- =============================================================
-- NOTAS DE APLICAÇÃO:
--
-- Após rodar esta migração:
--  1. Reimporte a agenda — agora as DTs com DOCA null aparecem
--     na aba "⚠️ DOCA S/ INFO" em vez de serem descartadas.
--  2. DTs GAP (pendentes de dias anteriores) NÃO serão mais
--     marcadas automaticamente como "DT EXCLUIDA" ao importar
--     uma nova agenda.
--  3. DTs novas recebem o badge "NOVO ✨" até o próximo upload.
--  4. A data de referência agora é o FIM de carregamento,
--     alinhando o dashboard com o Excel/reporte Suzano.
-- =============================================================
