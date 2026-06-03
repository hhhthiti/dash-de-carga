create table if not exists public.report_delivery_config (
  id text primary key default 'default',
  email_provider text not null default 'brevo',
  email_from text,
  email_to text[] not null default '{}',
  email_cc text[] not null default '{}',
  auto_email_enabled boolean not null default true,
  auto_email_interval_minutes integer not null default 60,
  report_inicial text not null default '00:00',
  report_final text not null default '11:59',
  send_next_day_at_final boolean not null default true,
  updated_at timestamptz not null default now()
);

insert into public.report_delivery_config (
  id,
  email_provider,
  email_from,
  auto_email_enabled,
  auto_email_interval_minutes,
  report_inicial,
  report_final,
  send_next_day_at_final
) values (
  'default',
  'brevo',
  'jonn3224@gmail.com',
  true,
  60,
  '00:00',
  '11:59',
  true
) on conflict (id) do nothing;

create table if not exists public.report_delivery_runs (
  slot_key text primary key,
  sent_at timestamptz not null default now(),
  details jsonb not null default '{}'::jsonb
);

create index if not exists ix_report_delivery_runs_sent_at
  on public.report_delivery_runs (sent_at desc);
