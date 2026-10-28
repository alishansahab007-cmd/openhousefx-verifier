-- Optional metadata-only logging for the OpenhouseFX verifier.
-- Execute once in the Supabase SQL editor as the project database administrator.
-- The app uses a server-side legacy service_role JWT key; no public policies.
create table if not exists public.verifier_runs (
    batch_id uuid primary key,
    total_validated bigint not null check (total_validated >= 0),
    counts jsonb not null,
    created_at timestamptz not null default now()
);

alter table public.verifier_runs enable row level security;
revoke all on table public.verifier_runs from public, anon, authenticated;
grant select, insert, update on table public.verifier_runs to service_role;

comment on table public.verifier_runs is
  'OpenhouseFX run totals only. No lead rows or filenames; backend service_role access.';
