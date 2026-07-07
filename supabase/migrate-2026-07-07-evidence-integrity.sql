alter table public.operational_incident_attachments
  add column if not exists sha256 text not null default '',
  add column if not exists integrity_status text not null default 'Unverified'
    check (integrity_status in ('Unverified', 'Verified', 'Mismatch')),
  add column if not exists verified_at timestamptz,
  add column if not exists verified_by text references public.profiles(user_id) on delete set null;

create table if not exists public.operational_evidence_access_log (
  access_id text primary key default ('EACC_' || replace(gen_random_uuid()::text, '-', '')),
  attachment_id text not null,
  incident_id text not null,
  actor_user_id text references public.profiles(user_id) on delete set null,
  action text not null check (action in ('Opened', 'Verified', 'Integrity Mismatch', 'Downloaded')),
  expected_hash text not null default '',
  observed_hash text not null default '',
  details jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists idx_evidence_access_attachment
  on public.operational_evidence_access_log(attachment_id, created_at desc);
create index if not exists idx_evidence_access_incident
  on public.operational_evidence_access_log(incident_id, created_at desc);

alter table public.operational_evidence_access_log enable row level security;

drop policy if exists "evidence access read authenticated" on public.operational_evidence_access_log;
create policy "evidence access read authenticated" on public.operational_evidence_access_log
for select using (auth.role() = 'authenticated');

drop policy if exists "evidence access insert authenticated" on public.operational_evidence_access_log;
create policy "evidence access insert authenticated" on public.operational_evidence_access_log
for insert with check (actor_user_id = public.current_user_id());

drop policy if exists "evidence integrity update authenticated" on public.operational_incident_attachments;
create policy "evidence integrity update authenticated" on public.operational_incident_attachments
for update using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
