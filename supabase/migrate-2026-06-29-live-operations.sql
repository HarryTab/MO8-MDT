alter table public.operational_messages
  add column if not exists incident_id text references public.operational_incidents(incident_id) on delete set null,
  add column if not exists channel text not null default 'Main',
  add column if not exists message_type text not null default 'Transmission',
  add column if not exists status text not null default 'Sent',
  add column if not exists acknowledged_by text[] not null default array[]::text[];

create index if not exists idx_operational_messages_channel
  on public.operational_messages(channel, created_at desc);
create index if not exists idx_operational_messages_incident
  on public.operational_messages(incident_id, created_at desc);

create table if not exists public.operational_entity_relationships (
  relationship_id text primary key default ('REL_' || replace(gen_random_uuid()::text, '-', '')),
  source_type text not null check (source_type in ('Person', 'Vehicle', 'Incident', 'Operation', 'Location')),
  source_id text not null,
  target_type text not null check (target_type in ('Person', 'Vehicle', 'Incident', 'Operation', 'Location')),
  target_id text not null,
  relationship_type text not null default 'Associated with',
  confidence text not null default 'Confirmed' check (confidence in ('Confirmed', 'Probable', 'Possible')),
  notes text not null default '',
  valid_from timestamptz,
  valid_to timestamptz,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint operational_relationship_not_self check (source_type <> target_type or source_id <> target_id)
);

create unique index if not exists idx_operational_relationship_unique
  on public.operational_entity_relationships(source_type, source_id, target_type, target_id, relationship_type);
create index if not exists idx_operational_relationship_source
  on public.operational_entity_relationships(source_type, source_id);
create index if not exists idx_operational_relationship_target
  on public.operational_entity_relationships(target_type, target_id);

create table if not exists public.shift_wrapups (
  wrapup_id text primary key default ('WRAP_' || replace(gen_random_uuid()::text, '-', '')),
  shift_id text not null unique references public.shift_logs(shift_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  summary text not null default '',
  carried_forward_notes text not null default '',
  incident_ids text[] not null default array[]::text[],
  completed_report_count integer not null default 0,
  incomplete_report_count integer not null default 0,
  arrest_count integer not null default 0,
  fpn_count integer not null default 0,
  warning_count integer not null default 0,
  power_use_count integer not null default 0,
  snapshot jsonb not null default '{}'::jsonb,
  submitted_by text references public.profiles(user_id) on delete set null,
  submitted_at timestamptz not null default now()
);

create index if not exists idx_shift_wrapups_officer
  on public.shift_wrapups(officer_id, submitted_at desc);

alter table public.operational_entity_relationships enable row level security;
alter table public.shift_wrapups enable row level security;

drop policy if exists "operational relationships read authenticated" on public.operational_entity_relationships;
create policy "operational relationships read authenticated" on public.operational_entity_relationships
for select using (auth.role() = 'authenticated');

drop policy if exists "operational relationships write authenticated" on public.operational_entity_relationships;
create policy "operational relationships write authenticated" on public.operational_entity_relationships
for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "shift wrapups read own or supervisors" on public.shift_wrapups;
create policy "shift wrapups read own or supervisors" on public.shift_wrapups
for select using (
  public.has_permission('VIEW_TASKS')
  or exists (
    select 1 from public.officers officer
    where officer.officer_id = shift_wrapups.officer_id
      and officer.member_id = public.current_member_id()
  )
);

drop policy if exists "shift wrapups create own" on public.shift_wrapups;
create policy "shift wrapups create own" on public.shift_wrapups
for insert with check (
  exists (
    select 1 from public.officers officer
    where officer.officer_id = shift_wrapups.officer_id
      and officer.member_id = public.current_member_id()
  )
);

alter table public.operational_messages replica identity full;
alter table public.operational_incidents replica identity full;
alter table public.operational_units replica identity full;

do $$
declare
  table_name text;
begin
  foreach table_name in array array[
    'operational_messages',
    'operational_incidents',
    'operational_incident_logs',
    'operational_units',
    'operational_unit_members',
    'operational_entity_relationships'
  ] loop
    if not exists (
      select 1
      from pg_publication_tables
      where pubname = 'supabase_realtime'
        and schemaname = 'public'
        and tablename = table_name
    ) then
      execute format('alter publication supabase_realtime add table public.%I', table_name);
    end if;
  end loop;
end $$;
