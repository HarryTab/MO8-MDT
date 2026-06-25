alter table public.operational_incidents
  add column if not exists assigned_unit_ids text[] not null default array[]::text[],
  add column if not exists required_capabilities text[] not null default array[]::text[],
  add column if not exists patrol_area text not null default '',
  add column if not exists radio_channel text not null default '';

create table if not exists public.operational_units (
  unit_id text primary key default ('UNIT_' || replace(gen_random_uuid()::text, '-', '')),
  callsign text not null,
  status text not null default 'Available' check (status in ('Available', 'Assigned', 'En Route', 'On Scene', 'Transporting', 'At Station', 'Out of Service')),
  patrol_area text not null default '',
  patrol_type text not null default 'Roads Policing',
  radio_channel text not null default '',
  lead_officer_id text references public.officers(officer_id) on delete set null,
  current_incident_id text references public.operational_incidents(incident_id) on delete set null,
  created_by text references public.profiles(user_id) on delete set null,
  notes text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  closed_at timestamptz
);

create table if not exists public.operational_unit_members (
  membership_id text primary key default ('UMEM_' || replace(gen_random_uuid()::text, '-', '')),
  unit_id text not null references public.operational_units(unit_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  role text not null default 'Crew' check (role in ('Lead', 'Driver', 'Operator', 'Crew', 'Observer', 'Supervisor')),
  status text not null default 'Active' check (status in ('Active', 'Left')),
  joined_at timestamptz not null default now(),
  left_at timestamptz,
  unique (unit_id, officer_id)
);

create table if not exists public.operational_unit_join_requests (
  request_id text primary key default ('UJ_' || replace(gen_random_uuid()::text, '-', '')),
  unit_id text not null references public.operational_units(unit_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  status text not null default 'Pending' check (status in ('Pending', 'Approved', 'Denied', 'Cancelled')),
  message text not null default '',
  review_reason text not null default '',
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  created_at timestamptz not null default now()
);

create table if not exists public.operational_incident_logs (
  log_id text primary key default ('ILOG_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  unit_id text references public.operational_units(unit_id) on delete set null,
  author_user_id text references public.profiles(user_id) on delete set null,
  entry_type text not null default 'Note',
  body text not null default '',
  created_at timestamptz not null default now()
);

create table if not exists public.operational_incident_links (
  link_id text primary key default ('ILNK_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  title text not null default '',
  url text not null,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.operational_briefings (
  briefing_id text primary key default ('BRF_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  objectives text not null default '',
  patrol_area text not null default '',
  radio_channel text not null default '',
  commander_user_id text references public.profiles(user_id) on delete set null,
  assigned_unit_ids text[] not null default array[]::text[],
  starts_at timestamptz,
  ends_at timestamptz,
  status text not null default 'Active' check (status in ('Draft', 'Active', 'Completed', 'Cancelled')),
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.operational_messages (
  message_id text primary key default ('OMSG_' || replace(gen_random_uuid()::text, '-', '')),
  target_unit_id text references public.operational_units(unit_id) on delete cascade,
  target_officer_id text references public.officers(officer_id) on delete cascade,
  sender_user_id text references public.profiles(user_id) on delete set null,
  priority text not null default 'Normal' check (priority in ('Urgent', 'High', 'Normal', 'Low')),
  message text not null,
  created_at timestamptz not null default now()
);

create index if not exists idx_operational_units_status on public.operational_units(status, created_at desc);
create index if not exists idx_operational_unit_members_unit on public.operational_unit_members(unit_id, status);
create index if not exists idx_operational_join_requests_unit on public.operational_unit_join_requests(unit_id, status);
create index if not exists idx_operational_logs_incident on public.operational_incident_logs(incident_id, created_at desc);
create index if not exists idx_operational_briefings_status on public.operational_briefings(status, starts_at);

alter table public.operational_units enable row level security;
alter table public.operational_unit_members enable row level security;
alter table public.operational_unit_join_requests enable row level security;
alter table public.operational_incident_logs enable row level security;
alter table public.operational_incident_links enable row level security;
alter table public.operational_briefings enable row level security;
alter table public.operational_messages enable row level security;

drop policy if exists "operational units read authenticated" on public.operational_units;
create policy "operational units read authenticated" on public.operational_units for select using (auth.role() = 'authenticated');
drop policy if exists "operational units write authenticated" on public.operational_units;
create policy "operational units write authenticated" on public.operational_units for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "operational unit members read authenticated" on public.operational_unit_members;
create policy "operational unit members read authenticated" on public.operational_unit_members for select using (auth.role() = 'authenticated');
drop policy if exists "operational unit members write authenticated" on public.operational_unit_members;
create policy "operational unit members write authenticated" on public.operational_unit_members for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "operational join requests read authenticated" on public.operational_unit_join_requests;
create policy "operational join requests read authenticated" on public.operational_unit_join_requests for select using (auth.role() = 'authenticated');
drop policy if exists "operational join requests write authenticated" on public.operational_unit_join_requests;
create policy "operational join requests write authenticated" on public.operational_unit_join_requests for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "operational logs read authenticated" on public.operational_incident_logs;
create policy "operational logs read authenticated" on public.operational_incident_logs for select using (auth.role() = 'authenticated');
drop policy if exists "operational logs write authenticated" on public.operational_incident_logs;
create policy "operational logs write authenticated" on public.operational_incident_logs for insert with check (auth.role() = 'authenticated');

drop policy if exists "operational links read authenticated" on public.operational_incident_links;
create policy "operational links read authenticated" on public.operational_incident_links for select using (auth.role() = 'authenticated');
drop policy if exists "operational links write authenticated" on public.operational_incident_links;
create policy "operational links write authenticated" on public.operational_incident_links for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "operational briefings read authenticated" on public.operational_briefings;
create policy "operational briefings read authenticated" on public.operational_briefings for select using (auth.role() = 'authenticated');
drop policy if exists "operational briefings write supervisors" on public.operational_briefings;
create policy "operational briefings write supervisors" on public.operational_briefings for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "operational messages read authenticated" on public.operational_messages;
create policy "operational messages read authenticated" on public.operational_messages for select using (auth.role() = 'authenticated');
drop policy if exists "operational messages write authenticated" on public.operational_messages;
create policy "operational messages write authenticated" on public.operational_messages for insert with check (auth.role() = 'authenticated');
