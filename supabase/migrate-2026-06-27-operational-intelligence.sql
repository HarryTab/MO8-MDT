alter table public.operational_units
  add column if not exists callsign_number text not null default '',
  add column if not exists assigned_vehicle_registration text not null default '';

alter table public.operational_incidents
  add column if not exists incident_data jsonb not null default '{}'::jsonb;

create table if not exists public.operational_persons (
  person_id text primary key default ('PERS_' || replace(gen_random_uuid()::text, '-', '')),
  display_name text not null,
  roblox_username text not null default '',
  aliases text[] not null default array[]::text[],
  notes text not null default '',
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.operational_vehicles (
  vehicle_id text primary key default ('VEH_' || replace(gen_random_uuid()::text, '-', '')),
  registration text not null unique,
  make text not null default '',
  model text not null default '',
  colour text not null default '',
  owner_person_id text references public.operational_persons(person_id) on delete set null,
  status text not null default 'Active' check (status in ('Active', 'Stolen', 'Seized', 'Destroyed', 'Archived')),
  notes text not null default '',
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.operational_offences (
  offence_id text primary key default ('OFF_' || replace(gen_random_uuid()::text, '-', '')),
  code text not null,
  title text not null,
  category text not null default 'General',
  description text not null default '',
  default_fine integer not null default 0,
  default_prison_minutes integer not null default 0,
  default_points integer not null default 0,
  version integer not null default 1,
  effective_from date not null default current_date,
  active boolean not null default true,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (code, version)
);

create table if not exists public.operational_incident_entities (
  involvement_id text primary key default ('INV_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  person_id text references public.operational_persons(person_id) on delete cascade,
  vehicle_id text references public.operational_vehicles(vehicle_id) on delete cascade,
  involvement_role text not null default 'Subject',
  notes text not null default '',
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  check (person_id is not null or vehicle_id is not null)
);

create table if not exists public.operational_disposals (
  disposal_id text primary key default ('DISP_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  person_id text references public.operational_persons(person_id) on delete set null,
  vehicle_id text references public.operational_vehicles(vehicle_id) on delete set null,
  offence_id text references public.operational_offences(offence_id) on delete set null,
  outcome_type text not null default 'No Further Action',
  fine_amount integer not null default 0,
  prison_minutes integer not null default 0,
  penalty_points integer not null default 0,
  notes text not null default '',
  issued_by text references public.profiles(user_id) on delete set null,
  issued_at timestamptz not null default now()
);

create table if not exists public.operational_intel_markers (
  marker_id text primary key default ('MARK_' || replace(gen_random_uuid()::text, '-', '')),
  person_id text references public.operational_persons(person_id) on delete cascade,
  vehicle_id text references public.operational_vehicles(vehicle_id) on delete cascade,
  marker_type text not null,
  severity text not null default 'Information' check (severity in ('Critical', 'High', 'Caution', 'Information')),
  details text not null default '',
  source_incident_id text references public.operational_incidents(incident_id) on delete set null,
  review_at date,
  expires_at date,
  status text not null default 'Active' check (status in ('Active', 'Expired', 'Removed')),
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  check (person_id is not null or vehicle_id is not null)
);

create table if not exists public.operational_operations (
  operation_id text primary key default ('OP_' || replace(gen_random_uuid()::text, '-', '')),
  reference text not null unique default ('OP-' || to_char(now(), 'YYYYMMDD') || '-' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 4))),
  name text not null,
  status text not null default 'Active' check (status in ('Planned', 'Active', 'Paused', 'Completed', 'Cancelled')),
  classification text not null default 'Routine',
  objectives text not null default '',
  lead_user_id text references public.profiles(user_id) on delete set null,
  starts_at timestamptz,
  ends_at timestamptz,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.operational_operation_links (
  operation_link_id text primary key default ('OPL_' || replace(gen_random_uuid()::text, '-', '')),
  operation_id text not null references public.operational_operations(operation_id) on delete cascade,
  subject_type text not null check (subject_type in ('Incident', 'Person', 'Vehicle', 'BOLO')),
  subject_id text not null,
  relationship text not null default 'Linked',
  notes text not null default '',
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  unique (operation_id, subject_type, subject_id)
);

create table if not exists public.operational_unit_invitations (
  invitation_id text primary key default ('UINV_' || replace(gen_random_uuid()::text, '-', '')),
  unit_id text not null references public.operational_units(unit_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  role text not null default 'Crew' check (role in ('Driver', 'Operator', 'Crew', 'Observer', 'Supervisor')),
  message text not null default '',
  status text not null default 'Pending' check (status in ('Pending', 'Accepted', 'Declined', 'Cancelled')),
  invited_by text references public.profiles(user_id) on delete set null,
  responded_at timestamptz,
  created_at timestamptz not null default now()
);

create index if not exists idx_operational_persons_name on public.operational_persons(lower(display_name));
create index if not exists idx_operational_vehicles_registration on public.operational_vehicles(lower(registration));
create index if not exists idx_operational_offences_code on public.operational_offences(lower(code), active);
create index if not exists idx_operational_entities_incident on public.operational_incident_entities(incident_id);
create index if not exists idx_operational_entities_person on public.operational_incident_entities(person_id);
create index if not exists idx_operational_entities_vehicle on public.operational_incident_entities(vehicle_id);
create index if not exists idx_operational_disposals_incident on public.operational_disposals(incident_id, issued_at desc);
create index if not exists idx_operational_markers_person on public.operational_intel_markers(person_id, status);
create index if not exists idx_operational_markers_vehicle on public.operational_intel_markers(vehicle_id, status);
create index if not exists idx_operational_operation_links_operation on public.operational_operation_links(operation_id, subject_type);
create index if not exists idx_operational_unit_invitations_officer on public.operational_unit_invitations(officer_id, status);

alter table public.operational_persons enable row level security;
alter table public.operational_vehicles enable row level security;
alter table public.operational_offences enable row level security;
alter table public.operational_incident_entities enable row level security;
alter table public.operational_disposals enable row level security;
alter table public.operational_intel_markers enable row level security;
alter table public.operational_operations enable row level security;
alter table public.operational_operation_links enable row level security;
alter table public.operational_unit_invitations enable row level security;

drop policy if exists "operational people authenticated" on public.operational_persons;
create policy "operational people authenticated" on public.operational_persons for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational vehicles authenticated" on public.operational_vehicles;
create policy "operational vehicles authenticated" on public.operational_vehicles for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational entities authenticated" on public.operational_incident_entities;
create policy "operational entities authenticated" on public.operational_incident_entities for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational disposals authenticated" on public.operational_disposals;
create policy "operational disposals authenticated" on public.operational_disposals for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational markers read authenticated" on public.operational_intel_markers;
create policy "operational markers read authenticated" on public.operational_intel_markers for select using (auth.role() = 'authenticated');
drop policy if exists "operational markers manage supervisors" on public.operational_intel_markers;
create policy "operational markers manage supervisors" on public.operational_intel_markers for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));
drop policy if exists "operational offences read authenticated" on public.operational_offences;
create policy "operational offences read authenticated" on public.operational_offences for select using (auth.role() = 'authenticated');
drop policy if exists "operational offences manage supervisors" on public.operational_offences;
create policy "operational offences manage supervisors" on public.operational_offences for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));
drop policy if exists "operational operations read authenticated" on public.operational_operations;
create policy "operational operations read authenticated" on public.operational_operations for select using (auth.role() = 'authenticated');
drop policy if exists "operational operations manage supervisors" on public.operational_operations;
create policy "operational operations manage supervisors" on public.operational_operations for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));
drop policy if exists "operational operation links authenticated" on public.operational_operation_links;
create policy "operational operation links authenticated" on public.operational_operation_links for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational unit invitations authenticated" on public.operational_unit_invitations;
create policy "operational unit invitations authenticated" on public.operational_unit_invitations for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "operational incident attach own unit" on public.operational_incidents;
create policy "operational incident attach own unit" on public.operational_incidents for update
using (auth.role() = 'authenticated')
with check (
  public.has_permission('VIEW_TASKS')
  or created_by = public.current_user_id()
  or exists (
    select 1
    from public.operational_unit_members um
    join public.officers o on o.officer_id = um.officer_id
    join public.profiles p on p.member_id = o.member_id
    where p.user_id = public.current_user_id()
      and um.status = 'Active'
      and um.unit_id = any(assigned_unit_ids)
  )
);
