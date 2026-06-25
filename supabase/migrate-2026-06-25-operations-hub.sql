alter table public.shift_logs
  add column if not exists operational_status text not null default 'Available',
  add column if not exists patrol_type text not null default 'Roads Policing',
  add column if not exists current_incident_id text;

drop policy if exists "shift read active units authenticated" on public.shift_logs;
create policy "shift read active units authenticated" on public.shift_logs for select
using (auth.role() = 'authenticated' and status = 'On Duty' and ended_at is null);

create table if not exists public.operational_incidents (
  incident_id text primary key default ('CAD_' || replace(gen_random_uuid()::text, '-', '')),
  incident_number text not null default ('MO8-' || to_char(now(), 'YYYYMMDD-HH24MISS')),
  title text not null,
  incident_type text not null default 'Traffic',
  priority text not null default 'Routine' check (priority in ('Emergency', 'Immediate', 'Priority', 'Routine')),
  location text not null default '',
  description text not null default '',
  status text not null default 'Open' check (status in ('Open', 'Assigned', 'En Route', 'On Scene', 'Resolved', 'Cancelled')),
  assigned_officer_ids text[] not null default array[]::text[],
  created_by text references public.profiles(user_id) on delete set null,
  closed_by text references public.profiles(user_id) on delete set null,
  closed_at timestamptz,
  outcome text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.operational_bolos (
  bolo_id text primary key default ('BOLO_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  bolo_type text not null default 'Vehicle',
  priority text not null default 'Normal' check (priority in ('Critical', 'High', 'Normal', 'Low')),
  description text not null default '',
  location text not null default '',
  expires_at timestamptz,
  status text not null default 'Active' check (status in ('Active', 'Expired', 'Cancelled')),
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_operational_incidents_status on public.operational_incidents(status, priority, created_at desc);
create index if not exists idx_operational_bolos_status on public.operational_bolos(status, created_at desc);

alter table public.operational_incidents enable row level security;
alter table public.operational_bolos enable row level security;

drop policy if exists "operational incidents read authenticated" on public.operational_incidents;
create policy "operational incidents read authenticated" on public.operational_incidents for select
using (auth.role() = 'authenticated');

drop policy if exists "operational incidents create authenticated" on public.operational_incidents;
create policy "operational incidents create authenticated" on public.operational_incidents for insert
with check (auth.role() = 'authenticated');

drop policy if exists "operational incidents update managers" on public.operational_incidents;
create policy "operational incidents update managers" on public.operational_incidents for update
using (public.has_permission('VIEW_TASKS') or created_by = public.current_user_id())
with check (public.has_permission('VIEW_TASKS') or created_by = public.current_user_id());

drop policy if exists "operational incidents delete command" on public.operational_incidents;
create policy "operational incidents delete command" on public.operational_incidents for delete
using (public.has_permission('FULL_ACCESS'));

drop policy if exists "operational bolos read authenticated" on public.operational_bolos;
create policy "operational bolos read authenticated" on public.operational_bolos for select
using (auth.role() = 'authenticated');

drop policy if exists "operational bolos create supervisors" on public.operational_bolos;
create policy "operational bolos create supervisors" on public.operational_bolos for insert
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "operational bolos update supervisors" on public.operational_bolos;
create policy "operational bolos update supervisors" on public.operational_bolos for update
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));
