alter table public.operational_incidents
  add column if not exists closure_code text not null default '',
  add column if not exists incident_commander_user_id text references public.profiles(user_id) on delete set null,
  add column if not exists command_roles jsonb not null default '{}'::jsonb,
  add column if not exists linked_incident_ids text[] not null default array[]::text[],
  add column if not exists linked_bolo_ids text[] not null default array[]::text[],
  add column if not exists review_status text not null default 'Not Required',
  add column if not exists review_due_at timestamptz,
  add column if not exists reviewed_at timestamptz,
  add column if not exists duplicate_of_incident_id text references public.operational_incidents(incident_id) on delete set null,
  add column if not exists archived_at timestamptz;

create table if not exists public.operational_incident_attendance (
  attendance_id text primary key default ('IATT_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  unit_id text references public.operational_units(unit_id) on delete set null,
  callsign_snapshot text not null default '',
  assigned_at timestamptz not null default now(),
  en_route_at timestamptz,
  on_scene_at timestamptz,
  cleared_at timestamptz,
  outcome text not null default '',
  created_by text references public.profiles(user_id) on delete set null
);

create table if not exists public.operational_incident_reviews (
  review_id text primary key default ('IREV_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  reviewer_user_id text references public.profiles(user_id) on delete set null,
  status text not null default 'Pending' check (status in ('Pending', 'Approved', 'Amendments Required', 'Completed')),
  feedback text not null default '',
  created_at timestamptz not null default now(),
  completed_at timestamptz
);

create table if not exists public.operational_incident_attachments (
  attachment_id text primary key default ('IATTACH_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  title text not null default '',
  storage_path text not null,
  file_name text not null default '',
  file_type text not null default '',
  file_size bigint not null default 0,
  uploaded_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.operational_incident_templates (
  template_id text primary key default ('ITPL_' || replace(gen_random_uuid()::text, '-', '')),
  name text not null unique,
  incident_type text not null default 'Traffic',
  default_priority text not null default 'Routine',
  title_template text not null default '',
  description_template text not null default '',
  required_capabilities text[] not null default array[]::text[],
  closure_codes text[] not null default array[]::text[],
  active boolean not null default true,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

insert into public.operational_incident_templates
  (name, incident_type, default_priority, title_template, description_template, required_capabilities, closure_codes)
values
  ('Road Traffic Collision', 'RTC', 'Priority', 'Road traffic collision', 'Vehicles involved:\nInjuries:\nRoad condition:\nObstruction:', array['Response'], array['No Further Action', 'Reported', 'Arrest', 'Words of Advice']),
  ('Vehicle Pursuit', 'Pursuit', 'Immediate', 'Vehicle pursuit', 'Vehicle:\nDirection of travel:\nGrounds:\nRisk level:', array['IPP'], array['Stopped', 'Abandoned', 'Lost', 'Arrest']),
  ('Officer Assistance', 'Assist Officer', 'Emergency', 'Officer requires assistance', 'Requesting unit:\nThreat:\nResources required:', array['Supervisor'], array['Resolved', 'Arrest', 'No Further Action']),
  ('Traffic Stop', 'Traffic', 'Routine', 'Traffic stop', 'Vehicle:\nReason for stop:\nOccupants:', array['Basic'], array['No Further Action', 'Words of Advice', 'Reported', 'Seized', 'Arrest'])
on conflict (name) do nothing;

create index if not exists idx_operational_incidents_archive on public.operational_incidents(status, closed_at desc);
create index if not exists idx_operational_incidents_location on public.operational_incidents(lower(location));
create index if not exists idx_operational_attendance_incident on public.operational_incident_attendance(incident_id, assigned_at);
create index if not exists idx_operational_reviews_status on public.operational_incident_reviews(status, created_at desc);
create index if not exists idx_operational_attachments_incident on public.operational_incident_attachments(incident_id, created_at desc);

alter table public.operational_incident_attendance enable row level security;
alter table public.operational_incident_reviews enable row level security;
alter table public.operational_incident_attachments enable row level security;
alter table public.operational_incident_templates enable row level security;

drop policy if exists "operational attendance authenticated" on public.operational_incident_attendance;
create policy "operational attendance authenticated" on public.operational_incident_attendance for all
using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

drop policy if exists "operational reviews read authenticated" on public.operational_incident_reviews;
create policy "operational reviews read authenticated" on public.operational_incident_reviews for select
using (auth.role() = 'authenticated');
drop policy if exists "operational reviews manage supervisors" on public.operational_incident_reviews;
create policy "operational reviews manage supervisors" on public.operational_incident_reviews for all
using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "operational attachments read authenticated" on public.operational_incident_attachments;
create policy "operational attachments read authenticated" on public.operational_incident_attachments for select
using (auth.role() = 'authenticated');
drop policy if exists "operational attachments write authenticated" on public.operational_incident_attachments;
create policy "operational attachments write authenticated" on public.operational_incident_attachments for insert
with check (auth.role() = 'authenticated');
drop policy if exists "operational attachments delete supervisors" on public.operational_incident_attachments;
create policy "operational attachments delete supervisors" on public.operational_incident_attachments for delete
using (public.has_permission('VIEW_TASKS'));

drop policy if exists "operational templates read authenticated" on public.operational_incident_templates;
create policy "operational templates read authenticated" on public.operational_incident_templates for select
using (auth.role() = 'authenticated');
drop policy if exists "operational templates manage supervisors" on public.operational_incident_templates;
create policy "operational templates manage supervisors" on public.operational_incident_templates for all
using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

insert into storage.buckets (id, name, public, file_size_limit)
values ('mo8-cad-evidence', 'mo8-cad-evidence', false, 10485760)
on conflict (id) do update set file_size_limit = excluded.file_size_limit;

drop policy if exists "cad evidence read authenticated" on storage.objects;
create policy "cad evidence read authenticated" on storage.objects for select
using (bucket_id = 'mo8-cad-evidence' and auth.role() = 'authenticated');
drop policy if exists "cad evidence upload authenticated" on storage.objects;
create policy "cad evidence upload authenticated" on storage.objects for insert
with check (bucket_id = 'mo8-cad-evidence' and auth.role() = 'authenticated');
drop policy if exists "cad evidence delete supervisors" on storage.objects;
create policy "cad evidence delete supervisors" on storage.objects for delete
using (bucket_id = 'mo8-cad-evidence' and public.has_permission('VIEW_TASKS'));
