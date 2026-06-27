alter table public.operational_offences
  add column if not exists suggested_outcome text not null default 'FPN',
  add column if not exists allowed_outcomes text[] not null default array['FPN', 'Warning', 'Arrested', 'Reported for Summons', 'No Further Action']::text[],
  add column if not exists escalation_guidance text not null default '';

alter table public.operational_disposals
  add column if not exists issuing_officer_id text references public.officers(officer_id) on delete set null,
  add column if not exists assisting_officer_ids text[] not null default array[]::text[];

update public.operational_disposals
set outcome_type = 'Arrested'
where outcome_type in ('Arrest', 'Prison Sentence');

update public.operational_incidents
set closure_code = 'Arrested'
where closure_code = 'Arrest';

alter table public.operational_incident_reviews
  add column if not exists assigned_reviewer_user_id text references public.profiles(user_id) on delete set null,
  add column if not exists task_status text not null default 'Pending' check (task_status in ('Pending', 'In Progress', 'Completed', 'Cancelled')),
  add column if not exists due_at timestamptz;

update public.operational_incident_reviews review
set assigned_reviewer_user_id = officer.supervisor_user_id,
    due_at = coalesce(review.due_at, review.created_at + interval '7 days')
from public.operational_incidents incident
join public.profiles creator on creator.user_id = incident.created_by
join public.officers officer on officer.member_id = creator.member_id
where review.incident_id = incident.incident_id
  and review.assigned_reviewer_user_id is null
  and officer.supervisor_user_id is not null;

create table if not exists public.operational_officer_actions (
  action_id text primary key default ('OACT_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  assisting_officer_ids text[] not null default array[]::text[],
  action_type text not null,
  person_id text references public.operational_persons(person_id) on delete set null,
  vehicle_id text references public.operational_vehicles(vehicle_id) on delete set null,
  offence_id text references public.operational_offences(offence_id) on delete set null,
  disposal_id text references public.operational_disposals(disposal_id) on delete set null,
  notes text not null default '',
  occurred_at timestamptz not null default now(),
  created_by text references public.profiles(user_id) on delete set null
);

update public.operational_disposals disposal
set issuing_officer_id = officer.officer_id
from public.profiles profile
join public.officers officer on officer.member_id = profile.member_id
where disposal.issued_by = profile.user_id
  and disposal.issuing_officer_id is null;

insert into public.operational_officer_actions (
  incident_id, officer_id, action_type, person_id, vehicle_id, offence_id,
  disposal_id, notes, occurred_at, created_by
)
select disposal.incident_id,
       disposal.issuing_officer_id,
       case disposal.outcome_type
         when 'FPN' then 'FPN Issued'
         when 'Warning' then 'Warning Issued'
         when 'Reported for Summons' then 'Report'
         else disposal.outcome_type
       end,
       disposal.person_id, disposal.vehicle_id, disposal.offence_id,
       disposal.disposal_id, disposal.notes, disposal.issued_at, disposal.issued_by
from public.operational_disposals disposal
where disposal.issuing_officer_id is not null
  and not exists (
    select 1 from public.operational_officer_actions action
    where action.disposal_id = disposal.disposal_id
  );

create table if not exists public.operational_intel_acknowledgements (
  acknowledgement_id text primary key default ('IACK_' || replace(gen_random_uuid()::text, '-', '')),
  marker_id text not null references public.operational_intel_markers(marker_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  incident_id text references public.operational_incidents(incident_id) on delete set null,
  acknowledged_at timestamptz not null default now(),
  unique (marker_id, officer_id, incident_id)
);

alter table public.operational_bolos
  add column if not exists image_storage_path text not null default '',
  add column if not exists image_file_name text not null default '',
  add column if not exists image_file_type text not null default '';

create table if not exists public.operational_bolo_sightings (
  sighting_id text primary key default ('BSIGHT_' || replace(gen_random_uuid()::text, '-', '')),
  bolo_id text not null references public.operational_bolos(bolo_id) on delete cascade,
  incident_id text references public.operational_incidents(incident_id) on delete set null,
  location text not null,
  details text not null default '',
  sighted_at timestamptz not null default now(),
  reported_by_officer_id text references public.officers(officer_id) on delete set null,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.operational_after_action_reviews (
  after_action_id text primary key default ('AAR_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null unique references public.operational_incidents(incident_id) on delete cascade,
  assigned_reviewer_user_id text references public.profiles(user_id) on delete set null,
  status text not null default 'Pending' check (status in ('Pending', 'In Progress', 'Completed', 'Cancelled')),
  decision_summary text not null default '',
  learning_points text not null default '',
  follow_up_actions text not null default '',
  completed_by text references public.profiles(user_id) on delete set null,
  due_at timestamptz,
  completed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_operational_actions_officer on public.operational_officer_actions(officer_id, occurred_at desc);
create index if not exists idx_operational_actions_incident on public.operational_officer_actions(incident_id, occurred_at desc);
create index if not exists idx_operational_actions_type on public.operational_officer_actions(action_type, occurred_at desc);
create index if not exists idx_operational_acknowledgements_marker on public.operational_intel_acknowledgements(marker_id, officer_id);
create index if not exists idx_operational_bolo_sightings_bolo on public.operational_bolo_sightings(bolo_id, sighted_at desc);
create index if not exists idx_operational_reviews_assignee on public.operational_incident_reviews(assigned_reviewer_user_id, task_status, due_at);
create index if not exists idx_operational_aar_assignee on public.operational_after_action_reviews(assigned_reviewer_user_id, status, due_at);

alter table public.operational_officer_actions enable row level security;
alter table public.operational_intel_acknowledgements enable row level security;
alter table public.operational_bolo_sightings enable row level security;
alter table public.operational_after_action_reviews enable row level security;

drop policy if exists "operational officer actions authenticated" on public.operational_officer_actions;
create policy "operational officer actions authenticated" on public.operational_officer_actions for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational marker acknowledgements authenticated" on public.operational_intel_acknowledgements;
create policy "operational marker acknowledgements authenticated" on public.operational_intel_acknowledgements for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational bolo sightings authenticated" on public.operational_bolo_sightings;
create policy "operational bolo sightings authenticated" on public.operational_bolo_sightings for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational after action read authenticated" on public.operational_after_action_reviews;
create policy "operational after action read authenticated" on public.operational_after_action_reviews for select using (auth.role() = 'authenticated');
drop policy if exists "operational after action manage supervisors" on public.operational_after_action_reviews;
create policy "operational after action manage supervisors" on public.operational_after_action_reviews for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

insert into storage.buckets (id, name, public, file_size_limit)
values ('mo8-bolo-images', 'mo8-bolo-images', false, 10485760)
on conflict (id) do update set file_size_limit = excluded.file_size_limit;

drop policy if exists "bolo images read authenticated" on storage.objects;
create policy "bolo images read authenticated" on storage.objects for select
using (bucket_id = 'mo8-bolo-images' and auth.role() = 'authenticated');
drop policy if exists "bolo images upload supervisors" on storage.objects;
create policy "bolo images upload supervisors" on storage.objects for insert
with check (bucket_id = 'mo8-bolo-images' and public.has_permission('VIEW_TASKS'));
drop policy if exists "bolo images update supervisors" on storage.objects;
create policy "bolo images update supervisors" on storage.objects for update
using (bucket_id = 'mo8-bolo-images' and public.has_permission('VIEW_TASKS'))
with check (bucket_id = 'mo8-bolo-images' and public.has_permission('VIEW_TASKS'));
drop policy if exists "bolo images delete supervisors" on storage.objects;
create policy "bolo images delete supervisors" on storage.objects for delete
using (bucket_id = 'mo8-bolo-images' and public.has_permission('VIEW_TASKS'));
