alter table public.operational_incident_reviews
  add column if not exists amendment_assignee_user_id text references public.profiles(user_id) on delete set null,
  add column if not exists amendment_response text not null default '',
  add column if not exists amendment_submitted_at timestamptz;

alter table public.operational_action_reviews
  add column if not exists amendment_response text not null default '',
  add column if not exists amendment_submitted_at timestamptz;

alter table public.operational_operations
  add column if not exists case_type text not null default 'Casework',
  add column if not exists summary text not null default '',
  add column if not exists sensitivity text not null default 'Internal',
  add column if not exists review_date date,
  add column if not exists command_structure jsonb not null default '{}'::jsonb;

create table if not exists public.operational_case_actions (
  action_id text primary key default ('CACT_' || replace(gen_random_uuid()::text, '-', '')),
  operation_id text not null references public.operational_operations(operation_id) on delete cascade,
  title text not null,
  details text not null default '',
  priority text not null default 'Normal' check (priority in ('Critical', 'High', 'Normal', 'Low')),
  status text not null default 'Open' check (status in ('Open', 'In Progress', 'Blocked', 'Completed', 'Cancelled')),
  assigned_officer_id text references public.officers(officer_id) on delete set null,
  due_at timestamptz,
  completed_at timestamptz,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.operational_case_updates (
  update_id text primary key default ('CUPD_' || replace(gen_random_uuid()::text, '-', '')),
  operation_id text not null references public.operational_operations(operation_id) on delete cascade,
  update_type text not null default 'Update',
  body text not null,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.operational_briefing_acknowledgements (
  acknowledgement_id text primary key default ('BACK_' || replace(gen_random_uuid()::text, '-', '')),
  briefing_id text not null references public.operational_briefings(briefing_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  acknowledged_at timestamptz not null default now(),
  unique (briefing_id, officer_id)
);

create table if not exists public.operational_deployments (
  deployment_id text primary key default ('DEPL_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  deployment_type text not null default 'Planned Operation',
  status text not null default 'Planned' check (status in ('Draft', 'Planned', 'Open for Registration', 'Confirmed', 'Active', 'Completed', 'Cancelled')),
  starts_at timestamptz not null,
  ends_at timestamptz,
  location text not null default '',
  briefing text not null default '',
  required_capabilities text[] not null default array[]::text[],
  assigned_officer_ids text[] not null default array[]::text[],
  capacity integer not null default 0,
  lead_user_id text references public.profiles(user_id) on delete set null,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.operational_control_events (
  event_id text primary key default ('CEVT_' || replace(gen_random_uuid()::text, '-', '')),
  event_type text not null default 'Control',
  summary text not null,
  incident_id text references public.operational_incidents(incident_id) on delete set null,
  unit_id text references public.operational_units(unit_id) on delete set null,
  actor_user_id text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.operational_deployment_responses (
  response_id text primary key default ('DRES_' || replace(gen_random_uuid()::text, '-', '')),
  deployment_id text not null references public.operational_deployments(deployment_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  response text not null default 'Available' check (response in ('Available', 'Maybe', 'Unavailable')),
  note text not null default '',
  responded_at timestamptz not null default now(),
  unique (deployment_id, officer_id)
);

create index if not exists idx_case_actions_owner on public.operational_case_actions(assigned_officer_id, status, due_at);
create index if not exists idx_case_updates_case on public.operational_case_updates(operation_id, created_at desc);
create index if not exists idx_briefing_ack_officer on public.operational_briefing_acknowledgements(officer_id, briefing_id);
create index if not exists idx_deployments_schedule on public.operational_deployments(status, starts_at);
create index if not exists idx_control_events_time on public.operational_control_events(created_at desc);
create index if not exists idx_deployment_responses_deployment on public.operational_deployment_responses(deployment_id, response);

alter table public.operational_case_actions enable row level security;
alter table public.operational_case_updates enable row level security;
alter table public.operational_briefing_acknowledgements enable row level security;
alter table public.operational_deployments enable row level security;
alter table public.operational_control_events enable row level security;
alter table public.operational_deployment_responses enable row level security;

drop policy if exists "case actions read authenticated" on public.operational_case_actions;
create policy "case actions read authenticated" on public.operational_case_actions for select using (auth.role() = 'authenticated');
drop policy if exists "case actions write supervisors" on public.operational_case_actions;
create policy "case actions write supervisors" on public.operational_case_actions for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));
drop policy if exists "case actions update assignee" on public.operational_case_actions;
create policy "case actions update assignee" on public.operational_case_actions for update
using (exists (select 1 from public.officers o join public.profiles p on p.member_id = o.member_id where o.officer_id = operational_case_actions.assigned_officer_id and p.user_id = public.current_user_id()))
with check (exists (select 1 from public.officers o join public.profiles p on p.member_id = o.member_id where o.officer_id = operational_case_actions.assigned_officer_id and p.user_id = public.current_user_id()));

drop policy if exists "case updates read authenticated" on public.operational_case_updates;
create policy "case updates read authenticated" on public.operational_case_updates for select using (auth.role() = 'authenticated');
drop policy if exists "case updates write authenticated" on public.operational_case_updates;
create policy "case updates write authenticated" on public.operational_case_updates for insert with check (auth.role() = 'authenticated');

drop policy if exists "briefing acknowledgements read authenticated" on public.operational_briefing_acknowledgements;
create policy "briefing acknowledgements read authenticated" on public.operational_briefing_acknowledgements for select using (auth.role() = 'authenticated');
drop policy if exists "briefing acknowledgements write own" on public.operational_briefing_acknowledgements;
create policy "briefing acknowledgements write own" on public.operational_briefing_acknowledgements for insert
with check (exists (select 1 from public.officers o join public.profiles p on p.member_id = o.member_id where o.officer_id = operational_briefing_acknowledgements.officer_id and p.user_id = public.current_user_id()));

drop policy if exists "deployments read authenticated" on public.operational_deployments;
create policy "deployments read authenticated" on public.operational_deployments for select using (auth.role() = 'authenticated');
drop policy if exists "deployments manage supervisors" on public.operational_deployments;
create policy "deployments manage supervisors" on public.operational_deployments for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "control events read authenticated" on public.operational_control_events;
create policy "control events read authenticated" on public.operational_control_events for select using (auth.role() = 'authenticated');
drop policy if exists "control events write authenticated" on public.operational_control_events;
create policy "control events write authenticated" on public.operational_control_events for insert with check (auth.role() = 'authenticated');

drop policy if exists "deployment responses read authenticated" on public.operational_deployment_responses;
create policy "deployment responses read authenticated" on public.operational_deployment_responses for select using (auth.role() = 'authenticated');
drop policy if exists "deployment responses write own" on public.operational_deployment_responses;
create policy "deployment responses write own" on public.operational_deployment_responses for all
using (exists (select 1 from public.officers o join public.profiles p on p.member_id = o.member_id where o.officer_id = operational_deployment_responses.officer_id and p.user_id = public.current_user_id()))
with check (exists (select 1 from public.officers o join public.profiles p on p.member_id = o.member_id where o.officer_id = operational_deployment_responses.officer_id and p.user_id = public.current_user_id()));

-- Correct the subject-officer policies introduced by the review amendment migration.
drop policy if exists "operational reviews amend assigned officer" on public.operational_incident_reviews;
create policy "operational reviews amend assigned officer" on public.operational_incident_reviews for update
using (amendment_assignee_user_id = public.current_user_id())
with check (amendment_assignee_user_id = public.current_user_id());

drop policy if exists "operational action reviews amend subject officer" on public.operational_action_reviews;
create policy "operational action reviews amend subject officer" on public.operational_action_reviews for update
using (exists (select 1 from public.officers o join public.profiles p on p.member_id = o.member_id where o.officer_id = operational_action_reviews.officer_id and p.user_id = public.current_user_id()))
with check (exists (select 1 from public.officers o join public.profiles p on p.member_id = o.member_id where o.officer_id = operational_action_reviews.officer_id and p.user_id = public.current_user_id()));

drop policy if exists "operational action reviews read subject officer" on public.operational_action_reviews;
create policy "operational action reviews read subject officer" on public.operational_action_reviews for select
using (exists (select 1 from public.officers o join public.profiles p on p.member_id = o.member_id where o.officer_id = operational_action_reviews.officer_id and p.user_id = public.current_user_id()));
