create table if not exists public.discord_alert_defaults (
  scope_type text not null check (scope_type in ('category', 'task')),
  scope_key text not null,
  label text not null,
  enabled boolean not null default true,
  locked boolean not null default false,
  description text not null default '',
  updated_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now(),
  primary key (scope_type, scope_key)
);

create table if not exists public.discord_alert_preferences (
  user_id text not null references public.profiles(user_id) on delete cascade,
  scope_type text not null check (scope_type in ('category', 'task')),
  scope_key text not null,
  enabled boolean not null default true,
  updated_at timestamptz not null default now(),
  primary key (user_id, scope_type, scope_key)
);

alter table public.discord_alert_defaults enable row level security;
alter table public.discord_alert_preferences enable row level security;

drop policy if exists "discord alert defaults read" on public.discord_alert_defaults;
create policy "discord alert defaults read" on public.discord_alert_defaults for select
using (public.current_user_id() is not null);

drop policy if exists "discord alert defaults admin write" on public.discord_alert_defaults;
create policy "discord alert defaults admin write" on public.discord_alert_defaults for all
using (public.has_permission('FULL_ACCESS'))
with check (public.has_permission('FULL_ACCESS'));

drop policy if exists "discord alert preferences read own" on public.discord_alert_preferences;
create policy "discord alert preferences read own" on public.discord_alert_preferences for select
using (user_id = public.current_user_id() or public.has_permission('FULL_ACCESS'));

drop policy if exists "discord alert preferences write own" on public.discord_alert_preferences;
create policy "discord alert preferences write own" on public.discord_alert_preferences for all
using (user_id = public.current_user_id() or public.has_permission('FULL_ACCESS'))
with check (user_id = public.current_user_id() or public.has_permission('FULL_ACCESS'));

insert into public.discord_alert_defaults (scope_type, scope_key, label, enabled, locked, description)
values
  ('category', 'critical', 'Critical alerts', true, true, 'Urgent or safety-critical alerts. These always send.'),
  ('category', 'task', 'Task alerts', true, false, 'Assigned work, reviews, approvals and amendments.'),
  ('category', 'profile', 'Profile updates', true, false, 'Training, discipline, supervisor, AIP and officer-record changes.'),
  ('category', 'calendar', 'Calendar alerts', true, false, 'Calendar assignments, reminders and changes.'),
  ('category', 'course', 'Training course alerts', true, false, 'Course booking requests, approvals, waitlists and cancellations.'),
  ('category', 'loa', 'LOA and transfer alerts', true, false, 'Leave, task-cover and transfer request decisions.'),
  ('category', 'cad', 'CAD and operations alerts', true, false, 'CAD assignments, BOLOs, callsign and operational actions.'),
  ('category', 'message', 'Messages and handovers', true, false, 'Formal messages, chat acknowledgements and command handovers.'),
  ('category', 'recruitment', 'Recruitment alerts', true, false, 'Applications, vacancies and account request updates.'),
  ('task', 'LOA Approval', 'LOA approvals', true, false, 'Tasks to review leave of absence requests.'),
  ('task', 'Transfer Request', 'Transfer requests', true, false, 'Tasks to review transfer requests.'),
  ('task', 'Supervisor Request', 'Supervisor support requests', true, false, 'Tasks for supervisee support and contact requests.'),
  ('task', 'Appeal / Review', 'Appeals and reviews', true, false, 'Tasks to review submitted appeals.'),
  ('task', 'Course Booking', 'Course booking requests', true, false, 'Tasks to approve, deny or waitlist course seats.'),
  ('task', 'Account Request', 'Account creation requests', true, false, 'Command tasks for pending account requests.'),
  ('task', 'Retrospective Shift', 'Retrospective shift requests', true, false, 'Tasks to approve backdated shift logs.'),
  ('task', 'Probation Review', 'Probation reviews', true, false, 'Supervisor probation review tasks.'),
  ('task', 'Performance Review', 'Performance reviews', true, false, 'Supervisor performance review tasks.'),
  ('task', 'Restriction Review', 'Restriction reviews', true, false, 'Tasks to review temporary restrictions.'),
  ('task', 'Activity Review', 'Activity reviews', true, false, 'Tasks to review low or missing activity.'),
  ('task', 'Training Review', 'Training reviews', true, false, 'Tasks to review training gaps or expiry.'),
  ('task', 'AIP Signature', 'AIP signatures', true, false, 'Tasks to sign Activity Improvement Notices.'),
  ('task', 'AIP Review', 'AIP reviews', true, false, 'Tasks to review AIP outcomes.'),
  ('task', 'CAD Review', 'CAD reviews', true, false, 'Supervisor CAD review tasks.'),
  ('task', 'After Action Review', 'After-action reviews', true, false, 'Tasks for operational learning reviews.'),
  ('task', 'Operational Action Review', 'Operational action reviews', true, false, 'Tasks to review arrests, powers and outcomes.'),
  ('task', 'CAD Amendment', 'CAD amendments', true, false, 'Returned CAD records requiring amendments.'),
  ('task', 'Operational Action Amendment', 'Operational action amendments', true, false, 'Returned operational actions requiring amendments.'),
  ('task', 'Shift Review', 'Shift reviews', true, false, 'Tasks to review completed shift debriefs.'),
  ('task', 'Shift Amendment', 'Shift amendments', true, false, 'Tasks to amend returned shift debriefs.'),
  ('task', 'Casework Action', 'Casework actions', true, false, 'Assigned casework actions.'),
  ('task', 'Recruitment Application', 'Recruitment applications', true, false, 'Application review tasks.'),
  ('task', 'Assigned Task', 'Manual assigned tasks', true, false, 'Tasks manually assigned by another officer.')
on conflict (scope_type, scope_key) do update
set label = excluded.label,
    description = excluded.description,
    locked = excluded.locked;

create index if not exists idx_discord_alert_preferences_user on public.discord_alert_preferences(user_id);
