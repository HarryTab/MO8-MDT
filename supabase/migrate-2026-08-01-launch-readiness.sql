create table if not exists public.feature_flags (
  feature_key text primary key,
  label text not null,
  category text not null default 'Other',
  enabled boolean not null default true,
  description text not null default '',
  updated_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now()
);

create table if not exists public.bug_reports (
  report_id text primary key default ('BUG_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  area text not null default 'General',
  severity text not null default 'Normal',
  status text not null default 'Open',
  description text not null default '',
  reporter_user_id text references public.profiles(user_id) on delete set null,
  reporter_member_id text,
  reporter_username text not null default '',
  page_url text not null default '',
  user_agent text not null default '',
  review_notes text not null default '',
  reviewed_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.loa_requests add column if not exists handover_user_id text references public.profiles(user_id) on delete set null;
alter table public.loa_requests add column if not exists handover_notes text not null default '';

insert into public.feature_flags (feature_key, label, category, enabled, description)
values
  ('personnelHub', 'Personnel Hub', 'Core', true, 'Core personnel workspace shell'),
  ('dashboard', 'Dashboard', 'Home', true, 'Personnel overview and widgets'),
  ('inbox', 'Inbox', 'Home', true, 'Personal MDT inbox'),
  ('messaging', 'Chat and Messages', 'Home', true, 'Internal chat and formal messages'),
  ('tasks', 'Tasks', 'Home', true, 'Task queue and approvals'),
  ('calendar', 'Calendar', 'Home', true, 'Calendar, LOA and course events'),
  ('timeline', 'Timeline', 'Home', true, 'System and record timelines'),
  ('shift', 'Shift Logging', 'Personnel', true, 'Shift logging and activity reviews'),
  ('supervisor', 'Supervisor', 'Personnel', true, 'Supervisor dashboards and supervisee management'),
  ('officers', 'Officer Records', 'Personnel', true, 'Officer database and profiles'),
  ('development', 'Development and AIP', 'Personnel', true, 'Development plans, reviews, restrictions and AIPs'),
  ('rankChanges', 'Rank Log', 'Personnel', true, 'Promotion and rank history'),
  ('discipline', 'Discipline', 'Personnel', true, 'Disciplinary records'),
  ('loa', 'Leave of Absence', 'Personnel', true, 'LOA requests and approvals'),
  ('handover', 'Handover', 'Personnel', true, 'Command handovers'),
  ('courses', 'Courses', 'Learning', true, 'Course bookings and course management'),
  ('training', 'Training Matrix', 'Learning', true, 'Training records and capabilities'),
  ('documents', 'Documents', 'Learning', true, 'Document explorer and acknowledgements'),
  ('announcements', 'Notice Board', 'Management', true, 'Announcements and notices'),
  ('recruitment', 'Recruitment', 'Management', true, 'Vacancies and applications'),
  ('reports', 'Reports', 'Management', true, 'Command reports'),
  ('users', 'Users and Account Requests', 'Management', true, 'User accounts and access requests'),
  ('permissions', 'Permissions', 'Management', true, 'Role and user permissions'),
  ('audit', 'Audit Log', 'Management', true, 'Audit history'),
  ('dataQuality', 'Data Quality', 'Management', true, 'Data quality centre'),
  ('settings', 'Settings', 'Management', true, 'Administration and launch controls'),
  ('cad', 'Control and CAD', 'Operations', true, 'Operations hub and CAD')
on conflict (feature_key) do nothing;

alter table public.feature_flags enable row level security;
alter table public.bug_reports enable row level security;

drop policy if exists "feature flags read" on public.feature_flags;
create policy "feature flags read" on public.feature_flags for select using (public.current_user_id() is not null);

drop policy if exists "feature flags admin write" on public.feature_flags;
create policy "feature flags admin write" on public.feature_flags for all
using (public.has_permission('FULL_ACCESS'))
with check (public.has_permission('FULL_ACCESS'));

drop policy if exists "bug reports insert own" on public.bug_reports;
create policy "bug reports insert own" on public.bug_reports for insert
with check (reporter_user_id = public.current_user_id());

drop policy if exists "bug reports read relevant" on public.bug_reports;
create policy "bug reports read relevant" on public.bug_reports for select
using (reporter_user_id = public.current_user_id() or public.has_permission('FULL_ACCESS'));

drop policy if exists "bug reports admin update" on public.bug_reports;
create policy "bug reports admin update" on public.bug_reports for update
using (public.has_permission('FULL_ACCESS'))
with check (public.has_permission('FULL_ACCESS'));

create index if not exists idx_bug_reports_status on public.bug_reports(status, created_at desc);
create index if not exists idx_bug_reports_reporter on public.bug_reports(reporter_user_id, created_at desc);
