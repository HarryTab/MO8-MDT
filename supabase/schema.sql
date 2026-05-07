-- MO8 MDT Supabase schema
-- Run this in Supabase SQL Editor after creating a new project.

create extension if not exists pgcrypto;

create table public.profiles (
  user_id text primary key default ('USR_' || replace(gen_random_uuid()::text, '-', '')),
  auth_user_id uuid unique references auth.users(id) on delete set null,
  member_id text not null default ('MBR_' || replace(gen_random_uuid()::text, '-', '')),
  roblox_username text not null unique,
  discord_id text,
  rank text not null default 'Police Constable',
  role text not null default 'Constable',
  status text not null default 'Active',
  last_login timestamptz,
  created_at timestamptz not null default now(),
  created_by text
);

create table public.officers (
  officer_id text primary key default ('OFF_' || replace(gen_random_uuid()::text, '-', '')),
  member_id text not null,
  roblox_username text not null unique,
  discord_id text,
  callsign text,
  rank text not null default 'Police Constable',
  status text not null default 'Active',
  join_date date,
  supervisor_user_id text references public.profiles(user_id) on delete set null,
  tags text[] not null default '{}',
  notes text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table public.training_options (
  option_id text primary key default ('TOP_' || replace(gen_random_uuid()::text, '-', '')),
  name text not null unique,
  type text not null,
  status text not null default 'Active',
  sort_order integer not null default 0,
  updated_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now()
);

create table public.training_records (
  training_id text primary key default ('TRN_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  standard text not null,
  status text not null default 'Not Started',
  assessor text,
  date_completed date,
  expiry_date date,
  notes text,
  updated_at timestamptz not null default now()
);

create table public.training_matrix (
  officer_id text primary key references public.officers(officer_id) on delete cascade,
  member_id text not null,
  roblox_username text not null,
  taser boolean not null default false,
  moe boolean not null default false,
  blue_ticket boolean not null default false,
  motorbike boolean not null default false,
  driving_standard text,
  review_date date,
  updated_at timestamptz not null default now(),
  updated_by text
);

create table public.training_courses (
  course_id text primary key default ('CRS_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  standard text not null,
  trainer_user_id text references public.profiles(user_id) on delete set null,
  course_date timestamptz,
  location text,
  capacity integer not null default 0,
  status text not null default 'Scheduled',
  notes text,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table public.course_bookings (
  booking_id text primary key default ('CBK_' || replace(gen_random_uuid()::text, '-', '')),
  course_id text not null references public.training_courses(course_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  status text not null default 'Requested',
  outcome text,
  notes text,
  requested_at timestamptz not null default now(),
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  unique (course_id, officer_id)
);

create table public.disciplinary_actions (
  action_id text primary key default ('DSC_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  type text not null,
  summary text not null,
  details text,
  issued_by text references public.profiles(user_id) on delete set null,
  issued_at timestamptz not null default now(),
  status text not null default 'Active'
);

create table public.loa_requests (
  request_id text primary key default ('LOA_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  start_date date not null,
  end_date date not null,
  reason text,
  status text not null default 'Pending',
  review_reason text,
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  created_at timestamptz not null default now()
);

create table public.transfer_requests (
  request_id text primary key default ('TRF_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  current_division text,
  target_division text not null,
  time_in_mo8 text,
  reason text,
  has_permission text,
  notes text,
  status text not null default 'Pending',
  review_reason text,
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  created_at timestamptz not null default now()
);

create table public.supervisor_requests (
  request_id text primary key default ('SPR_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  supervisor_user_id text references public.profiles(user_id) on delete set null,
  category text,
  subject text not null,
  details text,
  status text not null default 'Pending',
  review_reason text,
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  created_at timestamptz not null default now()
);

create table public.supervisor_checkins (
  checkin_id text primary key default ('CHK_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  supervisor_user_id text references public.profiles(user_id) on delete set null,
  checkin_date date not null,
  summary text,
  concerns text,
  development_goals text,
  follow_up_date date,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now()
);

create table public.development_plans (
  plan_id text primary key default ('PLN_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  supervisor_user_id text references public.profiles(user_id) on delete set null,
  goal text not null,
  category text,
  status text not null default 'Open',
  due_date date,
  notes text,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table public.appeals (
  appeal_id text primary key default ('APL_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  source_type text not null,
  source_id text not null,
  reason text,
  status text not null default 'Pending',
  review_reason text,
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  created_at timestamptz not null default now()
);

create table public.documents (
  document_id text primary key default ('DOC_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  category text not null default 'General',
  drive_url text not null,
  required_role text,
  required_tags text[] not null default '{}',
  requires_acknowledgement boolean not null default false,
  status text not null default 'Published',
  updated_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now()
);

create table public.document_acknowledgements (
  acknowledgement_id text primary key default ('ACK_' || replace(gen_random_uuid()::text, '-', '')),
  document_id text not null references public.documents(document_id) on delete cascade,
  member_id text,
  officer_id text references public.officers(officer_id) on delete cascade,
  user_id text references public.profiles(user_id) on delete cascade,
  acknowledged_at timestamptz not null default now(),
  unique (document_id, user_id)
);

create table public.dashboard_widgets (
  user_id text not null references public.profiles(user_id) on delete cascade,
  widget_key text not null,
  enabled boolean not null default true,
  updated_at timestamptz not null default now(),
  primary key (user_id, widget_key)
);

create table public.shift_logs (
  shift_id text primary key default ('SFT_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  member_id text,
  roblox_username text,
  callsign text,
  rank text,
  started_at timestamptz not null default now(),
  ended_at timestamptz,
  summary text,
  status text not null default 'On Duty',
  updated_at timestamptz not null default now()
);

create table public.announcements (
  announcement_id text primary key default ('ANN_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  body text,
  audience text,
  status text not null default 'Published',
  pinned boolean not null default false,
  expires_at date,
  updated_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now()
);

create table public.permissions (
  role text not null,
  permission text not null,
  allowed boolean not null default false,
  primary key (role, permission)
);

create table public.user_permissions (
  user_id text not null references public.profiles(user_id) on delete cascade,
  permission text not null,
  allowed text not null default 'Inherit',
  updated_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now(),
  primary key (user_id, permission)
);

create table public.audit_log (
  audit_id text primary key default ('AUD_' || replace(gen_random_uuid()::text, '-', '')),
  timestamp timestamptz not null default now(),
  actor_user_id text references public.profiles(user_id) on delete set null,
  action text not null,
  target_type text,
  target_id text,
  details jsonb not null default '{}'::jsonb
);

create table public.rank_changes (
  change_id text primary key default ('RCH_' || replace(gen_random_uuid()::text, '-', '')),
  member_id text,
  officer_id text references public.officers(officer_id) on delete set null,
  user_id text references public.profiles(user_id) on delete set null,
  roblox_username text,
  previous_rank text,
  new_rank text not null,
  reason text,
  changed_by text references public.profiles(user_id) on delete set null,
  changed_at timestamptz not null default now()
);

create table public.notifications (
  notification_id text primary key default ('NTF_' || replace(gen_random_uuid()::text, '-', '')),
  member_id text not null,
  title text not null,
  message text,
  created_at timestamptz not null default now(),
  read_at timestamptz,
  actor_user_id text references public.profiles(user_id) on delete set null
);

create index idx_officers_member_id on public.officers(member_id);
create index idx_officers_supervisor on public.officers(supervisor_user_id);
create index idx_training_records_officer on public.training_records(officer_id);
create index idx_course_bookings_course on public.course_bookings(course_id);
create index idx_course_bookings_officer on public.course_bookings(officer_id);
create index idx_loa_officer_status on public.loa_requests(officer_id, status);
create index idx_shift_logs_officer_started on public.shift_logs(officer_id, started_at desc);
create index idx_notifications_member_created on public.notifications(member_id, created_at desc);
create index idx_documents_category on public.documents(category);

create or replace function public.current_profile()
returns public.profiles
language sql
stable
security definer
set search_path = public
as $$
  select * from public.profiles where auth_user_id = auth.uid() limit 1
$$;

create or replace function public.current_user_id()
returns text
language sql
stable
security definer
set search_path = public
as $$
  select user_id from public.profiles where auth_user_id = auth.uid() limit 1
$$;

create or replace function public.current_member_id()
returns text
language sql
stable
security definer
set search_path = public
as $$
  select member_id from public.profiles where auth_user_id = auth.uid() limit 1
$$;

create or replace function public.has_permission(permission_name text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  with me as (
    select user_id, role from public.profiles where auth_user_id = auth.uid() limit 1
  )
  select coalesce((
    select case
      when up.allowed = 'Allow' then true
      when up.allowed = 'Deny' then false
      else null
    end
    from me
    join public.user_permissions up on up.user_id = me.user_id and up.permission = permission_name
    limit 1
  ), (
    select true
    from me
    join public.permissions p on p.role = me.role
    where p.allowed = true and p.permission in (permission_name, 'FULL_ACCESS')
    limit 1
  ), false)
$$;

create or replace function public.touch_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create trigger touch_officers_updated_at before update on public.officers for each row execute function public.touch_updated_at();
create trigger touch_training_courses_updated_at before update on public.training_courses for each row execute function public.touch_updated_at();
create trigger touch_documents_updated_at before update on public.documents for each row execute function public.touch_updated_at();
create trigger touch_announcements_updated_at before update on public.announcements for each row execute function public.touch_updated_at();
create trigger touch_development_plans_updated_at before update on public.development_plans for each row execute function public.touch_updated_at();

alter table public.profiles enable row level security;
alter table public.officers enable row level security;
alter table public.training_options enable row level security;
alter table public.training_records enable row level security;
alter table public.training_matrix enable row level security;
alter table public.training_courses enable row level security;
alter table public.course_bookings enable row level security;
alter table public.disciplinary_actions enable row level security;
alter table public.loa_requests enable row level security;
alter table public.transfer_requests enable row level security;
alter table public.supervisor_requests enable row level security;
alter table public.supervisor_checkins enable row level security;
alter table public.development_plans enable row level security;
alter table public.appeals enable row level security;
alter table public.documents enable row level security;
alter table public.document_acknowledgements enable row level security;
alter table public.dashboard_widgets enable row level security;
alter table public.shift_logs enable row level security;
alter table public.announcements enable row level security;
alter table public.permissions enable row level security;
alter table public.user_permissions enable row level security;
alter table public.audit_log enable row level security;
alter table public.rank_changes enable row level security;
alter table public.notifications enable row level security;

create policy "profiles can read self or managers can read users" on public.profiles for select
using (auth_user_id = auth.uid() or public.has_permission('MANAGE_USERS') or public.has_permission('VIEW_OFFICERS'));

create policy "managers can write profiles" on public.profiles for all
using (public.has_permission('MANAGE_USERS'))
with check (public.has_permission('MANAGE_USERS'));

create policy "officers visible to authorised users or self" on public.officers for select
using (public.has_permission('VIEW_OFFICERS') or member_id = public.current_member_id());

create policy "officer managers can write officers" on public.officers for all
using (public.has_permission('EDIT_OFFICERS') or public.has_permission('ADD_OFFICERS') or public.has_permission('ARCHIVE_OFFICERS'))
with check (public.has_permission('EDIT_OFFICERS') or public.has_permission('ADD_OFFICERS') or public.has_permission('ARCHIVE_OFFICERS'));

create policy "training read" on public.training_records for select
using (public.has_permission('VIEW_TRAINING') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "training write" on public.training_records for all
using (public.has_permission('MANAGE_TRAINING'))
with check (public.has_permission('MANAGE_TRAINING'));

create policy "training options read" on public.training_options for select
using (public.has_permission('VIEW_TRAINING') or public.has_permission('VIEW_COURSES'));

create policy "training options write" on public.training_options for all
using (public.has_permission('MANAGE_TRAINING_OPTIONS'))
with check (public.has_permission('MANAGE_TRAINING_OPTIONS'));

create policy "courses read" on public.training_courses for select
using (public.has_permission('VIEW_COURSES'));

create policy "courses write" on public.training_courses for all
using (public.has_permission('MANAGE_COURSES'))
with check (public.has_permission('MANAGE_COURSES'));

create policy "course bookings read" on public.course_bookings for select
using (public.has_permission('MANAGE_COURSES') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "course bookings insert own or managers" on public.course_bookings for insert
with check (public.has_permission('MANAGE_COURSES') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "course bookings update managers" on public.course_bookings for update
using (public.has_permission('MANAGE_COURSES'))
with check (public.has_permission('MANAGE_COURSES'));

create policy "discipline read" on public.disciplinary_actions for select
using (public.has_permission('VIEW_DISCIPLINE') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "discipline write" on public.disciplinary_actions for all
using (public.has_permission('ADD_DISCIPLINE'))
with check (public.has_permission('ADD_DISCIPLINE'));

create policy "loa read" on public.loa_requests for select
using (public.has_permission('VIEW_LOA') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "loa insert own or authorised" on public.loa_requests for insert
with check (public.has_permission('CREATE_LOA') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "loa update approvers" on public.loa_requests for update
using (public.has_permission('APPROVE_LOA') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id() and status = 'Pending'))
with check (public.has_permission('APPROVE_LOA') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "documents read" on public.documents for select
using (public.has_permission('VIEW_DOCUMENTS') and status = 'Published' or public.has_permission('MANAGE_DOCUMENTS'));

create policy "documents write" on public.documents for all
using (public.has_permission('MANAGE_DOCUMENTS'))
with check (public.has_permission('MANAGE_DOCUMENTS'));

create policy "announcements read" on public.announcements for select
using (public.has_permission('VIEW_ANNOUNCEMENTS') and status = 'Published' or public.has_permission('MANAGE_ANNOUNCEMENTS'));

create policy "announcements write" on public.announcements for all
using (public.has_permission('MANAGE_ANNOUNCEMENTS'))
with check (public.has_permission('MANAGE_ANNOUNCEMENTS'));

create policy "notifications read own" on public.notifications for select
using (member_id = public.current_member_id());

create policy "notifications update own" on public.notifications for update
using (member_id = public.current_member_id())
with check (member_id = public.current_member_id());

create policy "dashboard widgets own" on public.dashboard_widgets for all
using (user_id = public.current_user_id())
with check (user_id = public.current_user_id());

create policy "permissions read managers" on public.permissions for select
using (public.has_permission('MANAGE_PERMISSIONS'));

create policy "permissions write managers" on public.permissions for all
using (public.has_permission('MANAGE_PERMISSIONS'))
with check (public.has_permission('MANAGE_PERMISSIONS'));

create policy "user permissions read managers" on public.user_permissions for select
using (public.has_permission('MANAGE_PERMISSIONS'));

create policy "user permissions write managers" on public.user_permissions for all
using (public.has_permission('MANAGE_PERMISSIONS'))
with check (public.has_permission('MANAGE_PERMISSIONS'));

create policy "audit read authorised" on public.audit_log for select
using (public.has_permission('VIEW_AUDIT_LOG'));

create policy "rank changes read" on public.rank_changes for select
using (public.has_permission('VIEW_RANK_LOG') or member_id = public.current_member_id());

create policy "shift read own or managers" on public.shift_logs for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "shift write own" on public.shift_logs for insert
with check (officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

create policy "shift update own or managers" on public.shift_logs for update
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()))
with check (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));

insert into public.training_options (name, type, sort_order) values
  ('Taser', 'Specialist', 10),
  ('MOE', 'Specialist', 20),
  ('Blue Ticket', 'Specialist', 30),
  ('Motorbike', 'Specialist', 40),
  ('Basic', 'Driving', 50),
  ('Response', 'Driving', 60),
  ('IPP', 'Driving', 70),
  ('Advanced', 'Driving', 80),
  ('Advanced + TPAC', 'Driving', 90)
on conflict (name) do nothing;

insert into public.permissions (role, permission, allowed) values
  ('Constable', 'VIEW_DOCUMENTS', true),
  ('Constable', 'VIEW_ANNOUNCEMENTS', true),
  ('Constable', 'VIEW_COURSES', true),
  ('Constable', 'CHANGE_OWN_PASSWORD', true),
  ('Trainer', 'VIEW_DOCUMENTS', true),
  ('Trainer', 'VIEW_ANNOUNCEMENTS', true),
  ('Trainer', 'VIEW_TRAINING', true),
  ('Trainer', 'VIEW_COURSES', true),
  ('Trainer', 'MANAGE_COURSES', true),
  ('Trainer', 'CHANGE_OWN_PASSWORD', true),
  ('Sergeant', 'VIEW_DASHBOARD', true),
  ('Sergeant', 'VIEW_TASKS', true),
  ('Sergeant', 'VIEW_OFFICERS', true),
  ('Sergeant', 'VIEW_RANK_LOG', true),
  ('Sergeant', 'ASSIGN_SUPERVISORS', true),
  ('Sergeant', 'VIEW_TRAINING', true),
  ('Sergeant', 'MANAGE_TRAINING', true),
  ('Sergeant', 'VIEW_COURSES', true),
  ('Sergeant', 'VIEW_LOA', true),
  ('Sergeant', 'CREATE_LOA', true),
  ('Sergeant', 'APPROVE_LOA', true),
  ('Sergeant', 'VIEW_DOCUMENTS', true),
  ('Sergeant', 'VIEW_ANNOUNCEMENTS', true),
  ('Sergeant', 'CHANGE_OWN_PASSWORD', true),
  ('Inspector', 'FULL_ACCESS', true),
  ('Chief Inspector', 'FULL_ACCESS', true),
  ('Command', 'FULL_ACCESS', true)
on conflict (role, permission) do update set allowed = excluded.allowed;
