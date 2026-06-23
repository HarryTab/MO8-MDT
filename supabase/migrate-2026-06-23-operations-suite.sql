-- MO8 MDT operations suite
-- Run once in the Supabase SQL Editor.

create table if not exists public.saved_views (
  view_id text primary key default ('VIW_' || replace(gen_random_uuid()::text, '-', '')),
  user_id text not null references public.profiles(user_id) on delete cascade,
  name text not null,
  module text not null default 'Search',
  query text,
  filters jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  unique (user_id, name)
);

create table if not exists public.calendar_events (
  event_id text primary key default ('EVT_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  event_type text not null default 'Operational',
  starts_at timestamptz not null,
  ends_at timestamptz,
  location text,
  details text,
  audience text default 'All Officers',
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.probation_records (
  probation_id text primary key default ('PRB_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  stage text not null default 'Initial',
  status text not null default 'Active',
  start_date date,
  target_date date,
  progress integer not null default 0 check (progress between 0 and 100),
  requirements text,
  notes text,
  reviewer_user_id text references public.profiles(user_id) on delete set null,
  updated_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now()
);

create table if not exists public.performance_reviews (
  review_id text primary key default ('REV_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  review_date date not null default current_date,
  period_start date,
  period_end date,
  rating text not null default 'Meets Expectations',
  activity_summary text,
  strengths text,
  improvements text,
  objectives text,
  next_review_date date,
  reviewer_user_id text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.officer_restrictions (
  restriction_id text primary key default ('RST_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  restriction_type text not null,
  details text,
  starts_on date not null default current_date,
  ends_on date,
  status text not null default 'Active',
  imposed_by text references public.profiles(user_id) on delete set null,
  reviewed_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now()
);

create table if not exists public.handover_entries (
  handover_id text primary key default ('HND_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  category text not null default 'Operational',
  priority text not null default 'Normal',
  details text not null,
  owner_user_id text references public.profiles(user_id) on delete set null,
  due_at timestamptz,
  status text not null default 'Open',
  resolution text,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_saved_views_user on public.saved_views(user_id);
create index if not exists idx_calendar_events_start on public.calendar_events(starts_at);
create index if not exists idx_probation_officer on public.probation_records(officer_id);
create index if not exists idx_reviews_officer_date on public.performance_reviews(officer_id, review_date desc);
create index if not exists idx_restrictions_officer_status on public.officer_restrictions(officer_id, status);
create index if not exists idx_handover_status_priority on public.handover_entries(status, priority);

alter table public.saved_views enable row level security;
alter table public.calendar_events enable row level security;
alter table public.probation_records enable row level security;
alter table public.performance_reviews enable row level security;
alter table public.officer_restrictions enable row level security;
alter table public.handover_entries enable row level security;

drop policy if exists "saved views own" on public.saved_views;
create policy "saved views own" on public.saved_views for all
using (user_id = public.current_user_id())
with check (user_id = public.current_user_id());

drop policy if exists "calendar read authenticated" on public.calendar_events;
drop policy if exists "calendar manage supervisors" on public.calendar_events;
create policy "calendar read authenticated" on public.calendar_events for select
using (auth.uid() is not null);
create policy "calendar manage supervisors" on public.calendar_events for all
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "probation read own or supervisors" on public.probation_records;
drop policy if exists "probation manage supervisors" on public.probation_records;
create policy "probation read own or supervisors" on public.probation_records for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "probation manage supervisors" on public.probation_records for all
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "reviews read own or supervisors" on public.performance_reviews;
drop policy if exists "reviews manage supervisors" on public.performance_reviews;
create policy "reviews read own or supervisors" on public.performance_reviews for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "reviews manage supervisors" on public.performance_reviews for all
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "restrictions read own or supervisors" on public.officer_restrictions;
drop policy if exists "restrictions manage supervisors" on public.officer_restrictions;
create policy "restrictions read own or supervisors" on public.officer_restrictions for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "restrictions manage supervisors" on public.officer_restrictions for all
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "handover supervisors" on public.handover_entries;
create policy "handover supervisors" on public.handover_entries for all
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

