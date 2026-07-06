create table if not exists public.task_resolutions (
  resolution_id text primary key default ('TSKR_' || replace(gen_random_uuid()::text, '-', '')),
  task_key text not null,
  task_type text not null,
  source_id text not null default '',
  officer_id text references public.officers(officer_id) on delete set null,
  resolved_by text not null references public.profiles(user_id) on delete restrict,
  outcome text not null,
  notes text not null default '',
  follow_up_date date,
  suppress_until timestamptz,
  created_at timestamptz not null default now(),
  unique (task_key, resolved_by)
);

create index if not exists idx_task_resolutions_key on public.task_resolutions(task_key, suppress_until);
create index if not exists idx_task_resolutions_officer on public.task_resolutions(officer_id, created_at desc);

alter table public.task_resolutions enable row level security;

drop policy if exists "task resolutions read authenticated" on public.task_resolutions;
create policy "task resolutions read authenticated" on public.task_resolutions for select
using (
  resolved_by in (select user_id from public.profiles where member_id = public.current_member_id())
  or public.has_permission('VIEW_TASKS')
);

drop policy if exists "task resolutions create own" on public.task_resolutions;
create policy "task resolutions create own" on public.task_resolutions for insert
with check (
  resolved_by in (select user_id from public.profiles where member_id = public.current_member_id())
);

drop policy if exists "task resolutions update own or command" on public.task_resolutions;
create policy "task resolutions update own or command" on public.task_resolutions for update
using (
  resolved_by in (select user_id from public.profiles where member_id = public.current_member_id())
  or public.has_permission('VIEW_TASKS')
) with check (
  resolved_by in (select user_id from public.profiles where member_id = public.current_member_id())
  or public.has_permission('VIEW_TASKS')
);

alter table public.activity_improvement_notices
  add column if not exists archived_at timestamptz,
  add column if not exists archived_by text references public.profiles(user_id) on delete set null,
  add column if not exists archive_reason text not null default '';

create index if not exists idx_aip_archived on public.activity_improvement_notices(archived_at, created_at desc);
