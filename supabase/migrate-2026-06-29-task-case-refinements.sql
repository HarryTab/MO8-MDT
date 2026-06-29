alter table public.operational_operations
  add column if not exists archived_at timestamptz,
  add column if not exists archived_by text references public.profiles(user_id) on delete set null;

create table if not exists public.mdt_tasks (
  task_id text primary key default ('TASK_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  details text not null default '',
  category text not null default 'General',
  priority text not null default 'Normal' check (priority in ('Critical', 'High', 'Normal', 'Low')),
  status text not null default 'Open' check (status in ('Open', 'In Progress', 'Blocked', 'Completed', 'Cancelled')),
  assigned_officer_id text references public.officers(officer_id) on delete cascade,
  source_type text not null default '',
  source_id text not null default '',
  due_at timestamptz,
  completed_at timestamptz,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_mdt_tasks_assignee on public.mdt_tasks(assigned_officer_id, status, due_at);
create index if not exists idx_mdt_tasks_source on public.mdt_tasks(source_type, source_id);

alter table public.mdt_tasks enable row level security;

drop policy if exists "mdt tasks read assigned or supervisors" on public.mdt_tasks;
create policy "mdt tasks read assigned or supervisors" on public.mdt_tasks for select
using (
  public.has_permission('VIEW_TASKS')
  or exists (
    select 1 from public.officers officer
    where officer.officer_id = mdt_tasks.assigned_officer_id
      and officer.member_id = public.current_member_id()
  )
);

drop policy if exists "mdt tasks create supervisors" on public.mdt_tasks;
create policy "mdt tasks create supervisors" on public.mdt_tasks for insert
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "mdt tasks update assigned or supervisors" on public.mdt_tasks;
create policy "mdt tasks update assigned or supervisors" on public.mdt_tasks for update
using (
  public.has_permission('VIEW_TASKS')
  or exists (
    select 1 from public.officers officer
    where officer.officer_id = mdt_tasks.assigned_officer_id
      and officer.member_id = public.current_member_id()
  )
)
with check (
  public.has_permission('VIEW_TASKS')
  or exists (
    select 1 from public.officers officer
    where officer.officer_id = mdt_tasks.assigned_officer_id
      and officer.member_id = public.current_member_id()
  )
);

drop policy if exists "mdt tasks delete supervisors" on public.mdt_tasks;
create policy "mdt tasks delete supervisors" on public.mdt_tasks for delete
using (public.has_permission('VIEW_TASKS'));
