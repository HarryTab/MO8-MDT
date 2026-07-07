-- Shift duty workflow, supervisor assurance and immutable change history.
alter table public.shift_logs
  add column if not exists objectives text not null default '',
  add column if not exists start_notes text not null default '',
  add column if not exists voided_at timestamptz,
  add column if not exists voided_by text references public.profiles(user_id) on delete set null,
  add column if not exists void_reason text not null default '';

alter table public.shift_wrapups
  add column if not exists quality_score integer,
  add column if not exists review_status text not null default 'Pending',
  add column if not exists review_notes text not null default '',
  add column if not exists reviewed_by text references public.profiles(user_id) on delete set null,
  add column if not exists reviewed_at timestamptz;

create table if not exists public.shift_audit_events (
  event_id text primary key default ('SHFA_' || replace(gen_random_uuid()::text, '-', '')),
  shift_id text not null references public.shift_logs(shift_id) on delete restrict,
  officer_id text not null references public.officers(officer_id) on delete restrict,
  action text not null check (action in ('Created', 'Ended', 'Edited', 'Reviewed', 'Voided', 'Restored')),
  reason text not null default '',
  before_snapshot jsonb not null default '{}'::jsonb,
  after_snapshot jsonb not null default '{}'::jsonb,
  performed_by text references public.profiles(user_id) on delete set null,
  performed_at timestamptz not null default now()
);

create index if not exists idx_shift_audit_shift on public.shift_audit_events(shift_id, performed_at desc);
create index if not exists idx_shift_audit_officer on public.shift_audit_events(officer_id, performed_at desc);
alter table public.shift_audit_events enable row level security;

drop policy if exists "shift audit read authorised" on public.shift_audit_events;
create policy "shift audit read authorised" on public.shift_audit_events for select using (
  public.has_permission('VIEW_TASKS') or exists (
    select 1 from public.officers o where o.officer_id = shift_audit_events.officer_id and o.member_id = public.current_member_id()
  )
);
drop policy if exists "shift audit write supervisors" on public.shift_audit_events;
create policy "shift audit write supervisors" on public.shift_audit_events for insert with check (
  auth.role() = 'authenticated'
);

drop policy if exists "shift wrapups update supervisors" on public.shift_wrapups;
create policy "shift wrapups update supervisors" on public.shift_wrapups for update
using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

-- Existing installations differ slightly in their shift_logs policy names. These
-- explicit policies permit officers to manage their live shift and supervisors to
-- perform recorded assurance actions through the MDT.
drop policy if exists "shift logs update own or supervisors" on public.shift_logs;
create policy "shift logs update own or supervisors" on public.shift_logs for update using (
  public.has_permission('VIEW_TASKS') or exists (
    select 1 from public.officers o where o.officer_id = shift_logs.officer_id and o.member_id = public.current_member_id()
  )
) with check (
  public.has_permission('VIEW_TASKS') or exists (
    select 1 from public.officers o where o.officer_id = shift_logs.officer_id and o.member_id = public.current_member_id()
  )
);

drop policy if exists "mdt tasks create supervisors" on public.mdt_tasks;
create policy "mdt tasks create supervisors" on public.mdt_tasks for insert with check (
  public.has_permission('VIEW_TASKS')
  or (
    source_type = 'Shift Debrief'
    and created_by = auth.uid()::text
    and exists (
      select 1 from public.shift_logs s
      join public.officers o on o.officer_id = s.officer_id
      where s.shift_id = mdt_tasks.source_id and o.member_id = public.current_member_id()
    )
  )
);
