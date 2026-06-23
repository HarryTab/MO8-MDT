begin;

create table if not exists public.account_requests (
  request_id text primary key default ('ACR-' || upper(substr(md5(random()::text || clock_timestamp()::text), 1, 10))),
  roblox_username text not null,
  rank text not null,
  discord_id text not null,
  status text not null default 'Pending' check (status in ('Pending', 'Approved', 'Denied')),
  review_notes text not null default '',
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create unique index if not exists account_requests_pending_username_idx
  on public.account_requests (lower(roblox_username)) where status = 'Pending';

alter table public.account_requests enable row level security;
drop policy if exists "Public may request an account" on public.account_requests;
create policy "Public may request an account" on public.account_requests
  for insert to anon, authenticated with check (
    status = 'Pending' and review_notes = '' and reviewed_by is null and reviewed_at is null
  );
drop policy if exists "Account reviewers may read requests" on public.account_requests;
create policy "Account reviewers may read requests" on public.account_requests
  for select to authenticated using (public.has_permission('REVIEW_ACCOUNT_REQUESTS'));
drop policy if exists "Account reviewers may update requests" on public.account_requests;
create policy "Account reviewers may update requests" on public.account_requests
  for update to authenticated using (public.has_permission('REVIEW_ACCOUNT_REQUESTS'))
  with check (public.has_permission('REVIEW_ACCOUNT_REQUESTS'));

insert into public.permissions (role, permission, allowed) values
  ('Inspector', 'REVIEW_ACCOUNT_REQUESTS', true),
  ('Chief Inspector', 'REVIEW_ACCOUNT_REQUESTS', true),
  ('Command', 'REVIEW_ACCOUNT_REQUESTS', true)
on conflict (role, permission) do update set allowed = excluded.allowed;

alter table public.handover_entries
  add column if not exists recipient_user_ids text[] not null default '{}';

create table if not exists public.retrospective_shift_requests (
  request_id text primary key default ('RSR-' || upper(substr(md5(random()::text || clock_timestamp()::text), 1, 10))),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  started_at timestamptz not null,
  ended_at timestamptz not null,
  summary text not null default '',
  reason text not null,
  status text not null default 'Pending' check (status in ('Pending', 'Approved', 'Denied')),
  review_reason text not null default '',
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  created_at timestamptz not null default now()
);

alter table public.retrospective_shift_requests enable row level security;
drop policy if exists "Officers may create retrospective shift requests" on public.retrospective_shift_requests;
create policy "Officers may create retrospective shift requests" on public.retrospective_shift_requests
  for insert to authenticated with check (
    officer_id in (
      select o.officer_id from public.officers o
      join public.profiles p on p.member_id = o.member_id
      where p.auth_user_id = auth.uid()
    )
  );
drop policy if exists "Officers and task managers may view retrospective requests" on public.retrospective_shift_requests;
create policy "Officers and task managers may view retrospective requests" on public.retrospective_shift_requests
  for select to authenticated using (
    public.has_permission('VIEW_TASKS') or officer_id in (
      select o.officer_id from public.officers o
      join public.profiles p on p.member_id = o.member_id
      where p.auth_user_id = auth.uid()
    )
  );
drop policy if exists "Task managers may review retrospective requests" on public.retrospective_shift_requests;
create policy "Task managers may review retrospective requests" on public.retrospective_shift_requests
  for update to authenticated using (public.has_permission('VIEW_TASKS'))
  with check (public.has_permission('VIEW_TASKS'));

commit;
