create table if not exists public.pinned_officers (
  pin_id text primary key default ('PIN_' || replace(gen_random_uuid()::text, '-', '')),
  user_id text not null references public.profiles(user_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  reason text not null default '',
  created_at timestamptz not null default now(),
  unique (user_id, officer_id)
);

create index if not exists idx_pinned_officers_user on public.pinned_officers(user_id, created_at desc);
create index if not exists idx_pinned_officers_officer on public.pinned_officers(officer_id);

alter table public.pinned_officers enable row level security;

drop policy if exists "pinned officers read own" on public.pinned_officers;
create policy "pinned officers read own" on public.pinned_officers for select
using (user_id = public.current_user_id());

drop policy if exists "pinned officers write own" on public.pinned_officers;
create policy "pinned officers write own" on public.pinned_officers for all
using (user_id = public.current_user_id())
with check (user_id = public.current_user_id());
