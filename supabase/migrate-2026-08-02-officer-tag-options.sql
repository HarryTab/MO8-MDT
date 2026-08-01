create table if not exists public.officer_tag_options (
  tag_id text primary key default ('TAG_' || replace(gen_random_uuid()::text, '-', '')),
  name text not null unique,
  status text not null default 'Active',
  sort_order integer not null default 0,
  description text,
  updated_by text references public.profiles(user_id) on delete set null,
  updated_at timestamptz not null default now()
);

alter table public.officer_tag_options enable row level security;

drop policy if exists "officer tag options read" on public.officer_tag_options;
create policy "officer tag options read" on public.officer_tag_options for select
using (public.current_user_id() is not null);

drop policy if exists "officer tag options write" on public.officer_tag_options;
create policy "officer tag options write" on public.officer_tag_options for all
using (public.has_permission('MANAGE_OFFICER_TAGS'))
with check (public.has_permission('MANAGE_OFFICER_TAGS'));

insert into public.officer_tag_options (name, sort_order) values
  ('Roads Crime Team', 10),
  ('MO8 Command', 20),
  ('Roads and Traffic Policing Team', 30),
  ('Bronze Command', 40),
  ('Silver Command', 50),
  ('Gold Command', 60),
  ('Controller', 70),
  ('Control Supervisor', 80),
  ('Tactical Advisor', 90)
on conflict (name) do nothing;

insert into public.permissions (role, permission, allowed) values
  ('Inspector', 'MANAGE_OFFICER_TAGS', true),
  ('Chief Inspector', 'MANAGE_OFFICER_TAGS', true),
  ('Command', 'MANAGE_OFFICER_TAGS', true)
on conflict (role, permission) do update set allowed = excluded.allowed;

