create extension if not exists pgcrypto;

create table if not exists public.recruitment_vacancies (
  vacancy_id text primary key default ('VAC_' || replace(gen_random_uuid()::text, '-', '')),
  title text not null,
  team text not null default 'Metropolitan Operations 8',
  location text not null default 'London (Roleplay)',
  summary text not null default '',
  description text not null default '',
  vacancy_type text not null default 'External' check (vacancy_type in ('Internal', 'External', 'Internal and External')),
  status text not null default 'Draft' check (status in ('Draft', 'Open', 'Closed', 'Archived')),
  opens_at timestamptz,
  closes_at timestamptz,
  positions integer not null default 1 check (positions > 0),
  reviewer_min_rank text,
  reviewer_required_tags text[] not null default '{}',
  reviewer_officer_ids text[] not null default '{}',
  reviewer_match text not null default 'Any' check (reviewer_match in ('Any', 'All')),
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.recruitment_vacancy_fields (
  field_id text primary key default ('FLD_' || replace(gen_random_uuid()::text, '-', '')),
  vacancy_id text not null references public.recruitment_vacancies(vacancy_id) on delete cascade,
  field_key text not null,
  label text not null,
  help_text text not null default '',
  field_type text not null default 'Long text' check (field_type in ('Short text', 'Long text', 'Yes / No', 'Single choice', 'Multiple choice', 'Date', 'Number')),
  required boolean not null default false,
  options text[] not null default '{}',
  sort_order integer not null default 0,
  unique (vacancy_id, field_key)
);

create table if not exists public.recruitment_accounts (
  account_id uuid primary key default gen_random_uuid(),
  roblox_username text not null,
  username_normalised text generated always as (lower(trim(roblox_username))) stored unique,
  discord_id text not null,
  password_hash text not null,
  status text not null default 'Active' check (status in ('Active', 'Locked', 'Archived')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.recruitment_sessions (
  session_token uuid primary key default gen_random_uuid(),
  account_id uuid not null references public.recruitment_accounts(account_id) on delete cascade,
  expires_at timestamptz not null default (now() + interval '30 days'),
  created_at timestamptz not null default now()
);

create table if not exists public.recruitment_applications (
  application_id text primary key default ('APP_' || replace(gen_random_uuid()::text, '-', '')),
  vacancy_id text not null references public.recruitment_vacancies(vacancy_id) on delete cascade,
  applicant_account_id uuid references public.recruitment_accounts(account_id) on delete set null,
  internal_member_id text,
  roblox_username text not null,
  username_normalised text generated always as (lower(trim(roblox_username))) stored,
  discord_id text not null default '',
  answers jsonb not null default '{}'::jsonb,
  status text not null default 'Submitted' check (status in ('Submitted', 'Under Review', 'Shortlisted', 'Interview', 'Successful', 'Unsuccessful', 'Withdrawn')),
  applicant_message text not null default '',
  private_review_notes text not null default '',
  reviewer_user_id text references public.profiles(user_id) on delete set null,
  submitted_at timestamptz not null default now(),
  reviewed_at timestamptz,
  updated_at timestamptz not null default now()
);

create unique index if not exists recruitment_one_application_per_username
  on public.recruitment_applications(vacancy_id, username_normalised)
  where status <> 'Withdrawn';
create index if not exists recruitment_applications_status on public.recruitment_applications(vacancy_id, status, submitted_at desc);
create index if not exists recruitment_fields_order on public.recruitment_vacancy_fields(vacancy_id, sort_order);

create or replace function public.recruitment_rank_index(rank_name text)
returns integer language sql immutable as $$
  select coalesce(array_position(array[
    'Police Constable','Sergeant','Inspector','Chief Inspector','Superintendent',
    'Chief Superintendent','Commander','Deputy Assistant Commissioner',
    'Assistant Commissioner','Deputy Commissioner','Commissioner'
  ], rank_name), 0)
$$;

create or replace function public.can_review_recruitment_vacancy(target_vacancy_id text)
returns boolean language plpgsql security definer set search_path = public as $$
declare
  vacancy public.recruitment_vacancies;
  officer public.officers;
  profile public.profiles;
begin
  select * into vacancy from public.recruitment_vacancies where vacancy_id = target_vacancy_id;
  select * into profile from public.profiles where auth_user_id = auth.uid() limit 1;
  if vacancy is null or profile is null then return false; end if;
  if vacancy.created_by = profile.user_id then return true; end if;
  select * into officer from public.officers where member_id = profile.member_id limit 1;
  if officer is null then return false; end if;
  return officer.officer_id = any(coalesce(vacancy.reviewer_officer_ids, '{}'));
end;
$$;

create or replace function public.public_recruitment_vacancies()
returns table (vacancy_id text, title text, team text, location text, summary text, description text, vacancy_type text, opens_at timestamptz, closes_at timestamptz, positions integer)
language sql security definer set search_path = public as $$
  select v.vacancy_id, v.title, v.team, v.location, v.summary, v.description, v.vacancy_type, v.opens_at, v.closes_at, v.positions
  from public.recruitment_vacancies v
  where v.status = 'Open'
    and v.vacancy_type in ('External', 'Internal and External')
    and (v.opens_at is null or v.opens_at <= now())
    and (v.closes_at is null or v.closes_at >= now())
  order by v.closes_at nulls last, v.created_at desc
$$;

create or replace function public.public_recruitment_vacancy(target_vacancy_id text)
returns jsonb language sql security definer set search_path = public as $$
  select jsonb_build_object(
    'vacancy', jsonb_build_object(
      'vacancy_id', v.vacancy_id, 'title', v.title, 'team', v.team, 'location', v.location,
      'summary', v.summary, 'description', v.description, 'vacancy_type', v.vacancy_type,
      'opens_at', v.opens_at, 'closes_at', v.closes_at, 'positions', v.positions
    ),
    'fields', coalesce((select jsonb_agg(to_jsonb(f) order by f.sort_order, f.label) from public.recruitment_vacancy_fields f where f.vacancy_id = v.vacancy_id), '[]'::jsonb)
  )
  from public.recruitment_vacancies v
  where v.vacancy_id = target_vacancy_id
    and v.status = 'Open'
    and v.vacancy_type in ('External', 'Internal and External')
    and (v.opens_at is null or v.opens_at <= now())
    and (v.closes_at is null or v.closes_at >= now())
$$;

create or replace function public.recruitment_register(applicant_username text, applicant_discord_id text, applicant_password text)
returns jsonb language plpgsql security definer set search_path = public, extensions as $$
declare account public.recruitment_accounts; token uuid;
begin
  if length(trim(applicant_username)) < 3 or length(trim(applicant_username)) > 40 then raise exception 'Enter a valid Roblox username.'; end if;
  if applicant_discord_id !~ '^\d{15,22}$' then raise exception 'Enter a valid Discord user ID.'; end if;
  if length(applicant_password) < 8 then raise exception 'Password must be at least 8 characters.'; end if;
  insert into public.recruitment_accounts(roblox_username, discord_id, password_hash)
  values (trim(applicant_username), applicant_discord_id, crypt(applicant_password, gen_salt('bf')))
  returning * into account;
  insert into public.recruitment_sessions(account_id) values (account.account_id) returning session_token into token;
  return jsonb_build_object('token', token, 'username', account.roblox_username, 'discordId', account.discord_id);
exception when unique_violation then raise exception 'A recruitment account already exists for that Roblox username.';
end;
$$;

create or replace function public.recruitment_login(applicant_username text, applicant_password text)
returns jsonb language plpgsql security definer set search_path = public, extensions as $$
declare account public.recruitment_accounts; token uuid;
begin
  select * into account from public.recruitment_accounts where username_normalised = lower(trim(applicant_username)) and status = 'Active';
  if account is null or account.password_hash <> crypt(applicant_password, account.password_hash) then raise exception 'Username or password is incorrect.'; end if;
  delete from public.recruitment_sessions where expires_at < now();
  insert into public.recruitment_sessions(account_id) values (account.account_id) returning session_token into token;
  return jsonb_build_object('token', token, 'username', account.roblox_username, 'discordId', account.discord_id);
end;
$$;

create or replace function public.recruitment_session(session_token_input uuid)
returns jsonb language sql security definer set search_path = public as $$
  select jsonb_build_object('username', a.roblox_username, 'discordId', a.discord_id)
  from public.recruitment_sessions s join public.recruitment_accounts a on a.account_id = s.account_id
  where s.session_token = session_token_input and s.expires_at > now() and a.status = 'Active'
$$;

create or replace function public.submit_recruitment_application(session_token_input uuid, target_vacancy_id text, application_answers jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare account public.recruitment_accounts; vacancy public.recruitment_vacancies; required_field public.recruitment_vacancy_fields; application_id_output text;
begin
  select a.* into account from public.recruitment_sessions s join public.recruitment_accounts a on a.account_id = s.account_id where s.session_token = session_token_input and s.expires_at > now() and a.status = 'Active';
  if account is null then raise exception 'Your recruitment session has expired. Please sign in again.'; end if;
  select * into vacancy from public.recruitment_vacancies where vacancy_id = target_vacancy_id and status = 'Open' and vacancy_type in ('External', 'Internal and External') and (opens_at is null or opens_at <= now()) and (closes_at is null or closes_at >= now());
  if vacancy is null then raise exception 'This vacancy is not accepting external applications.'; end if;
  for required_field in select * from public.recruitment_vacancy_fields where vacancy_id = target_vacancy_id and required loop
    if not application_answers ? required_field.field_key or nullif(trim(application_answers ->> required_field.field_key), '') is null then raise exception 'Complete the required field: %', required_field.label; end if;
  end loop;
  insert into public.recruitment_applications(vacancy_id, applicant_account_id, roblox_username, discord_id, answers)
  values (target_vacancy_id, account.account_id, account.roblox_username, account.discord_id, coalesce(application_answers, '{}'::jsonb))
  returning application_id into application_id_output;
  return jsonb_build_object('applicationId', application_id_output, 'status', 'Submitted');
exception when unique_violation then raise exception 'You have already applied for this vacancy.';
end;
$$;

create or replace function public.recruitment_my_applications(session_token_input uuid)
returns table (application_id text, vacancy_id text, vacancy_title text, status text, applicant_message text, submitted_at timestamptz, updated_at timestamptz)
language sql security definer set search_path = public as $$
  select a.application_id, a.vacancy_id, v.title, a.status, a.applicant_message, a.submitted_at, a.updated_at
  from public.recruitment_sessions s
  join public.recruitment_applications a on a.applicant_account_id = s.account_id
  join public.recruitment_vacancies v on v.vacancy_id = a.vacancy_id
  where s.session_token = session_token_input and s.expires_at > now()
  order by a.submitted_at desc
$$;

create or replace function public.submit_internal_recruitment_application(target_vacancy_id text, application_answers jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare profile public.profiles; vacancy public.recruitment_vacancies; required_field public.recruitment_vacancy_fields; application_id_output text;
begin
  select * into profile from public.profiles where auth_user_id = auth.uid() limit 1;
  if profile is null then raise exception 'No MDT profile is linked to this login.'; end if;
  select * into vacancy from public.recruitment_vacancies where vacancy_id = target_vacancy_id and status = 'Open' and vacancy_type in ('Internal', 'Internal and External') and (opens_at is null or opens_at <= now()) and (closes_at is null or closes_at >= now());
  if vacancy is null then raise exception 'This vacancy is not accepting internal applications.'; end if;
  for required_field in select * from public.recruitment_vacancy_fields where vacancy_id = target_vacancy_id and required loop
    if not application_answers ? required_field.field_key or nullif(trim(application_answers ->> required_field.field_key), '') is null then raise exception 'Complete the required field: %', required_field.label; end if;
  end loop;
  insert into public.recruitment_applications(vacancy_id, internal_member_id, roblox_username, discord_id, answers)
  values (target_vacancy_id, profile.member_id, profile.roblox_username, profile.discord_id, coalesce(application_answers, '{}'::jsonb))
  returning application_id into application_id_output;
  return jsonb_build_object('applicationId', application_id_output, 'status', 'Submitted');
exception when unique_violation then raise exception 'You have already applied for this vacancy.';
end;
$$;

alter table public.recruitment_vacancies enable row level security;
alter table public.recruitment_vacancy_fields enable row level security;
alter table public.recruitment_accounts enable row level security;
alter table public.recruitment_sessions enable row level security;
alter table public.recruitment_applications enable row level security;

drop policy if exists "recruitment vacancies visible" on public.recruitment_vacancies;
create policy "recruitment vacancies visible" on public.recruitment_vacancies for select using (
  (status = 'Open' and auth.role() = 'authenticated')
  or public.can_review_recruitment_vacancy(vacancy_id)
);
drop policy if exists "recruitment vacancies managed" on public.recruitment_vacancies;
create policy "recruitment vacancies managed" on public.recruitment_vacancies for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "recruitment fields visible" on public.recruitment_vacancy_fields;
create policy "recruitment fields visible" on public.recruitment_vacancy_fields for select using (exists (select 1 from public.recruitment_vacancies v where v.vacancy_id = recruitment_vacancy_fields.vacancy_id));
drop policy if exists "recruitment fields managed" on public.recruitment_vacancy_fields;
create policy "recruitment fields managed" on public.recruitment_vacancy_fields for all using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "recruitment applications visible" on public.recruitment_applications;
create policy "recruitment applications visible" on public.recruitment_applications for select using (
  internal_member_id = public.current_member_id() or public.can_review_recruitment_vacancy(vacancy_id)
);
drop policy if exists "recruitment applications reviewed" on public.recruitment_applications;
create policy "recruitment applications reviewed" on public.recruitment_applications for update using (
  public.can_review_recruitment_vacancy(vacancy_id) and coalesce(internal_member_id, '') <> coalesce(public.current_member_id(), '')
) with check (
  public.can_review_recruitment_vacancy(vacancy_id) and coalesce(internal_member_id, '') <> coalesce(public.current_member_id(), '')
);

revoke all on function public.public_recruitment_vacancies() from public;
revoke all on function public.public_recruitment_vacancy(text) from public;
revoke all on function public.recruitment_register(text, text, text) from public;
revoke all on function public.recruitment_login(text, text) from public;
revoke all on function public.recruitment_session(uuid) from public;
revoke all on function public.submit_recruitment_application(uuid, text, jsonb) from public;
revoke all on function public.recruitment_my_applications(uuid) from public;
revoke all on function public.submit_internal_recruitment_application(text, jsonb) from public;
revoke all on function public.can_review_recruitment_vacancy(text) from public;
grant select, insert, update, delete on public.recruitment_vacancies to authenticated;
grant select, insert, update, delete on public.recruitment_vacancy_fields to authenticated;
grant select, update on public.recruitment_applications to authenticated;
grant execute on function public.public_recruitment_vacancies() to anon, authenticated;
grant execute on function public.public_recruitment_vacancy(text) to anon, authenticated;
grant execute on function public.recruitment_register(text, text, text) to anon, authenticated;
grant execute on function public.recruitment_login(text, text) to anon, authenticated;
grant execute on function public.recruitment_session(uuid) to anon, authenticated;
grant execute on function public.submit_recruitment_application(uuid, text, jsonb) to anon, authenticated;
grant execute on function public.recruitment_my_applications(uuid) to anon, authenticated;
grant execute on function public.submit_internal_recruitment_application(text, jsonb) to authenticated;
grant execute on function public.can_review_recruitment_vacancy(text) to authenticated;
