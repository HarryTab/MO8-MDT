begin;

alter table public.account_requests
  add column if not exists callsign text not null default '';

alter table public.training_courses
  add column if not exists duration_minutes integer not null default 60;

update public.training_courses
set duration_minutes = 60
where duration_minutes is null or duration_minutes < 1;

alter table public.training_courses
  drop constraint if exists training_courses_duration_minutes_check;
alter table public.training_courses
  add constraint training_courses_duration_minutes_check
  check (duration_minutes between 15 and 1440);

commit;
