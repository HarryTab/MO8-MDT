alter table public.calendar_events
  add column if not exists requires_rsvp boolean not null default true,
  add column if not exists reminder_minutes integer[] not null default array[1440, 60]::integer[];

create table if not exists public.calendar_rsvps (
  rsvp_id text primary key default ('RSVP_' || replace(gen_random_uuid()::text, '-', '')),
  event_id text not null references public.calendar_events(event_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  response text not null default 'Pending' check (response in ('Pending', 'Attending', 'Maybe', 'Not Attending')),
  note text not null default '',
  responded_at timestamptz,
  updated_at timestamptz not null default now(),
  unique (event_id, officer_id)
);

create table if not exists public.calendar_reminder_log (
  reminder_id text primary key default ('REM_' || replace(gen_random_uuid()::text, '-', '')),
  event_id text not null references public.calendar_events(event_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  minutes_before integer not null,
  sent_at timestamptz not null default now(),
  unique (event_id, officer_id, minutes_before)
);

create table if not exists public.officer_notes (
  note_id text primary key default ('NOTE_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  author_user_id text references public.profiles(user_id) on delete set null,
  visibility text not null default 'Supervisor' check (visibility in ('Visible to officer', 'Supervisor', 'Command', 'Training')),
  title text not null default '',
  body text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_calendar_rsvps_event on public.calendar_rsvps(event_id);
create index if not exists idx_calendar_rsvps_officer on public.calendar_rsvps(officer_id);
create index if not exists idx_calendar_reminders_event on public.calendar_reminder_log(event_id);
create index if not exists idx_officer_notes_officer on public.officer_notes(officer_id, created_at desc);

alter table public.calendar_rsvps enable row level security;
alter table public.calendar_reminder_log enable row level security;
alter table public.officer_notes enable row level security;

drop policy if exists "calendar rsvps read relevant" on public.calendar_rsvps;
create policy "calendar rsvps read relevant" on public.calendar_rsvps for select
using (
  public.has_permission('VIEW_TASKS')
  or officer_id in (select officer_id from public.officers where member_id = public.current_member_id())
);

drop policy if exists "calendar rsvps write own or managers" on public.calendar_rsvps;
create policy "calendar rsvps write own or managers" on public.calendar_rsvps for all
using (
  public.has_permission('VIEW_TASKS')
  or officer_id in (select officer_id from public.officers where member_id = public.current_member_id())
)
with check (
  public.has_permission('VIEW_TASKS')
  or officer_id in (select officer_id from public.officers where member_id = public.current_member_id())
);

drop policy if exists "calendar reminders managers" on public.calendar_reminder_log;
create policy "calendar reminders managers" on public.calendar_reminder_log for all
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "officer notes read scoped" on public.officer_notes;
create policy "officer notes read scoped" on public.officer_notes for select
using (
  public.has_permission('FULL_ACCESS')
  or (visibility = 'Training' and public.has_permission('MANAGE_TRAINING'))
  or (
    visibility in ('Supervisor', 'Training')
    and exists (
      select 1 from public.officers o
      where o.officer_id = officer_notes.officer_id
      and o.supervisor_user_id = public.current_user_id()
    )
  )
  or (
    visibility = 'Visible to officer'
    and officer_id in (select officer_id from public.officers where member_id = public.current_member_id())
  )
);

drop policy if exists "officer notes write managers" on public.officer_notes;
create policy "officer notes write managers" on public.officer_notes for all
using (
  public.has_permission('FULL_ACCESS')
  or public.has_permission('VIEW_TASKS')
  or public.has_permission('MANAGE_TRAINING')
)
with check (
  public.has_permission('FULL_ACCESS')
  or public.has_permission('VIEW_TASKS')
  or public.has_permission('MANAGE_TRAINING')
);
