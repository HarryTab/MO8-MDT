create table if not exists public.operational_powers (
  power_id text primary key default ('PWR_' || replace(gen_random_uuid()::text, '-', '')),
  code text not null unique,
  title text not null,
  category text not null default 'General',
  definition text not null default '',
  legal_test text not null default '',
  required_record_fields text[] not null default array[]::text[],
  legislation_source text not null default 'Police and Criminal Evidence Act 2025',
  section_reference text not null default '',
  requires_supervisor_review boolean not null default false,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.operational_power_uses (
  power_use_id text primary key default ('PUSE_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  power_id text not null references public.operational_powers(power_id) on delete restrict,
  person_id text references public.operational_persons(person_id) on delete set null,
  vehicle_id text references public.operational_vehicles(vehicle_id) on delete set null,
  grounds text not null,
  necessity text not null default '',
  outcome text not null,
  items_seized text not null default '',
  duration_minutes integer not null default 0,
  occurred_at timestamptz not null default now(),
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now(),
  check (person_id is not null or vehicle_id is not null)
);

create table if not exists public.operational_action_reviews (
  action_review_id text primary key default ('AREV_' || replace(gen_random_uuid()::text, '-', '')),
  incident_id text not null references public.operational_incidents(incident_id) on delete cascade,
  officer_id text not null references public.officers(officer_id) on delete cascade,
  action_type text not null,
  source_type text not null check (source_type in ('Disposal', 'Officer Action', 'Power Use')),
  source_id text not null,
  assigned_reviewer_user_id text references public.profiles(user_id) on delete set null,
  rationale text not null,
  outcome_report text not null,
  status text not null default 'Pending' check (status in ('Pending', 'Approved', 'Amendments Required', 'Completed', 'Cancelled')),
  supervisor_feedback text not null default '',
  due_at timestamptz not null default (now() + interval '7 days'),
  reviewed_by text references public.profiles(user_id) on delete set null,
  reviewed_at timestamptz,
  created_at timestamptz not null default now(),
  unique (source_type, source_id)
);

create table if not exists public.operational_controller_sessions (
  controller_session_id text primary key default ('CTRL_' || replace(gen_random_uuid()::text, '-', '')),
  officer_id text not null references public.officers(officer_id) on delete cascade,
  controller_role text not null default 'Controller',
  capabilities text[] not null default array[]::text[],
  status text not null default 'Active' check (status in ('Active', 'Break', 'Ended')),
  notes text not null default '',
  started_at timestamptz not null default now(),
  ended_at timestamptz,
  created_by text references public.profiles(user_id) on delete set null
);

create table if not exists public.operational_callsign_presets (
  preset_id text primary key default ('CSP_' || replace(gen_random_uuid()::text, '-', '')),
  callsign text not null unique,
  division text not null default 'Roads Policing Unit',
  unit_role text not null default 'Traffic Vehicle',
  vehicle_requirement text not null default 'Any Traffic Vehicle',
  required_capabilities text[] not null default array[]::text[],
  notes text not null default '',
  sort_order integer not null default 0,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.operational_units
  add column if not exists preset_id text references public.operational_callsign_presets(preset_id) on delete set null;

create index if not exists idx_operational_power_uses_subject_person on public.operational_power_uses(person_id, occurred_at desc);
create index if not exists idx_operational_power_uses_subject_vehicle on public.operational_power_uses(vehicle_id, occurred_at desc);
create index if not exists idx_operational_power_uses_incident on public.operational_power_uses(incident_id, occurred_at desc);
create index if not exists idx_operational_action_reviews_assignee on public.operational_action_reviews(assigned_reviewer_user_id, status, due_at);
create index if not exists idx_operational_controller_sessions_active on public.operational_controller_sessions(status, started_at desc);
create unique index if not exists idx_operational_controller_one_active on public.operational_controller_sessions(officer_id) where status in ('Active', 'Break');
create unique index if not exists idx_operational_units_active_preset on public.operational_units(preset_id) where preset_id is not null and closed_at is null;

alter table public.operational_powers enable row level security;
alter table public.operational_power_uses enable row level security;
alter table public.operational_action_reviews enable row level security;
alter table public.operational_controller_sessions enable row level security;
alter table public.operational_callsign_presets enable row level security;

drop policy if exists "operational powers read authenticated" on public.operational_powers;
create policy "operational powers read authenticated" on public.operational_powers for select using (auth.role() = 'authenticated');
drop policy if exists "operational powers manage command" on public.operational_powers;
create policy "operational powers manage command" on public.operational_powers for all using (public.has_permission('FULL_ACCESS')) with check (public.has_permission('FULL_ACCESS'));
drop policy if exists "operational power uses authenticated" on public.operational_power_uses;
create policy "operational power uses authenticated" on public.operational_power_uses for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational action reviews read supervisors" on public.operational_action_reviews;
create policy "operational action reviews read supervisors" on public.operational_action_reviews for select using (public.has_permission('VIEW_TASKS'));
drop policy if exists "operational action reviews create authenticated" on public.operational_action_reviews;
create policy "operational action reviews create authenticated" on public.operational_action_reviews for insert with check (auth.role() = 'authenticated');
drop policy if exists "operational action reviews update supervisors" on public.operational_action_reviews;
create policy "operational action reviews update supervisors" on public.operational_action_reviews for update using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));
drop policy if exists "operational action reviews delete supervisors" on public.operational_action_reviews;
create policy "operational action reviews delete supervisors" on public.operational_action_reviews for delete using (public.has_permission('VIEW_TASKS'));
drop policy if exists "operational controller sessions authenticated" on public.operational_controller_sessions;
create policy "operational controller sessions authenticated" on public.operational_controller_sessions for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
drop policy if exists "operational callsign presets read authenticated" on public.operational_callsign_presets;
create policy "operational callsign presets read authenticated" on public.operational_callsign_presets for select using (auth.role() = 'authenticated');
drop policy if exists "operational callsign presets manage command" on public.operational_callsign_presets;
create policy "operational callsign presets manage command" on public.operational_callsign_presets for all using (public.has_permission('FULL_ACCESS')) with check (public.has_permission('FULL_ACCESS'));

insert into public.operational_powers (
  code, title, category, definition, legal_test, required_record_fields,
  section_reference, requires_supervisor_review, updated_at
) values
  ('PACE-S2', 'Stop and search', 'Search', 'Detain and search a person, vehicle, or anything in or on a vehicle for stolen or prohibited articles in an accessible public place.', 'The officer must have objective reasonable grounds to suspect that stolen or prohibited articles will be found. A dwelling is excluded.', array['Grounds', 'Object of search', 'Outcome', 'Items seized'], 'Section 2', true, now()),
  ('PACE-S2A', 'Superintendent-authorised search', 'Search', 'Search a person or vehicle without individual reasonable grounds under a valid superintendent authorisation.', 'A valid authorisation must cover the location and time and relate to anticipated serious violence, weapons, or an area associated with criminal activity. Maximum duration is 24 hours.', array['Authorising officer', 'Authorisation area', 'Authorisation expiry', 'Outcome'], 'Section 2(2)(a)', true, now()),
  ('PACE-S2B', 'Superintendent-authorised area search', 'Search', 'Conduct searches within a defined area under a mass-search authority.', 'A superintendent or above must authorise the area because serious violence may occur or the area is associated with a high level of criminal activity. Maximum duration is 24 hours.', array['Authorising officer', 'Authorisation area', 'Authorisation expiry', 'Outcome'], 'Section 2(2)(b)', true, now()),
  ('PACE-S4', 'Entry and search without warrant', 'Entry and Seizure', 'Enter premises without warrant to execute an arrest warrant, arrest for an indictable offence, save life, or prevent serious property damage.', 'The officer must reasonably suspect the sought person is on the premises. Items may only be seized where lawfully present and necessary to preserve evidence.', array['Purpose of entry', 'Grounds', 'Premises', 'Items seized', 'Outcome'], 'Section 4', true, now()),
  ('PACE-S5', 'Arrest without warrant', 'Arrest', 'Arrest a person reasonably suspected of committing, having committed, or being about to commit an offence.', 'Arrest must be necessary to ascertain identity, prevent injury or property damage, enable prompt and effective investigation, or prevent disappearance.', array['Suspected offence', 'Grounds', 'Necessity', 'Information given', 'Outcome'], 'Section 5', true, now()),
  ('PACE-S10', 'Search upon arrest', 'Search', 'Search an arrested person, and in qualifying cases the immediate premises, for danger, escape items, evidence, or property obtained through offending.', 'The person must be lawfully arrested. The officer must record reasonable grounds and limit the search to what is reasonably required.', array['Arrest offence', 'Grounds', 'Object of search', 'Items seized', 'Outcome'], 'Section 10', true, now()),
  ('PACE-S16', 'Dispersal direction', 'Public Order', 'Direct a person to leave a locality and not return for the specified exclusion period.', 'A listed public-safety, public-order, highway-obstruction, violent-protest, or imminent-lawlessness condition must apply. An ordinary direction lasts no more than 20 minutes without renewal.', array['Condition relied upon', 'Locality', 'Exclusion period', 'Direction given', 'Outcome'], 'Section 16', false, now()),
  ('PACE-S17', 'Reasonable force when exercising a power', 'Use of Force', 'Use reasonable force where necessary to exercise a statutory power that does not require consent.', 'Force must be necessary, reasonable, and connected to a lawful power.', array['Underlying power', 'Necessity', 'Force used', 'Outcome'], 'Section 17', true, now()),
  ('PACE-S18', 'Reasonable force in arrest or crime prevention', 'Use of Force', 'Use reasonable force to prevent crime or effect or assist a lawful arrest.', 'Force must be reasonable in the circumstances and connected to crime prevention or lawful arrest.', array['Underlying offence', 'Necessity', 'Force used', 'Outcome'], 'Section 18', true, now())
on conflict (code) do update set
  title = excluded.title,
  category = excluded.category,
  definition = excluded.definition,
  legal_test = excluded.legal_test,
  required_record_fields = excluded.required_record_fields,
  section_reference = excluded.section_reference,
  requires_supervisor_review = excluded.requires_supervisor_review,
  active = true,
  updated_at = now();

insert into public.operational_callsign_presets (
  callsign, division, unit_role, vehicle_requirement, required_capabilities,
  notes, sort_order, updated_at
) values
  ('XG-01', 'Duty Command', 'Duty Gold', 'Any Service Vehicle', array['Gold Command'], 'Strategic command callsign.', 10, now()),
  ('SO-01', 'Duty Command', 'Duty Silver', 'Any Service Vehicle', array['Silver Command'], 'Tactical command callsign.', 20, now()),
  ('VB-01', 'Duty Command', 'Duty Inspector', 'Any Service Vehicle', array['Inspector'], 'Duty Inspector callsign.', 30, now()),
  ('VB-02', 'Duty Command', 'Duty Sergeant', 'Any Service Vehicle', array['Supervisor'], 'Duty Sergeant callsign.', 40, now()),
  ('OSCAR-1', 'Duty Command', 'Force Incident Manager', 'Control Room', array['Controller'], 'Force incident management.', 50, now()),
  ('VICTOR-BRAVO', 'Duty Command', 'Control', 'Control Room', array['Controller'], 'Primary control-room callsign.', 60, now()),
  ('RP-01', 'Roads Policing Unit', 'Traffic Supervisor', 'Any Traffic Vehicle', array['Advanced + TPAC', 'Supervisor'], 'Roads supervisor unit.', 100, now()),
  ('RP-20', 'Roads Policing Unit', 'Traffic Vehicle (TPAC)', 'Any Traffic Vehicle', array['Advanced + TPAC'], 'TPAC-capable traffic unit.', 110, now()),
  ('RP-23', 'Roads Policing Unit', 'Traffic Vehicle (TPAC)', 'Any Traffic Vehicle', array['Advanced + TPAC'], 'TPAC-capable traffic unit.', 120, now()),
  ('RP-28', 'Roads Policing Unit', 'Traffic Vehicle', 'Any Traffic Vehicle', array['Advanced'], 'General traffic unit.', 130, now()),
  ('RP-29', 'Roads Policing Unit', 'Traffic Vehicle', 'Any Traffic Vehicle', array['Advanced'], 'General traffic unit.', 140, now()),
  ('RP-51', 'Roads Policing Unit', 'Proactive Unit', 'Any Unmarked Traffic Vehicle', array['Advanced', 'Roads Crime Team'], 'Proactive roads-crime unit.', 150, now())
on conflict (callsign) do update set
  division = excluded.division,
  unit_role = excluded.unit_role,
  vehicle_requirement = excluded.vehicle_requirement,
  required_capabilities = excluded.required_capabilities,
  notes = excluded.notes,
  sort_order = excluded.sort_order,
  active = true,
  updated_at = now();
