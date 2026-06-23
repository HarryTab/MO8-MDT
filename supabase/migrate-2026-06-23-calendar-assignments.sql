begin;

alter table public.calendar_events
  add column if not exists audience_type text not null default 'Everyone',
  add column if not exists audience_values text[] not null default '{}',
  add column if not exists assigned_officer_ids text[] not null default '{}',
  add column if not exists priority text not null default 'Normal';

alter table public.calendar_events
  drop constraint if exists calendar_events_audience_type_check;
alter table public.calendar_events
  add constraint calendar_events_audience_type_check
  check (audience_type in ('Everyone', 'Role', 'Ranks', 'Minimum Rank', 'Tag', 'My Supervisees', 'Specific Officers'));

drop policy if exists "calendar read authenticated" on public.calendar_events;
drop policy if exists "calendar manage supervisors" on public.calendar_events;
drop policy if exists "calendar create supervisors" on public.calendar_events;
drop policy if exists "calendar update owners or command" on public.calendar_events;
drop policy if exists "calendar delete owners or command" on public.calendar_events;

create policy "calendar read authenticated" on public.calendar_events for select
using (
  auth.uid() is not null and (
    public.has_permission('FULL_ACCESS')
    or
    created_by = public.current_user_id()
    or audience_type = 'Everyone'
    or (audience_type = 'Role' and exists (
      select 1 from public.profiles p
      where p.auth_user_id = auth.uid() and p.role = any(audience_values)
    ))
    or (audience_type = 'Ranks' and exists (
      select 1 from public.profiles p
      where p.auth_user_id = auth.uid() and p.rank = any(audience_values)
    ))
    or (audience_type = 'Minimum Rank' and exists (
      select 1 from public.profiles p
      where p.auth_user_id = auth.uid()
        and array_position(array['Police Constable','Sergeant','Inspector','Chief Inspector','Superintendent','Chief Superintendent','Commander','Deputy Assistant Commissioner','Assistant Commissioner','Deputy Commissioner','Commissioner']::text[], p.rank)
          >= array_position(array['Police Constable','Sergeant','Inspector','Chief Inspector','Superintendent','Chief Superintendent','Commander','Deputy Assistant Commissioner','Assistant Commissioner','Deputy Commissioner','Commissioner']::text[], audience_values[1])
    ))
    or (audience_type = 'Tag' and exists (
      select 1 from public.officers o
      where o.member_id = public.current_member_id() and o.tags && audience_values
    ))
    or (audience_type = 'My Supervisees' and exists (
      select 1 from public.officers o
      where o.member_id = public.current_member_id()
        and o.supervisor_user_id = calendar_events.created_by
        and (cardinality(assigned_officer_ids) = 0 or o.officer_id = any(assigned_officer_ids))
    ))
    or (audience_type = 'Specific Officers' and exists (
      select 1 from public.officers o
      where o.member_id = public.current_member_id() and o.officer_id = any(assigned_officer_ids)
    ))
  )
);

create policy "calendar create supervisors" on public.calendar_events for insert
with check (public.has_permission('VIEW_TASKS') and created_by = public.current_user_id());

create policy "calendar update owners or command" on public.calendar_events for update
using (created_by = public.current_user_id() or public.has_permission('FULL_ACCESS'))
with check (created_by = public.current_user_id() or public.has_permission('FULL_ACCESS'));

create policy "calendar delete owners or command" on public.calendar_events for delete
using (created_by = public.current_user_id() or public.has_permission('FULL_ACCESS'));

commit;
