drop policy if exists "aip manage supervisors" on public.activity_improvement_notices;
drop policy if exists "aip create supervisors" on public.activity_improvement_notices;
drop policy if exists "aip update supervisors" on public.activity_improvement_notices;
drop policy if exists "aip delete full access" on public.activity_improvement_notices;

create policy "aip create supervisors" on public.activity_improvement_notices
for insert with check (public.has_permission('VIEW_TASKS'));

create policy "aip update supervisors" on public.activity_improvement_notices
for update using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

create policy "aip delete full access" on public.activity_improvement_notices
for delete using (public.has_permission('FULL_ACCESS'));

create or replace function public.permanently_delete_archived_aip(target_aip_id text, deletion_reason text)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  aip_record public.activity_improvement_notices%rowtype;
  officer_name text;
begin
  if not public.has_permission('FULL_ACCESS') then
    raise exception 'Full Access permission is required to permanently delete an AIP.';
  end if;

  if nullif(trim(deletion_reason), '') is null then
    raise exception 'A deletion reason is required.';
  end if;

  select * into aip_record
  from public.activity_improvement_notices
  where aip_id = target_aip_id
  for update;

  if not found then raise exception 'AIP not found.'; end if;
  if aip_record.archived_at is null then raise exception 'Only archived AIPs can be permanently deleted.'; end if;

  select roblox_username into officer_name
  from public.officers
  where officer_id = aip_record.officer_id;

  insert into public.audit_log (actor_user_id, action, target_type, target_id, details)
  values (
    public.current_user_id(),
    'PERMANENTLY_DELETE_AIP',
    'Activity Improvement Notice Tombstone',
    aip_record.aip_id,
    jsonb_build_object(
      'reference', aip_record.reference,
      'officer_id', aip_record.officer_id,
      'officer', coalesce(officer_name, ''),
      'status_at_deletion', aip_record.status,
      'archived_at', aip_record.archived_at,
      'archive_reason', aip_record.archive_reason,
      'deletion_reason', trim(deletion_reason),
      'deleted_at', now()
    )
  );

  delete from public.activity_improvement_notices where aip_id = target_aip_id;
end;
$$;

revoke all on function public.permanently_delete_archived_aip(text, text) from public;
grant execute on function public.permanently_delete_archived_aip(text, text) to authenticated;
