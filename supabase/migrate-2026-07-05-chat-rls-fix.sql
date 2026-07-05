-- Creates conversations and their initial membership atomically.
-- This avoids the RLS gap between inserting a conversation and adding its owner.

create or replace function public.create_chat_conversation(
  conversation_type_arg text,
  title_arg text,
  description_arg text default '',
  member_user_ids_arg text[] default '{}',
  linked_record_type_arg text default null,
  linked_record_id_arg text default null
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  caller_user_id text;
  new_conversation_id uuid;
begin
  caller_user_id := public.current_user_id();
  if caller_user_id is null then
    raise exception 'An authenticated MDT profile is required.';
  end if;

  if conversation_type_arg not in ('Direct','Group','Supervisor','CAD','Case','Callsign') then
    raise exception 'This conversation type cannot be created manually.';
  end if;

  insert into public.chat_conversations (
    conversation_type, title, description, linked_record_type,
    linked_record_id, created_by
  ) values (
    conversation_type_arg,
    coalesce(nullif(trim(title_arg), ''), conversation_type_arg || ' conversation'),
    coalesce(description_arg, ''),
    linked_record_type_arg,
    linked_record_id_arg,
    caller_user_id
  ) returning conversation_id into new_conversation_id;

  insert into public.chat_members (
    conversation_id, user_id, member_role, last_read_at
  ) values (
    new_conversation_id, caller_user_id, 'Owner', now()
  );

  insert into public.chat_members (conversation_id, user_id, member_role)
  select new_conversation_id, requested_user_id, 'Member'
  from unnest(coalesce(member_user_ids_arg, '{}'::text[])) requested_user_id
  where requested_user_id <> caller_user_id
    and exists (select 1 from public.profiles p where p.user_id = requested_user_id and p.status = 'Active')
  on conflict (conversation_id, user_id) do nothing;

  return new_conversation_id;
end;
$$;

revoke all on function public.create_chat_conversation(text,text,text,text[],text,text) from public;
grant execute on function public.create_chat_conversation(text,text,text,text[],text,text) to authenticated;
