-- Internal messaging, formal messages and secure chat attachments.

create or replace function public.mo8_rank_index(rank_name text)
returns integer language sql immutable as $$
  select coalesce(array_position(array[
    'Police Constable','Sergeant','Inspector','Chief Inspector','Superintendent',
    'Chief Superintendent','Commander','Deputy Assistant Commissioner',
    'Assistant Commissioner','Deputy Commissioner','Commissioner'
  ]::text[], rank_name), 0);
$$;

create table if not exists public.chat_conversations (
  conversation_id uuid primary key default gen_random_uuid(),
  conversation_type text not null check (conversation_type in ('Direct','Group','Rank','Supervisor','CAD','Case','Callsign')),
  title text not null,
  description text not null default '',
  minimum_rank text,
  linked_record_type text,
  linked_record_id text,
  created_by text not null references public.profiles(user_id) on delete cascade,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  archived_at timestamptz
);

create table if not exists public.chat_members (
  conversation_id uuid not null references public.chat_conversations(conversation_id) on delete cascade,
  user_id text not null references public.profiles(user_id) on delete cascade,
  member_role text not null default 'Member' check (member_role in ('Owner','Member')),
  last_read_at timestamptz,
  muted boolean not null default false,
  joined_at timestamptz not null default now(),
  primary key (conversation_id, user_id)
);

create table if not exists public.chat_messages (
  message_id uuid primary key default gen_random_uuid(),
  conversation_id uuid not null references public.chat_conversations(conversation_id) on delete cascade,
  sender_user_id text not null references public.profiles(user_id) on delete cascade,
  body text not null default '',
  reply_to_message_id uuid references public.chat_messages(message_id) on delete set null,
  importance text not null default 'Normal' check (importance in ('Normal','Important','Urgent')),
  requires_acknowledgement boolean not null default false,
  pinned boolean not null default false,
  attachments jsonb not null default '[]'::jsonb,
  linked_record jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  edited_at timestamptz,
  deleted_at timestamptz
);

create table if not exists public.chat_message_reactions (
  message_id uuid not null references public.chat_messages(message_id) on delete cascade,
  user_id text not null references public.profiles(user_id) on delete cascade,
  reaction text not null,
  created_at timestamptz not null default now(),
  primary key (message_id, user_id, reaction)
);

create table if not exists public.chat_message_acknowledgements (
  message_id uuid not null references public.chat_messages(message_id) on delete cascade,
  user_id text not null references public.profiles(user_id) on delete cascade,
  acknowledged_at timestamptz not null default now(),
  primary key (message_id, user_id)
);

create table if not exists public.formal_messages (
  formal_message_id uuid primary key default gen_random_uuid(),
  subject text not null,
  body text not null,
  importance text not null default 'Normal' check (importance in ('Normal','Important','Urgent')),
  requires_acknowledgement boolean not null default true,
  due_at timestamptz,
  linked_record jsonb not null default '{}'::jsonb,
  attachments jsonb not null default '[]'::jsonb,
  sender_user_id text not null references public.profiles(user_id) on delete cascade,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  withdrawn_at timestamptz
);

create table if not exists public.formal_message_recipients (
  formal_message_id uuid not null references public.formal_messages(formal_message_id) on delete cascade,
  recipient_user_id text not null references public.profiles(user_id) on delete cascade,
  read_at timestamptz,
  acknowledged_at timestamptz,
  primary key (formal_message_id, recipient_user_id)
);

create index if not exists idx_chat_messages_conversation_created on public.chat_messages(conversation_id, created_at desc);
create index if not exists idx_chat_members_user on public.chat_members(user_id, conversation_id);
create index if not exists idx_formal_recipients_user on public.formal_message_recipients(recipient_user_id, formal_message_id);

create or replace function public.can_access_chat(target_conversation uuid)
returns boolean language sql security definer stable set search_path = public as $$
  select exists (
    select 1 from public.chat_conversations c
    join public.profiles me on me.auth_user_id = auth.uid()
    where c.conversation_id = target_conversation
      and c.archived_at is null
      and (
        (c.conversation_type = 'Rank' and public.mo8_rank_index(me.rank) >= public.mo8_rank_index(c.minimum_rank))
        or (c.conversation_type <> 'Rank' and (
          c.created_by = me.user_id
          or exists (select 1 from public.chat_members cm where cm.conversation_id = c.conversation_id and cm.user_id = me.user_id)
        ))
      )
  );
$$;

create or replace function public.can_access_formal_message(target_message uuid)
returns boolean language sql security definer stable set search_path = public as $$
  select exists (
    select 1 from public.formal_messages fm
    join public.profiles me on me.auth_user_id = auth.uid()
    where fm.formal_message_id = target_message
      and (fm.sender_user_id = me.user_id or exists (
        select 1 from public.formal_message_recipients fr
        where fr.formal_message_id = fm.formal_message_id and fr.recipient_user_id = me.user_id
      ))
  );
$$;

drop function if exists public.messaging_directory();
create function public.messaging_directory()
returns table (user_id text, member_id text, roblox_username text, rank text, callsign text)
language sql security definer stable set search_path = public as $$
  select p.user_id, p.member_id, p.roblox_username, p.rank, coalesce(o.callsign, '')
  from public.profiles p
  left join public.officers o on o.member_id = p.member_id
  where p.status = 'Active' and p.auth_user_id is not null and auth.uid() is not null
  order by public.mo8_rank_index(p.rank) desc, p.roblox_username;
$$;
grant execute on function public.messaging_directory() to authenticated;

create or replace function public.touch_chat_conversation()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  update public.chat_conversations set updated_at = now() where conversation_id = new.conversation_id;
  return new;
end;
$$;
drop trigger if exists trg_touch_chat_conversation on public.chat_messages;
create trigger trg_touch_chat_conversation after insert or update on public.chat_messages
for each row execute function public.touch_chat_conversation();

alter table public.chat_conversations enable row level security;
alter table public.chat_members enable row level security;
alter table public.chat_messages enable row level security;
alter table public.chat_message_reactions enable row level security;
alter table public.chat_message_acknowledgements enable row level security;
alter table public.formal_messages enable row level security;
alter table public.formal_message_recipients enable row level security;

drop policy if exists "chat conversations accessible" on public.chat_conversations;
create policy "chat conversations accessible" on public.chat_conversations for select using (public.can_access_chat(conversation_id));
drop policy if exists "chat conversations create" on public.chat_conversations;
create policy "chat conversations create" on public.chat_conversations for insert with check (created_by = public.current_user_id());
drop policy if exists "chat conversations owners update" on public.chat_conversations;
create policy "chat conversations owners update" on public.chat_conversations for update using (created_by = public.current_user_id()) with check (created_by = public.current_user_id());
drop policy if exists "chat conversations owners delete" on public.chat_conversations;
create policy "chat conversations owners delete" on public.chat_conversations for delete using (created_by = public.current_user_id());

drop policy if exists "chat members accessible" on public.chat_members;
create policy "chat members accessible" on public.chat_members for select using (public.can_access_chat(conversation_id));
drop policy if exists "chat members add" on public.chat_members;
create policy "chat members add" on public.chat_members for insert with check (
  public.can_access_chat(conversation_id)
  or exists (select 1 from public.chat_conversations c where c.conversation_id = chat_members.conversation_id and c.created_by = public.current_user_id())
);
drop policy if exists "chat members update own or owner" on public.chat_members;
create policy "chat members update own or owner" on public.chat_members for update using (
  user_id = public.current_user_id() or exists (select 1 from public.chat_conversations c where c.conversation_id = chat_members.conversation_id and c.created_by = public.current_user_id())
) with check (public.can_access_chat(conversation_id));
drop policy if exists "chat members owner delete" on public.chat_members;
create policy "chat members owner delete" on public.chat_members for delete using (
  user_id = public.current_user_id() or exists (select 1 from public.chat_conversations c where c.conversation_id = chat_members.conversation_id and c.created_by = public.current_user_id())
);

drop policy if exists "chat messages accessible" on public.chat_messages;
create policy "chat messages accessible" on public.chat_messages for select using (public.can_access_chat(conversation_id));
drop policy if exists "chat messages send" on public.chat_messages;
create policy "chat messages send" on public.chat_messages for insert with check (sender_user_id = public.current_user_id() and public.can_access_chat(conversation_id));
drop policy if exists "chat messages sender update" on public.chat_messages;
create policy "chat messages sender update" on public.chat_messages for update using (sender_user_id = public.current_user_id()) with check (sender_user_id = public.current_user_id() and public.can_access_chat(conversation_id));

drop policy if exists "chat reactions accessible" on public.chat_message_reactions;
create policy "chat reactions accessible" on public.chat_message_reactions for select using (exists (select 1 from public.chat_messages m where m.message_id = chat_message_reactions.message_id and public.can_access_chat(m.conversation_id)));
drop policy if exists "chat reactions own" on public.chat_message_reactions;
create policy "chat reactions own" on public.chat_message_reactions for insert with check (user_id = public.current_user_id() and exists (select 1 from public.chat_messages m where m.message_id = chat_message_reactions.message_id and public.can_access_chat(m.conversation_id)));
drop policy if exists "chat reactions remove own" on public.chat_message_reactions;
create policy "chat reactions remove own" on public.chat_message_reactions for delete using (user_id = public.current_user_id());

drop policy if exists "chat acknowledgements accessible" on public.chat_message_acknowledgements;
create policy "chat acknowledgements accessible" on public.chat_message_acknowledgements for select using (exists (select 1 from public.chat_messages m where m.message_id = chat_message_acknowledgements.message_id and public.can_access_chat(m.conversation_id)));
drop policy if exists "chat acknowledgements own" on public.chat_message_acknowledgements;
create policy "chat acknowledgements own" on public.chat_message_acknowledgements for insert with check (user_id = public.current_user_id() and exists (select 1 from public.chat_messages m where m.message_id = chat_message_acknowledgements.message_id and public.can_access_chat(m.conversation_id)));

drop policy if exists "formal messages accessible" on public.formal_messages;
create policy "formal messages accessible" on public.formal_messages for select using (public.can_access_formal_message(formal_message_id));
drop policy if exists "formal messages create" on public.formal_messages;
create policy "formal messages create" on public.formal_messages for insert with check (sender_user_id = public.current_user_id());
drop policy if exists "formal messages sender update" on public.formal_messages;
create policy "formal messages sender update" on public.formal_messages for update using (sender_user_id = public.current_user_id()) with check (sender_user_id = public.current_user_id());
drop policy if exists "formal messages sender delete" on public.formal_messages;
create policy "formal messages sender delete" on public.formal_messages for delete using (sender_user_id = public.current_user_id());
drop policy if exists "formal recipients accessible" on public.formal_message_recipients;
create policy "formal recipients accessible" on public.formal_message_recipients for select using (public.can_access_formal_message(formal_message_id));
drop policy if exists "formal recipients sender add" on public.formal_message_recipients;
create policy "formal recipients sender add" on public.formal_message_recipients for insert with check (exists (select 1 from public.formal_messages fm where fm.formal_message_id = formal_message_recipients.formal_message_id and fm.sender_user_id = public.current_user_id()));
drop policy if exists "formal recipients own update" on public.formal_message_recipients;
create policy "formal recipients own update" on public.formal_message_recipients for update using (recipient_user_id = public.current_user_id()) with check (recipient_user_id = public.current_user_id());

insert into public.chat_conversations (conversation_type, title, description, minimum_rank, created_by)
select 'Rank', channel.title, channel.description, channel.minimum_rank, p.user_id
from (values
  ('Sergeant+','Supervisory discussion for Sergeant and above.','Sergeant'),
  ('Inspector+','Management discussion for Inspector and above.','Inspector'),
  ('Command','Command discussion for Superintendent and above.','Superintendent')
) as channel(title, description, minimum_rank)
cross join lateral (select user_id from public.profiles order by public.mo8_rank_index(rank) desc, created_at limit 1) p
where not exists (select 1 from public.chat_conversations existing where existing.conversation_type = 'Rank' and existing.title = channel.title);

insert into storage.buckets (id, name, public, file_size_limit)
values ('mo8-chat-files', 'mo8-chat-files', false, 10485760)
on conflict (id) do update set public = false, file_size_limit = 10485760;

drop policy if exists "chat files read" on storage.objects;
create policy "chat files read" on storage.objects for select using (
  bucket_id = 'mo8-chat-files' and (
    public.can_access_chat((storage.foldername(name))[1]::uuid)
    or public.can_access_formal_message((storage.foldername(name))[1]::uuid)
  )
);
drop policy if exists "chat files upload" on storage.objects;
create policy "chat files upload" on storage.objects for insert with check (
  bucket_id = 'mo8-chat-files' and (
    public.can_access_chat((storage.foldername(name))[1]::uuid)
    or public.can_access_formal_message((storage.foldername(name))[1]::uuid)
  )
);

do $$ begin
  alter publication supabase_realtime add table public.chat_messages;
exception when duplicate_object then null; end $$;
do $$ begin
  alter publication supabase_realtime add table public.chat_message_reactions;
exception when duplicate_object then null; end $$;
do $$ begin
  alter publication supabase_realtime add table public.formal_message_recipients;
exception when duplicate_object then null; end $$;
