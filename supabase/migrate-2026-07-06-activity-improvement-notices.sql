create table if not exists public.activity_improvement_notices (
  aip_id text primary key default ('AIP_' || replace(gen_random_uuid()::text, '-', '')),
  reference text not null unique default ('AIP-' || to_char(now(), 'YYYY') || '-' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 8))),
  officer_id text not null references public.officers(officer_id) on delete restrict,
  created_by text references public.profiles(user_id) on delete set null,
  line_manager_user_id text not null references public.profiles(user_id) on delete restrict,
  authorising_manager_user_id text not null references public.profiles(user_id) on delete restrict,
  status text not null default 'Awaiting Signatures' check (status in ('Draft','Awaiting Signatures','Issued','Under Review','Closed','Extended','Withdrawn')),
  issue_date date not null default current_date,
  review_end_date date not null,
  reason text not null,
  required_hours numeric(7,2) not null default 6,
  adjusted_required_hours numeric(7,2),
  expected_standards text not null default '',
  support_guidance text not null default '',
  consequences text not null default '',
  activity_snapshot jsonb not null default '{}'::jsonb,
  loa_snapshot jsonb not null default '[]'::jsonb,
  officer_response text not null default '',
  officer_acknowledged_at timestamptz,
  review_outcome text check (review_outcome is null or review_outcome in ('Requirements Met','Partial Improvement','No Improvement','Extended','Withdrawn')),
  review_comments text not null default '',
  reviewed_at timestamptz,
  reviewed_by text references public.profiles(user_id) on delete set null,
  withdrawn_reason text not null default '',
  withdrawn_at timestamptz,
  version integer not null default 1,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check (review_end_date >= issue_date),
  check (line_manager_user_id <> authorising_manager_user_id)
);

create table if not exists public.aip_signatures (
  signature_id text primary key default ('AIPSIG_' || replace(gen_random_uuid()::text, '-', '')),
  aip_id text not null references public.activity_improvement_notices(aip_id) on delete cascade,
  signature_role text not null check (signature_role in ('Line Manager','Manager Authorisation')),
  signer_user_id text references public.profiles(user_id) on delete set null,
  signer_name text not null,
  signer_rank text not null default '',
  signature_data text not null,
  signed_at timestamptz not null default now(),
  signed_via text not null default 'MDT' check (signed_via in ('MDT','Secure Link')),
  version integer not null default 1,
  unique (aip_id, signature_role, version)
);

create table if not exists public.aip_signing_links (
  link_id text primary key default ('AIPLINK_' || replace(gen_random_uuid()::text, '-', '')),
  aip_id text not null references public.activity_improvement_notices(aip_id) on delete cascade,
  signature_role text not null check (signature_role in ('Line Manager','Manager Authorisation')),
  signer_user_id text not null references public.profiles(user_id) on delete cascade,
  token text not null unique default replace(gen_random_uuid()::text || gen_random_uuid()::text, '-', ''),
  expires_at timestamptz not null default (now() + interval '7 days'),
  used_at timestamptz,
  revoked_at timestamptz,
  created_by text references public.profiles(user_id) on delete set null,
  created_at timestamptz not null default now()
);

create index if not exists idx_aip_officer on public.activity_improvement_notices(officer_id, created_at desc);
create index if not exists idx_aip_review on public.activity_improvement_notices(review_end_date, status);
create index if not exists idx_aip_line_manager on public.activity_improvement_notices(line_manager_user_id, status);
create index if not exists idx_aip_authoriser on public.activity_improvement_notices(authorising_manager_user_id, status);
create index if not exists idx_aip_signatures on public.aip_signatures(aip_id, version);

alter table public.activity_improvement_notices enable row level security;
alter table public.aip_signatures enable row level security;
alter table public.aip_signing_links enable row level security;

drop policy if exists "aip read relevant staff" on public.activity_improvement_notices;
create policy "aip read relevant staff" on public.activity_improvement_notices for select using (
  public.has_permission('VIEW_TASKS')
  or line_manager_user_id in (select user_id from public.profiles where member_id = public.current_member_id())
  or authorising_manager_user_id in (select user_id from public.profiles where member_id = public.current_member_id())
  or officer_id in (select officer_id from public.officers where member_id = public.current_member_id())
);

drop policy if exists "aip manage supervisors" on public.activity_improvement_notices;
create policy "aip manage supervisors" on public.activity_improvement_notices for all
using (public.has_permission('VIEW_TASKS')) with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "aip signatures read relevant" on public.aip_signatures;
create policy "aip signatures read relevant" on public.aip_signatures for select using (
  exists (select 1 from public.activity_improvement_notices aip where aip.aip_id = aip_signatures.aip_id)
);

drop policy if exists "aip signatures create assigned" on public.aip_signatures;
create policy "aip signatures create assigned" on public.aip_signatures for insert with check (
  signer_user_id in (select user_id from public.profiles where member_id = public.current_member_id())
  and exists (
    select 1 from public.activity_improvement_notices aip
    where aip.aip_id = aip_signatures.aip_id
      and ((signature_role = 'Line Manager' and aip.line_manager_user_id = signer_user_id)
        or (signature_role = 'Manager Authorisation' and aip.authorising_manager_user_id = signer_user_id))
  )
);

drop policy if exists "aip links supervisors read" on public.aip_signing_links;
create policy "aip links supervisors read" on public.aip_signing_links for select using (
  public.has_permission('VIEW_TASKS') or signer_user_id in (select user_id from public.profiles where member_id = public.current_member_id())
);
drop policy if exists "aip links supervisors create" on public.aip_signing_links;
create policy "aip links supervisors create" on public.aip_signing_links for insert with check (public.has_permission('VIEW_TASKS'));
drop policy if exists "aip links supervisors update" on public.aip_signing_links;
create policy "aip links supervisors update" on public.aip_signing_links for update using (public.has_permission('VIEW_TASKS'));

create or replace function public.refresh_aip_status(target_aip_id text)
returns void language plpgsql security definer set search_path = public as $$
declare signature_count integer;
begin
  select count(*) into signature_count from public.aip_signatures signature
  join public.activity_improvement_notices aip on aip.aip_id = signature.aip_id
  where signature.aip_id = target_aip_id and signature.version = aip.version;
  if signature_count >= 2 then
    update public.activity_improvement_notices
      set status = case when status = 'Awaiting Signatures' then 'Issued' else status end,
          updated_at = now()
      where aip_id = target_aip_id;
    insert into public.notifications(member_id, title, message)
    select officer.member_id, 'Activity Improvement Notice issued', aip.reference || ' is now issued. Activity will be reviewed on ' || to_char(aip.review_end_date, 'DD/MM/YYYY') || '.'
    from public.activity_improvement_notices aip join public.officers officer on officer.officer_id = aip.officer_id
    where aip.aip_id = target_aip_id and not exists (
      select 1 from public.notifications notice where notice.member_id = officer.member_id and notice.title = 'Activity Improvement Notice issued' and notice.message like aip.reference || '%'
    );
  end if;
end;
$$;

create or replace function public.acknowledge_aip(target_aip_id text, officer_response_text text default '')
returns jsonb language plpgsql security definer set search_path = public as $$
declare aip_record public.activity_improvement_notices%rowtype;
begin
  select aip.* into aip_record from public.activity_improvement_notices aip
  join public.officers officer on officer.officer_id = aip.officer_id
  where aip.aip_id = target_aip_id and officer.member_id = public.current_member_id();
  if not found then return jsonb_build_object('ok', false, 'error', 'This notice is not assigned to your officer account.'); end if;
  if aip_record.status not in ('Issued','Under Review','Extended','Closed') then return jsonb_build_object('ok', false, 'error', 'This notice has not yet been issued.'); end if;
  update public.activity_improvement_notices set officer_acknowledged_at = coalesce(officer_acknowledged_at, now()), officer_response = coalesce(officer_response_text, ''), updated_at = now() where aip_id = target_aip_id;
  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.get_aip_signing_request(signing_token text)
returns jsonb language plpgsql security definer set search_path = public as $$
declare result jsonb;
begin
  select jsonb_build_object(
    'ok', true, 'aipId', aip.aip_id, 'reference', aip.reference, 'role', link.signature_role,
    'officer', concat_ws(' ', officer.rank, officer.roblox_username), 'issueDate', aip.issue_date,
    'reviewEndDate', aip.review_end_date, 'reason', aip.reason, 'expectedStandards', aip.expected_standards,
    'supportGuidance', aip.support_guidance, 'signer', concat_ws(' ', profile.rank, profile.roblox_username)
  ) into result
  from public.aip_signing_links link
  join public.activity_improvement_notices aip on aip.aip_id = link.aip_id
  join public.officers officer on officer.officer_id = aip.officer_id
  join public.profiles profile on profile.user_id = link.signer_user_id
  where link.token = signing_token and link.used_at is null and link.revoked_at is null and link.expires_at > now()
    and aip.status = 'Awaiting Signatures';
  return coalesce(result, jsonb_build_object('ok', false, 'error', 'This signing link is invalid, expired or has already been used.'));
end;
$$;

create or replace function public.sign_aip_with_token(signing_token text, drawn_signature text)
returns jsonb language plpgsql security definer set search_path = public as $$
declare link_record public.aip_signing_links%rowtype;
declare aip_record public.activity_improvement_notices%rowtype;
declare signer public.profiles%rowtype;
begin
  if length(coalesce(drawn_signature, '')) < 100 then return jsonb_build_object('ok', false, 'error', 'A drawn signature is required.'); end if;
  select * into link_record from public.aip_signing_links where token = signing_token and used_at is null and revoked_at is null and expires_at > now() for update;
  if not found then return jsonb_build_object('ok', false, 'error', 'This signing link is invalid, expired or has already been used.'); end if;
  select * into aip_record from public.activity_improvement_notices where aip_id = link_record.aip_id;
  select * into signer from public.profiles where user_id = link_record.signer_user_id;
  insert into public.aip_signatures(aip_id, signature_role, signer_user_id, signer_name, signer_rank, signature_data, signed_via, version)
  values (aip_record.aip_id, link_record.signature_role, signer.user_id, signer.roblox_username, signer.rank, drawn_signature, 'Secure Link', aip_record.version)
  on conflict (aip_id, signature_role, version) do nothing;
  update public.aip_signing_links set used_at = now() where link_id = link_record.link_id;
  perform public.refresh_aip_status(aip_record.aip_id);
  return jsonb_build_object('ok', true, 'reference', aip_record.reference);
end;
$$;

grant execute on function public.get_aip_signing_request(text) to anon, authenticated;
grant execute on function public.sign_aip_with_token(text, text) to anon, authenticated;
grant execute on function public.refresh_aip_status(text) to authenticated;
grant execute on function public.acknowledge_aip(text, text) to authenticated;
