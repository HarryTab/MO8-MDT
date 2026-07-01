create table if not exists public.cid_referral_exports (
  referral_id text primary key default ('CIDEXP_' || replace(gen_random_uuid()::text, '-', '')),
  referral_reference text not null unique default ('CID-' || to_char(now(), 'YYYY') || '-' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 8))),
  version integer not null default 1 check (version > 0),
  source_type text not null check (source_type in ('CAD', 'Case')),
  source_id text not null,
  source_reference text not null default '',
  incident_ids text[] not null default array[]::text[],
  evidence_ids text[] not null default array[]::text[],
  submitting_officer_id text references public.officers(officer_id) on delete set null,
  created_by text references public.profiles(user_id) on delete set null,
  classification text not null default 'Official - Internal',
  referral_summary text not null default '',
  requested_action text not null default '',
  suspected_offences text not null default '',
  sent_to text not null default '',
  sent_at timestamptz,
  declaration text not null default '',
  signature_storage_path text not null,
  pdf_storage_path text not null,
  pdf_hash text not null,
  pdf_size bigint not null default 0,
  evidence_link_expires_at timestamptz,
  snapshot jsonb not null default '{}'::jsonb,
  status text not null default 'Exported' check (status in ('Exported', 'Sent', 'Superseded', 'Cancelled')),
  exported_at timestamptz not null default now()
);

create index if not exists idx_cid_referral_source
  on public.cid_referral_exports(source_type, source_id, exported_at desc);
create index if not exists idx_cid_referral_officer
  on public.cid_referral_exports(submitting_officer_id, exported_at desc);

alter table public.cid_referral_exports enable row level security;

drop policy if exists "cid referrals read authenticated" on public.cid_referral_exports;
create policy "cid referrals read authenticated" on public.cid_referral_exports
for select using (auth.role() = 'authenticated');

drop policy if exists "cid referrals create authenticated" on public.cid_referral_exports;
create policy "cid referrals create authenticated" on public.cid_referral_exports
for insert with check (auth.role() = 'authenticated');

drop policy if exists "cid referrals update owner or supervisors" on public.cid_referral_exports;
create policy "cid referrals update owner or supervisors" on public.cid_referral_exports
for update using (
  exists (
    select 1 from public.profiles profile
    where profile.user_id = cid_referral_exports.created_by
      and profile.member_id = public.current_member_id()
  )
  or public.has_permission('VIEW_TASKS')
) with check (
  exists (
    select 1 from public.profiles profile
    where profile.user_id = cid_referral_exports.created_by
      and profile.member_id = public.current_member_id()
  )
  or public.has_permission('VIEW_TASKS')
);

insert into storage.buckets (id, name, public, file_size_limit)
values ('mo8-cid-referrals', 'mo8-cid-referrals', false, 52428800)
on conflict (id) do update set file_size_limit = excluded.file_size_limit;

drop policy if exists "cid referral files read authenticated" on storage.objects;
create policy "cid referral files read authenticated" on storage.objects for select
using (bucket_id = 'mo8-cid-referrals' and auth.role() = 'authenticated');

drop policy if exists "cid referral files upload authenticated" on storage.objects;
create policy "cid referral files upload authenticated" on storage.objects for insert
with check (bucket_id = 'mo8-cid-referrals' and auth.role() = 'authenticated');

drop policy if exists "cid referral files update owner or supervisors" on storage.objects;
create policy "cid referral files update owner or supervisors" on storage.objects for update
using (bucket_id = 'mo8-cid-referrals' and auth.role() = 'authenticated')
with check (bucket_id = 'mo8-cid-referrals' and auth.role() = 'authenticated');

drop policy if exists "cid referral files delete authenticated" on storage.objects;
create policy "cid referral files delete authenticated" on storage.objects for delete
using (bucket_id = 'mo8-cid-referrals' and auth.role() = 'authenticated');
