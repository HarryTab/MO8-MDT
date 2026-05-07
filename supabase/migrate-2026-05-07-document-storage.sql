-- MO8 MDT document file upload support.
-- Run this in Supabase SQL Editor before uploading documents through the MDT.

insert into storage.buckets (id, name, public)
values ('mo8-documents', 'mo8-documents', false)
on conflict (id) do nothing;

alter table public.documents add column if not exists storage_path text;
alter table public.documents add column if not exists file_name text;
alter table public.documents add column if not exists file_size bigint;
alter table public.documents add column if not exists file_type text;

drop policy if exists "document files read" on storage.objects;
drop policy if exists "document files insert managers" on storage.objects;
drop policy if exists "document files update managers" on storage.objects;
drop policy if exists "document files delete managers" on storage.objects;

create policy "document files read" on storage.objects for select
using (
  bucket_id = 'mo8-documents'
  and public.has_permission('VIEW_DOCUMENTS')
);

create policy "document files insert managers" on storage.objects for insert
with check (
  bucket_id = 'mo8-documents'
  and public.has_permission('MANAGE_DOCUMENTS')
);

create policy "document files update managers" on storage.objects for update
using (
  bucket_id = 'mo8-documents'
  and public.has_permission('MANAGE_DOCUMENTS')
)
with check (
  bucket_id = 'mo8-documents'
  and public.has_permission('MANAGE_DOCUMENTS')
);

create policy "document files delete managers" on storage.objects for delete
using (
  bucket_id = 'mo8-documents'
  and public.has_permission('MANAGE_DOCUMENTS')
);
