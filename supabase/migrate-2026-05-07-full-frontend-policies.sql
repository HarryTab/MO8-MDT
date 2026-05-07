-- MO8 MDT Supabase policy patch for the full frontend migration.
-- Run this after supabase/schema.sql if your project was created before this patch.

drop policy if exists "transfer read" on public.transfer_requests;
drop policy if exists "transfer insert own" on public.transfer_requests;
drop policy if exists "transfer update managers" on public.transfer_requests;
create policy "transfer read" on public.transfer_requests for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "transfer insert own" on public.transfer_requests for insert
with check (officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "transfer update managers" on public.transfer_requests for update
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "supervisor requests read" on public.supervisor_requests;
drop policy if exists "supervisor requests insert own" on public.supervisor_requests;
drop policy if exists "supervisor requests update managers" on public.supervisor_requests;
create policy "supervisor requests read" on public.supervisor_requests for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "supervisor requests insert own" on public.supervisor_requests for insert
with check (officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "supervisor requests update managers" on public.supervisor_requests for update
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "supervisor checkins read" on public.supervisor_checkins;
drop policy if exists "supervisor checkins write managers" on public.supervisor_checkins;
create policy "supervisor checkins read" on public.supervisor_checkins for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "supervisor checkins write managers" on public.supervisor_checkins for all
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "development plans read" on public.development_plans;
drop policy if exists "development plans write managers" on public.development_plans;
create policy "development plans read" on public.development_plans for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "development plans write managers" on public.development_plans for all
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "appeals read" on public.appeals;
drop policy if exists "appeals insert own" on public.appeals;
drop policy if exists "appeals update managers" on public.appeals;
create policy "appeals read" on public.appeals for select
using (public.has_permission('VIEW_TASKS') or officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "appeals insert own" on public.appeals for insert
with check (officer_id in (select officer_id from public.officers where member_id = public.current_member_id()));
create policy "appeals update managers" on public.appeals for update
using (public.has_permission('VIEW_TASKS'))
with check (public.has_permission('VIEW_TASKS'));

drop policy if exists "document acknowledgements read own or managers" on public.document_acknowledgements;
drop policy if exists "document acknowledgements write own" on public.document_acknowledgements;
create policy "document acknowledgements read own or managers" on public.document_acknowledgements for select
using (user_id = public.current_user_id() or public.has_permission('MANAGE_DOCUMENTS'));
create policy "document acknowledgements write own" on public.document_acknowledgements for all
using (user_id = public.current_user_id())
with check (user_id = public.current_user_id());

drop policy if exists "audit insert authenticated" on public.audit_log;
create policy "audit insert authenticated" on public.audit_log for insert
with check (auth.uid() is not null);

drop policy if exists "rank changes insert authorised" on public.rank_changes;
create policy "rank changes insert authorised" on public.rank_changes for insert
with check (public.has_permission('EDIT_OFFICERS') or public.has_permission('MANAGE_USERS'));

drop policy if exists "notifications insert authenticated" on public.notifications;
create policy "notifications insert authenticated" on public.notifications for insert
with check (auth.uid() is not null);

drop policy if exists "profiles insert managers" on public.profiles;
create policy "profiles insert managers" on public.profiles for insert
with check (public.has_permission('MANAGE_USERS') or public.has_permission('ADD_OFFICERS'));

