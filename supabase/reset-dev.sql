-- MO8 MDT Supabase development reset
-- Only run this on a new/empty MDT Supabase project.
-- This deletes MDT tables so supabase/schema.sql can be run again cleanly.

drop table if exists public.notifications cascade;
drop table if exists public.rank_changes cascade;
drop table if exists public.audit_log cascade;
drop table if exists public.user_permissions cascade;
drop table if exists public.permissions cascade;
drop table if exists public.announcements cascade;
drop table if exists public.shift_logs cascade;
drop table if exists public.dashboard_widgets cascade;
drop table if exists public.document_acknowledgements cascade;
drop table if exists public.documents cascade;
drop table if exists public.appeals cascade;
drop table if exists public.development_plans cascade;
drop table if exists public.supervisor_checkins cascade;
drop table if exists public.supervisor_requests cascade;
drop table if exists public.transfer_requests cascade;
drop table if exists public.loa_requests cascade;
drop table if exists public.disciplinary_actions cascade;
drop table if exists public.course_bookings cascade;
drop table if exists public.training_courses cascade;
drop table if exists public.training_matrix cascade;
drop table if exists public.training_records cascade;
drop table if exists public.training_options cascade;
drop table if exists public.officers cascade;
drop table if exists public.profiles cascade;

drop function if exists public.current_profile() cascade;
drop function if exists public.current_user_id() cascade;
drop function if exists public.current_member_id() cascade;
drop function if exists public.has_permission(text) cascade;
drop function if exists public.touch_updated_at() cascade;
