-- Course workflow improvements: co-trainers, course deletion support, and LOA delete policy.
-- Run this once in the Supabase SQL Editor.

alter table public.training_courses
add column if not exists co_trainer_user_ids text[] not null default '{}';

drop policy if exists "course bookings read" on public.course_bookings;
create policy "course bookings read" on public.course_bookings for select
using (
  public.has_permission('MANAGE_COURSES')
  or officer_id in (select officer_id from public.officers where member_id = public.current_member_id())
  or exists (
    select 1 from public.training_courses course
    where course.course_id = course_bookings.course_id
    and (
      course.trainer_user_id = public.current_user_id()
      or public.current_user_id() = any(course.co_trainer_user_ids)
    )
  )
);

drop policy if exists "course bookings update managers" on public.course_bookings;
create policy "course bookings update managers" on public.course_bookings for update
using (
  public.has_permission('MANAGE_COURSES')
  or exists (
    select 1 from public.training_courses course
    where course.course_id = course_bookings.course_id
    and (
      course.trainer_user_id = public.current_user_id()
      or public.current_user_id() = any(course.co_trainer_user_ids)
    )
  )
)
with check (
  public.has_permission('MANAGE_COURSES')
  or exists (
    select 1 from public.training_courses course
    where course.course_id = course_bookings.course_id
    and (
      course.trainer_user_id = public.current_user_id()
      or public.current_user_id() = any(course.co_trainer_user_ids)
    )
  )
);

drop policy if exists "loa delete approvers" on public.loa_requests;
create policy "loa delete approvers" on public.loa_requests for delete
using (public.has_permission('APPROVE_LOA'));
