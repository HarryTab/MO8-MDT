-- Allow authorised training managers to update officer training tickboxes,
-- driving standards, and review dates stored in training_matrix.
-- Run once in Supabase SQL Editor.

drop policy if exists "training matrix read" on public.training_matrix;
create policy "training matrix read" on public.training_matrix for select
using (
  public.has_permission('VIEW_TRAINING')
  or officer_id in (
    select officer_id from public.officers
    where member_id = public.current_member_id()
  )
);

drop policy if exists "training matrix write" on public.training_matrix;
create policy "training matrix write" on public.training_matrix for all
using (public.has_permission('MANAGE_TRAINING'))
with check (public.has_permission('MANAGE_TRAINING'));
