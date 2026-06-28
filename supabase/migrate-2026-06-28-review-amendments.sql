alter table public.operational_incident_reviews
  add column if not exists amendment_assignee_user_id text references public.profiles(user_id) on delete set null,
  add column if not exists amendment_response text not null default '',
  add column if not exists amendment_submitted_at timestamptz;

alter table public.operational_action_reviews
  add column if not exists amendment_response text not null default '',
  add column if not exists amendment_submitted_at timestamptz;

create index if not exists idx_operational_reviews_amendment_assignee
  on public.operational_incident_reviews(amendment_assignee_user_id, status, task_status);

drop policy if exists "operational reviews amend assigned officer" on public.operational_incident_reviews;
create policy "operational reviews amend assigned officer"
on public.operational_incident_reviews for update
using (amendment_assignee_user_id = auth.uid()::text)
with check (amendment_assignee_user_id = auth.uid()::text);

drop policy if exists "operational action reviews amend subject officer" on public.operational_action_reviews;
create policy "operational action reviews amend subject officer"
on public.operational_action_reviews for update
using (
  exists (
    select 1
    from public.officers officer
    join public.profiles profile on profile.member_id = officer.member_id
    where officer.officer_id = operational_action_reviews.officer_id
      and profile.user_id = auth.uid()::text
  )
)
with check (
  exists (
    select 1
    from public.officers officer
    join public.profiles profile on profile.member_id = officer.member_id
    where officer.officer_id = operational_action_reviews.officer_id
      and profile.user_id = auth.uid()::text
  )
);

drop policy if exists "operational action reviews read subject officer" on public.operational_action_reviews;
create policy "operational action reviews read subject officer"
on public.operational_action_reviews for select
using (
  exists (
    select 1
    from public.officers officer
    join public.profiles profile on profile.member_id = officer.member_id
    where officer.officer_id = operational_action_reviews.officer_id
      and profile.user_id = auth.uid()::text
  )
);
