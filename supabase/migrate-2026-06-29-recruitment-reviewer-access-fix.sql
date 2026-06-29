create or replace function public.can_review_recruitment_vacancy(target_vacancy_id text)
returns boolean language plpgsql security definer set search_path = public as $$
declare
  vacancy public.recruitment_vacancies;
  officer public.officers;
  profile public.profiles;
begin
  select * into vacancy from public.recruitment_vacancies where vacancy_id = target_vacancy_id;
  select * into profile from public.profiles where auth_user_id = auth.uid() limit 1;
  if vacancy is null or profile is null then return false; end if;

  -- The officer who created the posting always retains review access.
  if vacancy.created_by = profile.user_id then return true; end if;

  select * into officer from public.officers where member_id = profile.member_id limit 1;
  if officer is null then return false; end if;
  return officer.officer_id = any(coalesce(vacancy.reviewer_officer_ids, '{}'));
end;
$$;

revoke all on function public.can_review_recruitment_vacancy(text) from public;
grant execute on function public.can_review_recruitment_vacancy(text) to authenticated;
