create extension if not exists pgcrypto;

create or replace function public.recruitment_register(applicant_username text, applicant_discord_id text, applicant_password text)
returns jsonb language plpgsql security definer set search_path = public, extensions as $$
declare account public.recruitment_accounts; token uuid;
begin
  if length(trim(applicant_username)) < 3 or length(trim(applicant_username)) > 40 then raise exception 'Enter a valid Roblox username.'; end if;
  if applicant_discord_id !~ '^\d{15,22}$' then raise exception 'Enter a valid Discord user ID.'; end if;
  if length(applicant_password) < 8 then raise exception 'Password must be at least 8 characters.'; end if;
  insert into public.recruitment_accounts(roblox_username, discord_id, password_hash)
  values (trim(applicant_username), applicant_discord_id, crypt(applicant_password, gen_salt('bf')))
  returning * into account;
  insert into public.recruitment_sessions(account_id) values (account.account_id) returning session_token into token;
  return jsonb_build_object('token', token, 'username', account.roblox_username, 'discordId', account.discord_id);
exception when unique_violation then raise exception 'A recruitment account already exists for that Roblox username.';
end;
$$;

create or replace function public.recruitment_login(applicant_username text, applicant_password text)
returns jsonb language plpgsql security definer set search_path = public, extensions as $$
declare account public.recruitment_accounts; token uuid;
begin
  select * into account from public.recruitment_accounts where username_normalised = lower(trim(applicant_username)) and status = 'Active';
  if account is null or account.password_hash <> crypt(applicant_password, account.password_hash) then raise exception 'Username or password is incorrect.'; end if;
  delete from public.recruitment_sessions where expires_at < now();
  insert into public.recruitment_sessions(account_id) values (account.account_id) returning session_token into token;
  return jsonb_build_object('token', token, 'username', account.roblox_username, 'discordId', account.discord_id);
end;
$$;

revoke all on function public.recruitment_register(text, text, text) from public;
revoke all on function public.recruitment_login(text, text) from public;
grant execute on function public.recruitment_register(text, text, text) to anon, authenticated;
grant execute on function public.recruitment_login(text, text) to anon, authenticated;
