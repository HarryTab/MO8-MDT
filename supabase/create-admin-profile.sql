-- MO8 MDT first admin profile link
-- 1. Create your login in Supabase Authentication > Users first.
-- 2. Copy that user's UUID.
-- 3. Replace the placeholder values below and run this in SQL Editor.

insert into public.profiles (
  auth_user_id,
  member_id,
  roblox_username,
  discord_id,
  rank,
  role,
  status,
  created_by
) values (
  'PASTE_AUTH_USER_UUID_HERE',
  'MBR_ADMIN',
  'YourRobloxUsername',
  'YourDiscordID',
  'Commissioner',
  'Command',
  'Active',
  'system'
)
on conflict (roblox_username) do update set
  auth_user_id = excluded.auth_user_id,
  rank = excluded.rank,
  role = excluded.role,
  status = excluded.status;

insert into public.officers (
  member_id,
  roblox_username,
  discord_id,
  callsign,
  rank,
  status,
  join_date,
  tags,
  notes
) values (
  'MBR_ADMIN',
  'YourRobloxUsername',
  'YourDiscordID',
  'MO8-1',
  'Commissioner',
  'Active',
  current_date,
  array['MO8 Command', 'Gold Command'],
  'Initial Supabase admin profile.'
)
on conflict (roblox_username) do update set
  member_id = excluded.member_id,
  rank = excluded.rank,
  status = excluded.status,
  tags = excluded.tags;
