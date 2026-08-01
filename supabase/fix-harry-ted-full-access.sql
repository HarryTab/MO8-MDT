-- One-off admin repair:
-- Sets the MDT account "harry_ted" to explicit FULL_ACCESS.
-- Run this in Supabase SQL Editor.

insert into public.user_permissions (user_id, permission, allowed, updated_at)
select user_id, 'FULL_ACCESS', 'Allow', now()
from public.profiles
where lower(roblox_username) = 'harry_ted'
on conflict (user_id, permission)
do update set
  allowed = 'Allow',
  updated_at = now();

insert into public.user_permissions (user_id, permission, allowed, updated_at)
select user_id, 'MANAGE_PERMISSIONS', 'Allow', now()
from public.profiles
where lower(roblox_username) = 'harry_ted'
on conflict (user_id, permission)
do update set
  allowed = 'Allow',
  updated_at = now();

insert into public.audit_log (actor_user_id, action, target_type, target_id, details)
select user_id, 'REPAIR_FULL_ACCESS', 'User', user_id, jsonb_build_object(
  'robloxUsername', roblox_username,
  'permissions', jsonb_build_array('FULL_ACCESS', 'MANAGE_PERMISSIONS'),
  'reason', 'One-off repair for command/admin account access'
)
from public.profiles
where lower(roblox_username) = 'harry_ted';
