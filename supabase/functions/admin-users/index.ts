import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

type SaveUserPayload = {
  action: 'saveUser' | 'resetPassword' | 'deleteUser';
  UserID?: string;
  RobloxUsername?: string;
  DiscordID?: string;
  Rank?: string;
  Role?: string;
  Status?: string;
  TemporaryPassword?: string;
};

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders });
  }

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL');
    const anonKey = Deno.env.get('SUPABASE_ANON_KEY');
    const serviceRoleKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY');
    if (!supabaseUrl || !anonKey || !serviceRoleKey) {
      return json({ ok: false, error: 'Supabase function secrets are not configured.' }, 500);
    }

    const authorization = request.headers.get('Authorization') || '';
    const userClient = createClient(supabaseUrl, anonKey, {
      global: { headers: { Authorization: authorization } },
    });
    const adminClient = createClient(supabaseUrl, serviceRoleKey);

    const { data: authData, error: authError } = await userClient.auth.getUser();
    if (authError || !authData.user) return json({ ok: false, error: 'Not signed in.' }, 401);

    const payload = await request.json() as SaveUserPayload;
    const requiredPermission = payload.action === 'resetPassword' ? 'RESET_PASSWORDS' : 'MANAGE_USERS';
    const { data: allowed, error: permissionError } = await userClient.rpc('has_permission', {
      permission_name: requiredPermission,
    });
    if (permissionError) return json({ ok: false, error: permissionError.message }, 500);
    if (!allowed) return json({ ok: false, error: 'You do not have permission to manage login accounts.' }, 403);

    const { data: actorProfile } = await adminClient
      .from('profiles')
      .select('*')
      .eq('auth_user_id', authData.user.id)
      .limit(1)
      .maybeSingle();

    if (payload.action === 'resetPassword') {
      return await resetPassword(adminClient, payload);
    }
    if (payload.action === 'deleteUser') {
      return await deleteUser(adminClient, payload);
    }
    return await saveUser(adminClient, payload, actorProfile?.user_id || null);
  } catch (error) {
    return json({ ok: false, error: error instanceof Error ? error.message : String(error) }, 500);
  }
});

async function saveUser(adminClient: ReturnType<typeof createClient>, payload: SaveUserPayload, actorUserId: string | null) {
  const username = String(payload.RobloxUsername || '').trim();
  if (!username) return json({ ok: false, error: 'Roblox username is required.' }, 400);

  const userId = payload.UserID || id('USR');
  const existingProfile = payload.UserID
    ? await byId(adminClient, 'profiles', 'user_id', payload.UserID)
    : await byId(adminClient, 'profiles', 'roblox_username', username);
  const memberId = existingProfile?.member_id || id('MBR');
  const email = `${username.toLowerCase().replace(/[^a-z0-9._-]/g, '')}@mo8.local`;
  const temporaryPassword = payload.TemporaryPassword || randomPassword();

  let authUserId = existingProfile?.auth_user_id || null;
  if (!authUserId) {
    const { data, error } = await adminClient.auth.admin.createUser({
      email,
      password: temporaryPassword,
      email_confirm: true,
      user_metadata: { roblox_username: username },
    });
    if (error) return json({ ok: false, error: error.message }, 400);
    authUserId = data.user?.id || null;
  } else if (payload.TemporaryPassword) {
    const { error } = await adminClient.auth.admin.updateUserById(authUserId, {
      password: payload.TemporaryPassword,
      email_confirm: true,
    });
    if (error) return json({ ok: false, error: error.message }, 400);
  }

  const profileRecord = {
    user_id: userId,
    auth_user_id: authUserId,
    member_id: memberId,
    roblox_username: username,
    discord_id: payload.DiscordID || '',
    rank: payload.Rank || 'Police Constable',
    role: payload.Role || rankToRole(payload.Rank || ''),
    status: payload.Status || 'Active',
    created_by: existingProfile?.created_by || actorUserId,
  };

  const profileResult = existingProfile
    ? await adminClient.from('profiles').update(profileRecord).eq('user_id', existingProfile.user_id)
    : await adminClient.from('profiles').insert(profileRecord);
  if (profileResult.error) return json({ ok: false, error: profileResult.error.message }, 400);

  const existingOfficer = await byId(adminClient, 'officers', 'member_id', memberId);
  const officerRecord = {
    officer_id: existingOfficer?.officer_id || id('OFF'),
    member_id: memberId,
    roblox_username: username,
    discord_id: payload.DiscordID || '',
    rank: payload.Rank || 'Police Constable',
    status: payload.Status || 'Active',
  };
  const officerResult = existingOfficer
    ? await adminClient.from('officers').update(officerRecord).eq('officer_id', existingOfficer.officer_id)
    : await adminClient.from('officers').insert(officerRecord);
  if (officerResult.error) return json({ ok: false, error: officerResult.error.message }, 400);

  return json({
    ok: true,
    UserID: existingProfile?.user_id || userId,
    authUserId,
    email,
    temporaryPassword: payload.TemporaryPassword ? '' : temporaryPassword,
  });
}

async function resetPassword(adminClient: ReturnType<typeof createClient>, payload: SaveUserPayload) {
  if (!payload.UserID) return json({ ok: false, error: 'UserID is required.' }, 400);
  const profile = await byId(adminClient, 'profiles', 'user_id', payload.UserID);
  if (!profile) return json({ ok: false, error: 'User profile not found.' }, 404);

  let authUserId = profile.auth_user_id;
  const temporaryPassword = payload.TemporaryPassword || randomPassword();
  if (!authUserId) {
    const email = `${String(profile.roblox_username || '').toLowerCase().replace(/[^a-z0-9._-]/g, '')}@mo8.local`;
    const { data, error } = await adminClient.auth.admin.createUser({
      email,
      password: temporaryPassword,
      email_confirm: true,
      user_metadata: { roblox_username: profile.roblox_username },
    });
    if (error) return json({ ok: false, error: error.message }, 400);
    authUserId = data.user?.id;
    await adminClient.from('profiles').update({ auth_user_id: authUserId }).eq('user_id', payload.UserID);
  } else {
    const { error } = await adminClient.auth.admin.updateUserById(authUserId, {
      password: temporaryPassword,
      email_confirm: true,
    });
    if (error) return json({ ok: false, error: error.message }, 400);
  }

  return json({ ok: true, temporaryPassword });
}

async function deleteUser(adminClient: ReturnType<typeof createClient>, payload: SaveUserPayload) {
  if (!payload.UserID) return json({ ok: false, error: 'UserID is required.' }, 400);
  const profile = await byId(adminClient, 'profiles', 'user_id', payload.UserID);
  if (profile?.auth_user_id) {
    const { error } = await adminClient.auth.admin.deleteUser(profile.auth_user_id);
    if (error) return json({ ok: false, error: error.message }, 400);
  }
  if (profile?.member_id) await adminClient.from('officers').delete().eq('member_id', profile.member_id);
  const { error } = await adminClient.from('profiles').delete().eq('user_id', payload.UserID);
  if (error) return json({ ok: false, error: error.message }, 400);
  return json({ ok: true });
}

async function byId(adminClient: ReturnType<typeof createClient>, table: string, column: string, value: string) {
  const { data, error } = await adminClient.from(table).select('*').eq(column, value).limit(1).maybeSingle();
  if (error) throw error;
  return data;
}

function json(body: Record<string, unknown>, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
  });
}

function id(prefix: string) {
  return `${prefix}_${crypto.randomUUID().replaceAll('-', '')}`;
}

function randomPassword() {
  return `MO8-${crypto.randomUUID().slice(0, 8)}!`;
}

function rankToRole(rank: string) {
  if (['Commissioner', 'Deputy Commissioner', 'Assistant Commissioner', 'Deputy Assistant Commissioner', 'Commander', 'Chief Superintendent', 'Superintendent'].includes(rank)) return 'Command';
  if (rank === 'Chief Inspector') return 'Chief Inspector';
  if (rank === 'Inspector') return 'Inspector';
  if (rank === 'Sergeant') return 'Sergeant';
  return 'Constable';
}
