import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

type DiscordAlertPayload = {
  action?: 'sendNotification' | 'testIdentity' | 'testDm' | 'recruitmentSubmitted' | 'recruitmentSubmittedInternal' | 'recruitmentStatus';
  notificationId?: string;
  applicationId?: string;
  recruitmentToken?: string;
  discordId?: string;
  title?: string;
  message?: string;
};

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });

  try {
    const supabaseUrl = Deno.env.get('SUPABASE_URL');
    const anonKey = Deno.env.get('SUPABASE_ANON_KEY');
    const serviceRoleKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY');
    const botToken = Deno.env.get('DISCORD_BOT_TOKEN');
    if (!supabaseUrl || !anonKey || !serviceRoleKey || !botToken) {
      return json({ ok: false, error: 'Discord alert function secrets are not configured.' }, 500);
    }

    const authorization = request.headers.get('Authorization') || '';
    const userClient = createClient(supabaseUrl, anonKey, {
      global: { headers: { Authorization: authorization } },
    });
    const adminClient = createClient(supabaseUrl, serviceRoleKey);
    const payload = await request.json() as DiscordAlertPayload;

    if (payload.action === 'recruitmentSubmitted') {
      return await sendRecruitmentDm(adminClient, botToken, payload, null);
    }

    const { data: authData, error: authError } = await userClient.auth.getUser();
    if (authError || !authData.user) return json({ ok: false, error: 'Not signed in.' }, 401);

    const { data: actorProfile } = await adminClient
      .from('profiles')
      .select('*')
      .eq('auth_user_id', authData.user.id)
      .limit(1)
      .maybeSingle();
    if (!actorProfile) return json({ ok: false, error: 'No MDT profile found for this login.' }, 403);

    if (payload.action === 'testIdentity') return await testIdentity(botToken);
    if (payload.action === 'testDm') return await testDm(botToken, payload, actorProfile.discord_id || '');
    if (payload.action === 'recruitmentSubmittedInternal' || payload.action === 'recruitmentStatus') return await sendRecruitmentDm(adminClient, botToken, payload, actorProfile);
    return await sendNotificationDm(adminClient, botToken, payload, actorProfile);
  } catch (error) {
    return json({ ok: false, error: error instanceof Error ? error.message : String(error) }, 500);
  }
});

async function sendNotificationDm(adminClient: ReturnType<typeof createClient>, botToken: string, payload: DiscordAlertPayload, actorProfile: Record<string, unknown>) {
  if (!payload.notificationId) return json({ ok: false, error: 'notificationId is required.' }, 400);

  const { data: notification, error: notificationError } = await adminClient
    .from('notifications')
    .select('*')
    .eq('notification_id', payload.notificationId)
    .limit(1)
    .maybeSingle();
  if (notificationError) return json({ ok: false, error: notificationError.message }, 500);
  if (!notification) return json({ ok: false, error: 'Notification not found.' }, 404);

  const actorUserId = String(actorProfile.user_id || '');
  if (notification.actor_user_id && notification.actor_user_id !== actorUserId) {
    return json({ ok: false, error: 'You can only send Discord alerts for notifications you created.' }, 403);
  }

  const { data: recipient } = await adminClient
    .from('profiles')
    .select('*')
    .eq('member_id', notification.member_id)
    .limit(1)
    .maybeSingle();
  const discordId = digitsOnly(recipient?.discord_id || '');
  if (!discordId) return json({ ok: true, sent: false, skipped: 'No Discord ID is stored for this officer.' });

  const result = await sendDiscordDm(botToken, discordId, notification.title || 'MO8 MDT notification', notification.message || '');
  return json({ ok: result.ok, sent: result.ok, error: result.error || '', discordId });
}

async function sendRecruitmentDm(adminClient: ReturnType<typeof createClient>, botToken: string, payload: DiscordAlertPayload, actorProfile: Record<string, unknown> | null) {
  if (!payload.applicationId) return json({ ok: false, error: 'applicationId is required.' }, 400);
  const { data: application, error } = await adminClient.from('recruitment_applications').select('*').eq('application_id', payload.applicationId).limit(1).maybeSingle();
  if (error || !application) return json({ ok: false, error: error?.message || 'Application not found.' }, 404);
  const { data: vacancy } = await adminClient.from('recruitment_vacancies').select('*').eq('vacancy_id', application.vacancy_id).limit(1).maybeSingle();

  if (!actorProfile) {
    if (!payload.recruitmentToken || !application.applicant_account_id) return json({ ok: false, error: 'Recruitment session is required.' }, 401);
    const { data: session } = await adminClient.from('recruitment_sessions').select('*').eq('session_token', payload.recruitmentToken).eq('account_id', application.applicant_account_id).gt('expires_at', new Date().toISOString()).limit(1).maybeSingle();
    if (!session || payload.action !== 'recruitmentSubmitted') return json({ ok: false, error: 'Recruitment session is invalid.' }, 401);
  } else if (payload.action === 'recruitmentSubmittedInternal') {
    if (application.internal_member_id !== actorProfile.member_id) return json({ ok: false, error: 'This is not your application.' }, 403);
  } else if (payload.action === 'recruitmentStatus' && application.reviewer_user_id !== actorProfile.user_id) {
    return json({ ok: false, error: 'Only the recorded reviewer can send this update.' }, 403);
  }

  const discordId = digitsOnly(application.discord_id || '');
  if (!discordId) return json({ ok: true, sent: false, skipped: 'No Discord ID is stored for this applicant.' });
  const submitted = payload.action !== 'recruitmentStatus';
  const title = submitted ? 'MO8 application received' : `MO8 application ${String(application.status || 'updated').toLowerCase()}`;
  const message = [
    submitted ? 'Your application has been received and added to the recruitment queue.' : (application.applicant_message || 'Your application status has been updated.'),
    `Role: ${vacancy?.title || application.vacancy_id}`,
    `Status: ${application.status}`,
    `Application: ${application.application_id}`,
    application.reviewed_at ? `Updated: ${new Date(application.reviewed_at).toLocaleString('en-GB', { timeZone: 'Europe/London' })}` : '',
  ].filter(Boolean).join('\n');
  const careersUrl = `${(Deno.env.get('MDT_URL') || 'https://harrytab.github.io/MO8-MDT/').replace(/\/$/, '')}/careers/`;
  const result = await sendDiscordDm(botToken, discordId, title, message, careersUrl, 'Open applicant portal');
  return json({ ok: result.ok, sent: result.ok, error: result.error || '', discordId }, result.ok ? 200 : 400);
}

async function testIdentity(botToken: string) {
  const response = await discordFetch(botToken, 'https://discord.com/api/v10/users/@me');
  const body = await response.json().catch(() => ({}));
  if (!response.ok) return json({ ok: false, error: discordError(response, body) }, response.status);
  return json({
    ok: true,
    botId: body.id,
    botUsername: body.username,
    discriminator: body.discriminator,
    isBot: Boolean(body.bot),
  });
}

async function testDm(botToken: string, payload: DiscordAlertPayload, actorDiscordId: string) {
  const discordId = digitsOnly(payload.discordId || actorDiscordId || '');
  if (!discordId) return json({ ok: false, error: 'No Discord ID supplied or stored on your profile.' }, 400);
  const result = await sendDiscordDm(botToken, discordId, payload.title || 'MO8 MDT test', payload.message || 'Discord DM alerts are connected.');
  return json({ ok: result.ok, sent: result.ok, error: result.error || '', discordId }, result.ok ? 200 : 400);
}

async function sendDiscordDm(botToken: string, discordId: string, title: string, message: string, destinationUrl = '', destinationLabel = '') {
  const channelResponse = await discordFetch(botToken, 'https://discord.com/api/v10/users/@me/channels', {
    method: 'POST',
    body: JSON.stringify({ recipient_id: discordId }),
  });
  const channelBody = await channelResponse.json().catch(() => ({}));
  if (!channelResponse.ok) return { ok: false, error: discordError(channelResponse, channelBody) };

  const messageResponse = await discordFetch(botToken, `https://discord.com/api/v10/channels/${channelBody.id}/messages`, {
    method: 'POST',
    body: JSON.stringify({
      embeds: [formatDiscordEmbed(title, message, destinationUrl, destinationLabel)],
    }),
  });
  const messageBody = await messageResponse.json().catch(() => ({}));
  if (!messageResponse.ok) return { ok: false, error: discordError(messageResponse, messageBody) };
  return { ok: true };
}

function discordFetch(botToken: string, url: string, init: RequestInit = {}) {
  return fetch(url, {
    ...init,
    headers: {
      'Authorization': `Bot ${botToken}`,
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
  });
}

function formatDiscordEmbed(title: string, message: string, destinationUrl = '', destinationLabel = '') {
  const mdtUrl = Deno.env.get('MDT_URL') || 'https://harrytab.github.io/MO8-MDT/';
  const parsed = parseDiscordDetails(message);
  const description = parsed.description || 'You have a new MDT notification.';
  return {
    title: String(title || 'MO8 MDT notification').slice(0, 256),
    description: description.slice(0, 3900),
    color: embedColour(title, message),
    timestamp: new Date().toISOString(),
    fields: [
      ...parsed.fields,
      {
        name: destinationLabel || 'Open MDT',
        value: `[${destinationLabel || 'Launch MO8 MDT'}](${destinationUrl || mdtUrl})`,
        inline: false,
      },
    ],
    footer: {
      text: 'MO8 MDT Alerts',
    },
  };
}

function parseDiscordDetails(message: string) {
  const fields: Array<{ name: string; value: string; inline: boolean }> = [];
  const descriptionLines: string[] = [];

  String(message || '').split('\n').forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) return;
    const separator = trimmed.indexOf(':');
    if (separator > 0 && fields.length < 12) {
      const name = trimmed.slice(0, separator).trim();
      const value = trimmed.slice(separator + 1).trim();
      if (name && value) {
        fields.push({
          name: name.slice(0, 256),
          value: value.slice(0, 1024),
          inline: shouldInlineField(name),
        });
        return;
      }
    }
    descriptionLines.push(trimmed);
  });

  return {
    description: descriptionLines.join('\n'),
    fields,
  };
}

function shouldInlineField(name: string) {
  return ['Officer', 'Rank', 'Status', 'Outcome', 'Start date', 'End date', 'Date', 'Location', 'Supervisor', 'Reviewed by', 'Updated by', 'Training', 'Course', 'Standard'].includes(name);
}

function embedColour(title: string, message: string) {
  const text = `${title || ''} ${message || ''}`.toLowerCase();
  if (['denied', 'cancelled', 'canceled', 'failed', 'unsuccessful', 'discipline', 'disciplinary', 'removed', 'suspended'].some((word) => text.includes(word))) return 0xd93025;
  if (['supervisor updated', 'new supervisee assigned', 'supervisee reassigned', 'assigned as your supervisor'].some((word) => text.includes(word))) return 0x1267d8;
  if (['approved', 'passed', 'completed', 'successful', 'added'].some((word) => text.includes(word))) return 0x188038;
  if (['waitlist', 'waitlisted', 'deferred', 'on hold', 'shortlisted', 'interview', 'under review'].some((word) => text.includes(word))) return 0xf9ab00;
  return 0x1267d8;
}

function discordError(response: Response, body: Record<string, unknown>) {
  return `Discord API ${response.status}: ${body.message || response.statusText || 'Request failed'}${body.code ? ` (${body.code})` : ''}`;
}

function digitsOnly(value: unknown) {
  return String(value || '').replace(/\D/g, '');
}

function json(body: Record<string, unknown>, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
  });
}
