const client = window.supabase.createClient(window.MO8_SUPABASE.url, window.MO8_SUPABASE.anonKey);
const SESSION_KEY = 'mo8_recruitment_session';
const state = { vacancies: [], vacancy: null, sessionToken: localStorage.getItem(SESSION_KEY) || '', applicant: null };
const $ = (selector) => document.querySelector(selector);

document.querySelectorAll('[data-page]').forEach((node) => node.addEventListener('click', () => showPage(node.dataset.page)));
$('#menuButton').addEventListener('click', () => $('.site-header').classList.toggle('menu-open'));
$('#vacancySearch').addEventListener('input', renderVacancies);
$('#searchButton').addEventListener('click', renderVacancies);
$('#applicantLoginForm').addEventListener('submit', loginApplicant);
$('#applicantRegisterForm').addEventListener('submit', registerApplicant);
$('#applicantLogout').addEventListener('click', logoutApplicant);
$('#applicationForm').addEventListener('submit', submitApplication);
$('#applicationCloseButton').addEventListener('click', closeApplication);
$('#applicationFields').addEventListener('input', scheduleApplicationDraftSave);
$('#applicationDialog').addEventListener('cancel', (event) => { event.preventDefault(); closeApplication(); });
document.addEventListener('click', async (event) => {
  const vacancy = event.target.closest('[data-vacancy]');
  if (vacancy) await openVacancy(vacancy.dataset.vacancy);
  const apply = event.target.closest('[data-apply]');
  if (apply) await openApplication(apply.dataset.apply);
});

function showPage(page) {
  ['jobs', 'vacancy', 'account'].forEach((name) => $(`#${name}Page`).hidden = name !== page);
  document.querySelectorAll('.nav-link').forEach((node) => node.classList.toggle('active', node.dataset.page === page));
  $('.site-header').classList.remove('menu-open');
  if (page === 'account') loadApplicant();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

async function loadVacancies() {
  const { data, error } = await client.rpc('public_recruitment_vacancies');
  if (error) return setVacancyStatus(error.message);
  state.vacancies = data || [];
  renderVacancies();
}

function renderVacancies() {
  const query = $('#vacancySearch').value.trim().toLowerCase();
  const rows = state.vacancies.filter((row) => !query || [row.title, row.team, row.location, row.summary].some((value) => String(value || '').toLowerCase().includes(query)));
  setVacancyStatus(`${rows.length} result${rows.length === 1 ? '' : 's'} matched`);
  $('#vacancyList').innerHTML = rows.length ? `<div class="vacancy-head"><span>Title</span><span>Closing date</span></div>${rows.map((row) => `<button class="vacancy-row" data-vacancy="${escapeHtml(row.vacancy_id)}"><span><strong>${escapeHtml(row.title)}</strong><small>${escapeHtml(row.team)} / ${escapeHtml(row.location)}</small></span><time>${row.closes_at ? formatDateTime(row.closes_at) : 'Open until filled'}</time></button>`).join('')}` : '<p class="empty">No current vacancies match your search.</p>';
}

function setVacancyStatus(message) { $('#vacancyStatus').textContent = message; }

async function openVacancy(id) {
  const { data, error } = await client.rpc('public_recruitment_vacancy', { target_vacancy_id: id });
  if (error || !data?.vacancy) return alert(error?.message || 'This vacancy is no longer available.');
  state.vacancy = { ...data.vacancy, fields: data.fields || [] };
  const row = state.vacancy;
  $('#vacancyDetail').innerHTML = `<header><span>${escapeHtml(row.team)} / ${escapeHtml(row.vacancy_type)}</span><h1>${escapeHtml(row.title)}</h1><p>${escapeHtml(row.summary || '')}</p></header><div class="vacancy-meta"><div><span>Location</span><strong>${escapeHtml(row.location)}</strong></div><div><span>Positions</span><strong>${escapeHtml(String(row.positions))}</strong></div><div><span>Applications close</span><strong>${row.closes_at ? formatDateTime(row.closes_at) : 'Open until filled'}</strong></div><div><span>Reference</span><strong>${escapeHtml(row.vacancy_id)}</strong></div></div><div class="vacancy-copy">${escapeHtml(row.description || row.summary || 'No additional details supplied.')}</div><section class="apply-panel"><div><strong>Ready to apply?</strong><p>A recruitment account is required to submit and track an external application.</p></div><button data-apply="${escapeHtml(row.vacancy_id)}">Apply for this role</button></section>`;
  showPage('vacancy');
}

async function loadApplicant() {
  if (!state.sessionToken) return renderApplicantState();
  const { data, error } = await client.rpc('recruitment_session', { session_token_input: state.sessionToken });
  if (error || !data) return logoutApplicant();
  state.applicant = data;
  renderApplicantState();
  const result = await client.rpc('recruitment_my_applications', { session_token_input: state.sessionToken });
  $('#applicationList').innerHTML = result.error ? `<p class="empty">${escapeHtml(result.error.message)}</p>` : renderApplicationCards(result.data || []);
}

function renderApplicantState() {
  $('#applicantSignedOut').hidden = Boolean(state.applicant);
  $('#applicantSignedIn').hidden = !state.applicant;
  if (state.applicant) $('#applicantName').textContent = state.applicant.username;
}

function renderApplicationCards(rows) {
  return rows.length ? rows.map((row) => `<article class="application-card status-${escapeHtml(row.status.toLowerCase().replaceAll(' ', '-'))}"><div><h3>${escapeHtml(row.vacancy_title)}</h3><p>${escapeHtml(row.applicant_message || 'No update has been provided by the recruitment team.')}</p><small>Submitted ${formatDateTime(row.submitted_at)} / ${escapeHtml(row.application_id)}</small></div><span>${escapeHtml(row.status)}</span></article>`).join('') : '<p class="empty">You have not submitted any applications yet.</p>';
}

async function loginApplicant(event) {
  event.preventDefault();
  const form = event.currentTarget; const status = form.querySelector('.form-status'); const values = Object.fromEntries(new FormData(form));
  status.textContent = 'Signing in...';
  const { data, error } = await client.rpc('recruitment_login', { applicant_username: values.username, applicant_password: values.password });
  if (error) return status.textContent = cleanError(error.message);
  setSession(data); form.reset(); status.textContent = ''; await loadApplicant();
}

async function registerApplicant(event) {
  event.preventDefault();
  const form = event.currentTarget; const status = form.querySelector('.form-status'); const values = Object.fromEntries(new FormData(form));
  status.textContent = 'Creating account...';
  const { data, error } = await client.rpc('recruitment_register', { applicant_username: values.username, applicant_discord_id: values.discordId, applicant_password: values.password });
  if (error) return status.textContent = cleanError(error.message);
  setSession(data); form.reset(); status.classList.add('success'); status.textContent = 'Account created.'; await loadApplicant();
}

function setSession(data) { state.sessionToken = data.token; state.applicant = { username: data.username, discordId: data.discordId }; localStorage.setItem(SESSION_KEY, data.token); }
function logoutApplicant() { localStorage.removeItem(SESSION_KEY); state.sessionToken = ''; state.applicant = null; renderApplicantState(); }

async function openApplication(vacancyId) {
  if (!state.sessionToken) { showPage('account'); $('#applicantLoginForm input').focus(); return; }
  if (!state.applicant) await loadApplicant();
  if (!state.applicant) return;
  if (!state.vacancy || state.vacancy.vacancy_id !== vacancyId) await openVacancy(vacancyId);
  const row = state.vacancy;
  $('#applicationTitle').textContent = row.title;
  $('#applicationFields').innerHTML = `<aside class="applicant-identity"><span>Applying as</span><strong>${escapeHtml(state.applicant.username)}</strong><small>Discord ID ${escapeHtml(state.applicant.discordId)} / supplied automatically from your recruitment account</small></aside>${(row.fields || []).map(renderField).join('') || '<p>No additional questions are required for this role.</p>'}`;
  const restored = restoreApplicationDraft();
  $('#applicationStatus').classList.toggle('success', restored);
  $('#applicationStatus').textContent = restored ? 'Your saved draft has been restored.' : '';
  $('#applicationDialog').showModal();
}

function renderField(field) {
  const required = field.required ? ' required' : ''; const help = field.help_text ? `<small class="field-help">${escapeHtml(field.help_text)}</small>` : '';
  if (field.field_type === 'Long text') return `<label>${escapeHtml(field.label)}${help}<textarea name="${escapeHtml(field.field_key)}"${required}></textarea></label>`;
  if (field.field_type === 'Yes / No') return `<label>${escapeHtml(field.label)}${help}<select name="${escapeHtml(field.field_key)}"${required}><option value="">Select</option><option>Yes</option><option>No</option></select></label>`;
  if (field.field_type === 'Single choice' || field.field_type === 'Multiple choice') { const inputType = field.field_type === 'Single choice' ? 'radio' : 'checkbox'; return `<fieldset><legend>${escapeHtml(field.label)}${field.required ? ' *' : ''}</legend>${help}${(field.options || []).map((option) => `<label><input type="${inputType}" name="${escapeHtml(field.field_key)}" value="${escapeHtml(option)}"${required}> ${escapeHtml(option)}</label>`).join('')}</fieldset>`; }
  const inputType = field.field_type === 'Date' ? 'date' : field.field_type === 'Number' ? 'number' : 'text';
  return `<label>${escapeHtml(field.label)}${help}<input type="${inputType}" name="${escapeHtml(field.field_key)}"${required}></label>`;
}

async function submitApplication(event) {
  event.preventDefault();
  if (event.submitter?.value !== 'submit') return;
  const form = event.currentTarget;
  if (!form.reportValidity()) return;
  const answers = collectApplicationAnswers();
  $('#applicationStatus').textContent = 'Submitting application...';
  const { data, error } = await client.rpc('submit_recruitment_application', { session_token_input: state.sessionToken, target_vacancy_id: state.vacancy.vacancy_id, application_answers: answers });
  if (error) return $('#applicationStatus').textContent = cleanError(error.message);
  await sendApplicantDiscord('recruitmentSubmitted', data.applicationId);
  localStorage.removeItem(applicationDraftKey());
  $('#applicationStatus').classList.add('success'); $('#applicationStatus').textContent = `Application ${data.applicationId} submitted successfully.`;
  setTimeout(() => { $('#applicationDialog').close(); showPage('account'); loadApplicant(); }, 900);
}

let applicationDraftTimer = null;

function collectApplicationAnswers() {
  const answers = {};
  new FormData($('#applicationForm')).forEach((value, key) => {
    if (key === 'submit') return;
    if (Object.prototype.hasOwnProperty.call(answers, key)) answers[key] = Array.isArray(answers[key]) ? [...answers[key], value] : [answers[key], value];
    else answers[key] = value;
  });
  return answers;
}

function applicationDraftKey() {
  const username = String(state.applicant?.username || 'anonymous').trim().toLowerCase();
  return `mo8_recruitment_draft:${username}:${state.vacancy?.vacancy_id || 'unknown'}`;
}

function scheduleApplicationDraftSave() {
  clearTimeout(applicationDraftTimer);
  applicationDraftTimer = setTimeout(saveApplicationDraft, 250);
}

function saveApplicationDraft() {
  if (!state.applicant || !state.vacancy) return;
  localStorage.setItem(applicationDraftKey(), JSON.stringify({ answers: collectApplicationAnswers(), savedAt: new Date().toISOString() }));
  const status = $('#applicationStatus');
  if (status && !status.textContent.includes('Submitting')) { status.classList.add('success'); status.textContent = 'Draft saved on this device.'; }
}

function restoreApplicationDraft() {
  try {
    const draft = JSON.parse(localStorage.getItem(applicationDraftKey()) || 'null');
    if (!draft?.answers) return false;
    Object.entries(draft.answers).forEach(([name, saved]) => {
      const values = Array.isArray(saved) ? saved.map(String) : [String(saved)];
      [...$('#applicationForm').elements].filter((element) => element.name === name).forEach((element) => {
        if (['checkbox', 'radio'].includes(element.type)) element.checked = values.includes(element.value);
        else element.value = values[0] || '';
      });
    });
    return true;
  } catch (_) { return false; }
}

function closeApplication() {
  clearTimeout(applicationDraftTimer);
  saveApplicationDraft();
  $('#applicationDialog').close();
}

async function sendApplicantDiscord(action, applicationId) { try { await client.functions.invoke('discord-alerts', { body: { action, applicationId, recruitmentToken: state.sessionToken } }); } catch (_) {} }
function formatDateTime(value) { return new Intl.DateTimeFormat('en-GB', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Europe/London' }).format(new Date(value)); }
function cleanError(value) { return String(value || 'Something went wrong.').replace(/^.*?: /, ''); }
function escapeHtml(value) { return String(value ?? '').replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[char]); }

loadVacancies();
loadApplicant();
