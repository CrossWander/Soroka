// Jira: issue keys in the chats and the cards open in Jira, with the ticket's name and status beside them.
import { BrowserWindow, net, safeStorage, session, shell } from 'electron';
import fs from 'node:fs';
import * as teams from './teams';
import { file, log, errorText, handle, settings } from './base';
import { t } from './i18n';
import { APP_ICON, status, settingsView, toPanel } from './main';

// Issue keys found in messages open in Jira; nothing else is opened from here.
// A web page in a window of the app, with Teams' session: Microsoft 365 opens signed in, and so does Jira once signed in there.
function webWindow(url: string) {
  const win = new BrowserWindow({ width: 1200, height: 820, icon: APP_ICON, autoHideMenuBar: true, webPreferences: { partition: 'persist:teams' } });
  void win.loadURL(url);
  return win;
}
const jiraBase = () => settings.jiraUrl.replace(/\/+$/, '');
handle('open:jira', (_e, key: string) => openJira(String(key)));
function openJira(key: string) { // where the links of Teams open: here or in the browser
  if (!settings.jiraUrl || !teams.JIRA_KEY.test(key)) return;
  const url = `${jiraBase()}/browse/${key}`;
  if (settings.links === 'browser') void shell.openExternal(url); else webWindow(url);
}
// A ticket's name and status beside its key in the cards, read as the user: signed in to Jira once in a window of the app,
// its cookies stay in the app's session as in a browser. { login: true } while not signed in; null when Jira does not say.
const jiraSeen = new Map<string, { at: number; issue: { summary: string; status: string } }>(); // 5 minutes
handle('jira:issue', (_e, key: string) => jiraIssue(String(key)));
async function jiraIssue(key: string) {
  if (!settings.jiraUrl || !teams.JIRA_KEY.test(key)) return null;
  const hit = jiraSeen.get(key);
  if (hit && Date.now() - hit.at < 5 * 60_000) return hit.issue;
  const r = await jiraGet(`/rest/api/2/issue/${key}?fields=summary,status`);
  if (r?.status === 401) return { login: true };
  if (!r?.ok) return null; // no such ticket, no access to it, or Jira away
  const j = await r.json().catch(() => undefined);
  const issue = { summary: String(j?.fields?.summary ?? ''), status: String(j?.fields?.status?.name ?? '') };
  jiraSeen.set(key, { at: Date.now(), issue });
  return issue;
}
// The same after the Jira links and keys in Teams' chats; a click on a key opens it as a link would. Not signed in: a hint
// after the links, asked again in a minute (a click on the link opens Jira in a window of the app, where the user signs in).
// ponytail: re-asked on the page's next change, not on the sign-in itself
export const jiraLinks = () => teams.watchJira(jiraBase(), keys => {
  for (const key of keys) void jiraIssue(key).then(i => i && 'login' in i ? teams.jiraText(key, t('jira.login'), 60_000, false)
    : teams.jiraText(key, i ? [i.status, i.summary].filter(Boolean).join(' · ') : '', 5 * 60_000, !!i)).catch(e => log('jira links:', errorText(e)));
}, key => openJira(key));
// A GET to Jira as the user: with the saved token, when it was saved for this address and Jira still takes it; else with
// the cookies of a sign-in in a window of the app. The token never follows a redirect: one fails the request.
async function jiraGet(apiPath: string) {
  const base = jiraBase(), token = jiraToken(base);
  if (token) {
    const r = await net.fetch(`${base}${apiPath}`, { headers: { Accept: 'application/json', Authorization: `Bearer ${token}` }, credentials: 'omit', redirect: 'error' })
      .catch(() => undefined);
    jiraTokenBad = r?.status === 401;
    if (r && !jiraTokenBad) return r;
  }
  return session.fromPartition('persist:teams').fetch(`${base}${apiPath}`, { headers: { Accept: 'application/json' }, credentials: 'include' }) // the session's cookies
    .catch(() => undefined);
}
// The Jira personal access token: encrypted by Windows for the user's account on this computer (safeStorage: DPAPI), in a file
// of its own; never in the settings, never in plain text, never shown again. Bound to the Jira address it was saved for and
// sent only there. It leaves this process only in requests to that address: the settings page learns whether it works, and whose it is.
const JIRA_TOKEN = file('jira-token.bin');
let jiraTokenBad = false; // Jira turned the saved token down (expired, revoked): the cookies are tried instead
function jiraSaved(): { base: string; token: string; user: string } | undefined {
  try { return JSON.parse(safeStorage.decryptString(fs.readFileSync(JIRA_TOKEN))); } catch { return undefined; }
}
const jiraToken = (base: string) => { const s = jiraSaved(); return s?.base === base ? s.token : undefined; };
handle('jira:token', () => {
  const s = jiraSaved();
  return s ? { user: s.user, other: s.base !== jiraBase(), bad: jiraTokenBad } : null;
});
// Saves a token Jira takes for the address in the settings ('' removes it). Only from the settings page.
handle('jira:tokenSet', async (e, value: unknown) => {
  if (e.sender !== settingsView?.webContents) return;
  if (value === '') fs.rmSync(JIRA_TOKEN, { force: true });
  else {
    const base = jiraBase(), token = String(value).trim();
    if (!base) throw new Error(t('set.jiraTokenNoUrl'));
    if (!/^[\x21-\x7e]{10,500}$/.test(token)) throw new Error(t('set.jiraTokenBad')); // printable, no spaces: nothing to break a header
    if (!safeStorage.isEncryptionAvailable()) throw new Error(t('set.jiraTokenNoCrypt')); // never kept in plain text instead
    const r = await net.fetch(`${base}/rest/api/2/myself`, { headers: { Accept: 'application/json', Authorization: `Bearer ${token}` }, credentials: 'omit', redirect: 'error' })
      .catch(() => undefined);
    if (r?.status === 401 || r?.status === 403) throw new Error(t('set.jiraTokenBad'));
    if (!r?.ok) throw new Error(t('set.jiraTokenAway', { status: String(r?.status ?? '—') }));
    const me = await r.json().catch(() => ({}));
    fs.writeFileSync(JIRA_TOKEN, safeStorage.encryptString(JSON.stringify({ base, token, user: String(me.displayName || me.name || '') })), { mode: 0o600 });
  }
  log(value === '' ? 'jira token: removed' : 'jira token: saved'); // never the token itself
  jiraTokenBad = false;
  jiraSeen.clear();
  if (status === 'ready') void jiraLinks().catch(err => log('jira links:', errorText(err))); // the chat's links ask again
});
handle('jira:login', () => { // always here, whatever the links setting: the cards read Jira with the app's session
  if (settings.jiraUrl) webWindow(`${jiraBase()}/login.jsp`).on('closed', () => toPanel('drafts')); // the cards ask Jira again
});
