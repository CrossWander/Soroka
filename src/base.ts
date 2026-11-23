// Soroka's base, shared by the rest of the app: its data folder and log, the guard of its own IPC, the settings,
// the cards' state, the local API's token and the queue of the Teams page.
import { app, ipcMain, nativeTheme } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import * as ai from './ai';
import * as i18n from './i18n';

// Until 0.9.1 the app was called TeamsBridge: its data (Teams sign-in, settings, token) moves to the new name once.
// While the old app still runs and holds the folder, this run keeps using it and the move is tried at the next start.
const legacy = path.join(app.getPath('appData'), 'TeamsBridge');
if (fs.existsSync(path.join(legacy, 'settings.json')) && !fs.existsSync(path.join(app.getPath('userData'), 'settings.json'))) {
  try { fs.rmdirSync(app.getPath('userData')); } catch { /* not there yet, or not empty */ }
  try { fs.renameSync(legacy, app.getPath('userData')); } catch { app.setPath('userData', legacy); }
}
if (!app.requestSingleInstanceLock()) app.exit(0);
app.disableHardwareAcceleration(); // nothing here needs the GPU process
app.commandLine.appendSwitch('lang', 'ru-RU'); // the app's own Teams UI language; selectors don't depend on it

export const data = app.getPath('userData');
export const file = (name: string) => path.join(data, name);
export const ui = (name: string) => path.join(__dirname, '..', 'ui', name);
// A page of the app's own: a file of its ui folder, not any file: page (an HTML file dropped on a window goes there too)
const UI_DIR = (path.join(__dirname, '..', 'ui') + path.sep).toLowerCase();
export const ownPage = (url = '') => { try { return fileURLToPath(url).toLowerCase().startsWith(UI_DIR); } catch { return false; } }; // not a file: URL
try { if (!fs.existsSync(file('soroka.log'))) fs.renameSync(file('bridge.log'), file('soroka.log')); } catch { /* no log of the old name */ } // bridge.log until 0.9.5
// a log past 2 MB starts over at the app's start; the one before stays as *.old.log
for (const name of ['soroka.log', 'events.log']) try { if (fs.statSync(file(name)).size > 2 << 20) fs.renameSync(file(name), file(name.replace('.log', '.old.log'))); } catch { /* none yet */ }
export const log = (...parts: unknown[]) => fs.appendFileSync(file('soroka.log'), `${new Date().toISOString()} ${parts.join(' ')}\n`);
export const readJson = <T>(name: string, fallback: T): T => { try { return { ...fallback, ...JSON.parse(fs.readFileSync(file(name), 'utf8')) }; } catch { return { ...fallback }; } };
// a new file, then a rename: a crash in the middle leaves the old one whole, not half a file
export const writeJson = (name: string, value: unknown) => {
  const tmp = file(`${name}.tmp`);
  fs.writeFileSync(tmp, JSON.stringify(value, null, 2));
  try { fs.renameSync(tmp, file(name)); } catch { fs.writeFileSync(file(name), fs.readFileSync(tmp)); fs.rmSync(tmp, { force: true }); } // held open by another program
};
export const errorText = (e: unknown) => (e as Error).message;
// The app's own pages only (ownPage): a web page in one of its windows never reaches these, wherever it navigated.
export const handle = (channel: string, fn: (e: Electron.IpcMainInvokeEvent, ...a: any[]) => unknown) =>
  ipcMain.handle(channel, (e, ...a) => { if (!ownPage(e.senderFrame?.url)) throw new Error('not allowed'); return fn(e, ...a); });
// What fails also goes to soroka.log (chat names and errors, never message texts): a run that failed unseen can be traced.
export const logged = <A extends unknown[], T>(what: string, fn: (...a: A) => T | Promise<T>) => async (...a: A) => {
  try { return await fn(...a); } catch (e) { log(`${what}:`, errorText(e)); throw e; }
};

// ---------- settings & state ----------
export type Settings = { wakeMode: string; draftModel: string; summaryModel: string; label: boolean; autostart: boolean; summaryBackground: boolean; style: string; language: string;
  onlyFrom: string; keywords: string; quietFrom: string; quietTo: string; digestAt: string; jiraUrl: string;
  muteFrom: string; groups: boolean; quietDnd: boolean; tones: string; remindAfter: string; followUpAfter: string; quickReplies: string; triage: boolean; links: string; officeMenu: boolean; activity: boolean; theme: string; aiSkip: string; webChat: string; aiOff: boolean;
  agent: string; projectDir: string; projectAsk: boolean; handoffKeep: string }; // activity: mentions in channels, from Teams' Activity; theme: '' = the system's
// aiOff: no AI at all — the app only notifies; no AI program is started
// agent: the user's AI agent a card's chat goes to (agents.ts), '' = none chosen; projectDir: the folder it opens in, or
// projectAsk: a folder picked each time; handoffKeep: hours the files handed over to it stay
// wakeMode: what the AI does with a new message; the notification itself always comes
// remindAfter: minutes a personal message may wait for an answer before a reminder; '' = never
// followUpAfter: the same for the user's own question that nobody answered
// webChat: the site "Open in AI chat" opens; '' = the one of the model's provider, else claude.ai
// links: where web links from Teams and the cards open; '' = a window of the app, 'browser' = the default browser
// officeMenu: the button of Teams' Office menu (the dotted grid top left); false hides it
export const OPTIONS: Record<string, unknown[]> = { wakeMode: ['notify', 'drafts', 'monitor'], language: i18n.LANGS, remindAfter: ['', '60', '240', '1440'],
  followUpAfter: ['', '60', '240', '1440'], theme: ['', 'light', 'dark'], links: ['', 'browser'],
  webChat: ['', ...Object.keys(ai.WEB)], handoffKeep: ['1', '24', '168'] };
export const firstRun = !fs.existsSync(file('settings.json'));
export const settings = readJson<Settings>('settings.json', { wakeMode: 'notify', draftModel: 'claude:claude-haiku-5-5', summaryModel: 'claude:claude-opus-5-5',
  label: true, autostart: false, summaryBackground: true, style: '', language: '', // language '' = the system's, resolved at start
  onlyFrom: '', keywords: '', quietFrom: '', quietTo: '', digestAt: '', jiraUrl: '',
  muteFrom: '', groups: false, quietDnd: true, tones: '', remindAfter: '240', followUpAfter: '240', quickReplies: '', triage: false, links: '', officeMenu: true, activity: true, theme: '', aiSkip: '', webChat: '', aiOff: false,
  agent: '', projectDir: '', projectAsk: false, handoffKeep: '24' }); // a day: a request left in the agent's field still finds its chat.md
// The theme of every window, Teams' page included: its pages follow prefers-color-scheme, which nativeTheme sets.
nativeTheme.themeSource = (settings.theme || 'system') as 'system' | 'light' | 'dark';
settings.draftModel = ({ haiku: 'claude-haiku-5-5', sonnet: 'claude-sonnet-5-5' } as Record<string, string>)[settings.draftModel] ?? settings.draftModel; // aliases saved before 0.4.4
for (const k of ['draftModel', 'summaryModel'] as const) if (!settings[k].includes(':')) settings[k] = `claude:${settings[k]}`; // before 0.10 every model was Claude's
if (!OPTIONS.wakeMode.includes(settings.wakeMode)) settings.wakeMode = 'notify'; // "off" and "loop" before 0.5.3: notifications no longer switch off

// kind, what the card is for: incoming — a message to the user; fyi — a group's message that only "Every message in group
// chats" brought (no AI reply unasked, no reminder); request — made by "Suggest reply" or a Claude session on any chat;
// improve — "Improve my text": `last` is the user's own unsent text from the compose box, and sending replaces it; nudge — a
// reminder about the user's own question nobody answered (`last`), closed when an answer comes. A new message in the chat
// makes any but improve an incoming card.
// remindAt: snoozed until then; reminded: the "still waiting for your answer" reminder went out; answered: the user replied
// in Teams itself; read: read it there
// urgent: an @mention or a keyword hit, never held back; held: "AI sorts" found it waits for nothing from the user
// hushed: came in quiet hours, announced when they end; preview: the chat's last message when a card made by request was
// first seen in the list (a new one of the user's own closes the card)
// kept: `text` is the user's own, typed in a card or in "Replies" (draft:keep): it stays over the chat's next message and goes
// without the AI label
export type Draft = { chatId: string; title: string; note: string; last: string; text: string; variants?: string[]; busy?: boolean; error?: string;
  context?: string[]; kind: 'incoming' | 'fyi' | 'request' | 'improve' | 'nudge'; status: 'pending' | 'sent' | 'rejected' | 'answered' | 'read'; at: number;
  remindAt?: number; reminded?: boolean; urgent?: boolean; held?: boolean; hushed?: boolean; preview?: string; msgs?: Msg[]; kept?: boolean;
  from?: 'ai'; back?: boolean; failed?: boolean; ready?: number };
// from: an AI session's draft (POST /draft); back: came back by a reminder (its snooze ended, or "still no answer");
// failed: its send failed; ready: when a result the user asked for (variants, improved text) came, or failed (src/badge.ts)
// msgs: the messages of an incoming card, as they came one by one (the chat list's preview, or the opened chat's text in
// drafts mode); who: their author in a group; me: the message the card came for in a group (the mention, the keyword)
export type Msg = { who?: string; text: string; time: string; me?: boolean };
// seen: per chat, the last message announced ("time|preview"); models: last result of "Check models"; digestDay: date of the last daily digest
export const state = readJson('state.json', { nextDraft: 1, drafts: {} as Record<string, Draft>, seen: {} as Record<string, string>,
  models: {} as Record<string, string>, digestDay: '', summaries: {} as Record<string, Summary & { at: number }>, // summaries: the last one per chat
  spent: { day: '', tokens: 0 }, // spent: Claude's tokens today, as Claude Code reports them
  activitySeen: [] as string[], reactions: [] as string[], // activitySeen: Activity items already looked at; reactions: the user's recent ones
  asked: {} as Record<string, { title: string; text: string; at: number; told?: boolean }>, // asked: per chat, the user's own question still unanswered
  checkWarned: '', // checkWarned: "version|missing parts" the self-check last warned about
  toolsOk: [] as string[], // toolsOk: the AIs with tools of their own the user let read chats
  callBounds: null as Electron.Rectangle | null, // callBounds: where the user left the call window
  miniAt: null as { x: number; y: number } | null, // miniAt: where the user left the call's compact view
  miniMode: '', // miniMode: 'bar' when the user folded the compact view to its buttons
  lastDir: '' }); // lastDir: the project folder last picked for an agent, where the next pick starts
state.models = Object.fromEntries(Object.entries(state.models).map(([m, s]) => [m.includes(':') ? m : `claude:${m}`, s])); // before 0.10: Claude's only
ai.setup(data);
ai.usage.onSpent = tokens => {
  const day = new Date().toDateString();
  if (state.spent.day !== day) state.spent = { day, tokens: 0 };
  state.spent.tokens += tokens;
  saveState();
};
for (const d of Object.values(state.drafts)) delete d.busy; // a run cut short by a restart
// cards saved before kind: its four flags (an answer joined to a nudge had both incoming and nudge: it is incoming)
for (const d of Object.values(state.drafts) as (Draft & Partial<Record<'incoming' | 'fyi' | 'improve' | 'nudge', boolean>>)[]) if (!d.kind) {
  d.kind = d.improve ? 'improve' : d.incoming ? (d.fyi ? 'fyi' : 'incoming') : d.nudge ? 'nudge' : 'request';
  delete d.incoming; delete d.fyi; delete d.improve; delete d.nudge;
}
// the state stays small: handled cards a week, the last message seen of the 300 latest chats
for (const [n, d] of Object.entries(state.drafts)) if (d.status !== 'pending' && Date.now() - d.at > 7 * 86_400_000) delete state.drafts[n];
state.seen = Object.fromEntries(Object.entries(state.seen).slice(-300));
export const saveState = () => writeJson('state.json', state);
export const waiting = () => Object.entries(state.drafts).filter(([, d]) => d.status === 'pending'); // snoozed ones too
// The cards in the answers window: a snoozed one stays out until its time
export const pending = () => waiting().filter(([, d]) => !(d.remindAt && d.remindAt > Date.now()))
  .map(([n, { context, ...d }]) => ({ n, ...d })); // the transcript stays in the main process

// marks: for one chat, what its messages carry for a look into code ("Look into it in the agent" stands out then)
export type Summary = { text: string; count: number; from: string; marks?: string[] };

const tokenFile = file('token.txt');
// a missing or short token (hand-made) is replaced: 48 hex characters
if (!fs.existsSync(tokenFile) || fs.readFileSync(tokenFile, 'utf8').trim().length < 32) { fs.mkdirSync(data, { recursive: true }); fs.writeFileSync(tokenFile, crypto.randomBytes(24).toString('hex')); }
export const TOKEN = fs.readFileSync(tokenFile, 'utf8').trim();

// One page = one thing at a time.
let chain: Promise<unknown> = Promise.resolve();
export const serial = <T>(fn: () => Promise<T>): Promise<T> => { const r = chain.then(fn); chain = r.catch(() => {}); return r; };
