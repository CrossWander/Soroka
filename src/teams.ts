// Drives the Teams web page in the hidden window: DOM reads via executeJavaScript, real input via the DevTools protocol.
// Selectors use Teams' own data-tid / ids and image names, not UI texts, so any Teams language works.
import type { WebContents } from 'electron';

export const TEAMS_URL = 'https://teams.cloud.microsoft/';
// preview: the last message as the chat list shows it ("Author: text" in group chats, cut by Teams); time: "16:33" today, a date before
export type Chat = { id: string; title: string; preview: string; time: string; type: string; unread: boolean; mention: boolean };
// replyTo, quote: a quoted reply's author and text, kept out of the message's own text
export type Message = { id: string; author: string; mine: boolean; text: string; replyTo?: string; quote?: string };
type ChatRef = Pick<Chat, 'id' | 'title'>;

export const fail = (status: number, message: string) => Object.assign(new Error(message), { status });
const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));
// Teams may render emoji as images, so text checks compare without them.
const plain = (t: string) => t.replace(/[\p{Extended_Pictographic}️​-‍⁠﻿]/gu, '').replace(/\s+/g, ' ').trim();

let wc: WebContents;
const cdp = (method: string, params: object = {}) => wc.debugger.sendCommand(method, params);

export function attach(contents: WebContents) {
  wc = contents;
  wc.debugger.attach('1.3');
}
// Call once a page is loaded (hangs before that): the hidden window then takes keyboard input.
export const emulateFocus = () => cdp('Emulation.setFocusEmulationEnabled', { enabled: true });

// Runs fn inside the page. fn is serialized: it must not use anything from this module and must return plain data.
// A page that never loaded anything waits for a load to run a script, for ever: the fresh page a call leaves behind (popOutCall)
// loads in its turn in the app's queue, so a task already in it would wait for itself. There it fails at once.
const run = (fn: (arg: any) => unknown, arg: unknown = null): Promise<any> => !wc.getURL() ? Promise.reject(fail(503, 'Teams page not loaded yet'))
  : wc.executeJavaScript(`(${fn})(${JSON.stringify(arg)})`, true);

async function waitFor(fn: (arg: any) => unknown, arg: unknown = null, timeout = 15_000) {
  for (const end = Date.now() + timeout; ; await sleep(250)) {
    if (await run(fn, arg).catch(() => false)) return true; // fails while the page navigates
    if (Date.now() > end) return false;
  }
}
const exists = (sel: string, timeout?: number) => waitFor(s => !!document.querySelector(s), sel, timeout);
// Teams' left rail is there on every view (Chat, Activity, Teams), unlike the chat list; the loading screen keeps the focus.
export const ready = (timeout = 0) => waitFor(() => !!document.querySelector('[data-tid="app-bar-wrapper"]')
  && !document.getElementById('loading-screen')?.getClientRects().length, null, timeout);
// Teams' own app ids on the rail: Activity and Chat
const ACTIVITY = '[data-tid="14d6962d-6eeb-4f48-8890-de55454bb136"]', CHAT_APP = '[data-tid="86fcd49b-61a2-4701-b771-54728cd291fb"]';
// The page elements the app's features stand on, looked up without a click: after a Teams update a missing one means
// that feature has stopped working. null: cannot tell from the current view (no chat open, or Activity shown).
export const selfCheck = (): Promise<Record<string, boolean | null>> => run(() => {
  const has = (s: string) => !!document.querySelector(s);
  const listShown = has('[role="tree"]'), pane = document.querySelector('[data-tid="message-pane-list-viewport"]') as HTMLElement | null;
  const chatOpen = !!pane && pane.innerText.trim().length > 40; // an empty new chat has no messages to find
  return {
    rail: has('[data-tid="app-bar-wrapper"]'),
    chat: has('[data-tid="86fcd49b-61a2-4701-b771-54728cd291fb"]'),
    activity: has('[data-tid="14d6962d-6eeb-4f48-8890-de55454bb136"]'),
    list: listShown ? has('[role="treeitem"][data-item-type="chat"] [id^="title-chat-list-item_"]') : null,
    messages: chatOpen ? has('[data-tid="chat-pane-message"]') : null,
    compose: chatOpen ? has('[id^="new-message-"]') : null,
    menu: has('[data-tid="more-options-header"]'),
  };
});

// Back to the chat list from Activity or Teams (the user's, or the app's own visit to Activity).
export async function showChats() {
  if (await exists('[role="tree"]', 0)) return;
  await click(CHAT_APP);
  if (!(await exists('[role="tree"]', 10_000))) throw fail(502, 'the chat list did not open');
}

// Real mouse click in the middle of the index-th match (negative index counts from the end);
// `label` keeps only elements whose aria-label or text contains it.
async function click(sel: string, index = 0, label?: string) {
  const at = await run(({ sel, index, label }) => {
    const els = [...document.querySelectorAll(sel)]
      .filter(e => !label || (e.getAttribute('aria-label') || (e as HTMLElement).innerText || '').includes(label));
    const e = els.at(index);
    if (!e) return null;
    e.scrollIntoView({ block: 'center' });
    const r = e.getBoundingClientRect();
    return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
  }, { sel, index, label });
  if (!at) throw fail(404, `not found: ${sel}${label ? ` "${label}"` : ''}`);
  for (const type of ['mouseMoved', 'mousePressed', 'mouseReleased'])
    await cdp('Input.dispatchMouseEvent', { type, x: at.x, y: at.y, button: 'left', clickCount: 1 });
}

const KEYS: Record<string, { key: string; code: string; windowsVirtualKeyCode: number; text?: string }> = {
  Enter: { key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r' },
  Escape: { key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 },
  Tab: { key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9 },
  End: { key: 'End', code: 'End', windowsVirtualKeyCode: 35 },
  Delete: { key: 'Delete', code: 'Delete', windowsVirtualKeyCode: 46 },
  a: { key: 'a', code: 'KeyA', windowsVirtualKeyCode: 65 },
};
const CTRL = 2, SHIFT = 8;
async function press(name: string, modifiers = 0) {
  const { text, ...key } = KEYS[name];
  const typed = modifiers & CTRL ? undefined : text;
  await cdp('Input.dispatchKeyEvent', { type: typed ? 'keyDown' : 'rawKeyDown', modifiers, ...key, text: typed });
  await cdp('Input.dispatchKeyEvent', { type: 'keyUp', modifiers, ...key });
}
const insertText = (text: string) => cdp('Input.insertText', { text });
const typeChar = async (ch: string) => {
  await cdp('Input.dispatchKeyEvent', { type: 'keyDown', key: ch, text: ch, unmodifiedText: ch });
  await cdp('Input.dispatchKeyEvent', { type: 'keyUp', key: ch });
};

// Unicode emoji show up as plain font glyphs in Teams. Its own emoji come from the suggestion list that ":code" opens:
// [code to type, emoji id in the suggestion's image path]. Other emoji go in as Unicode.
const TEAMS_EMOJI: Record<string, [string, string]> = {
  '🤖': ['robot', 'smilerobot'], '😄': ['laugh', 'laugh'], '😆': ['laugh', 'laugh'], '😂': ['laugh', 'cwl'],
  '🤣': ['rofl', 'rofl'], '🙂': ['smile', 'smile'], '😊': ['smile', 'smileeyes'], '🤔': ['think', 'think'],
};
const EMOJI_RE = /(\p{Extended_Pictographic}️?)/u;
const SUGGESTION = '#autocomplete-picker-list [role="option"]';
const suggestionIndex = (id: string): Promise<number> => run(({ sel, id }) =>
  [...document.querySelectorAll(sel)].findIndex(o => (o.querySelector('img')?.getAttribute('src') || '').includes(`/emoticons/${id}/`)), { sel: SUGGESTION, id });
async function typeLine(line: string) {
  for (const part of line.split(EMOJI_RE).filter(Boolean)) {
    const emoji = TEAMS_EMOJI[part];
    if (!emoji) { await insertText(part); continue; }
    const [code, id] = emoji;
    for (const ch of `:${code}`) await typeChar(ch);
    for (const end = Date.now() + 6000; (await suggestionIndex(id)) < 0; await sleep(250))
      if (Date.now() > end) throw fail(500, `Teams emoji "${id}" not offered for :${code}, not sent`); // callers clear the box
    await click(SUGGESTION, await suggestionIndex(id)); // a click, never Enter: Enter without the list would send a half-typed message
    // the list stays ~0.9 s after the click, and while it is there Teams drops an Enter: the message would not go
    await waitFor(s => !document.querySelector(s), '#autocomplete-picker-list', 3000);
    await run(b => (document.querySelector(b) as HTMLElement | null)?.focus(), BOX);
  }
}

const TYPES: Record<string, string> = {
  OneGQL_OneOnOneChatConversation: 'direct', OneGQL_GroupChatConversation: 'group',
  OneGQL_MeetingChatConversation: 'meeting', OneGQL_SelfChatConversation: 'self',
};
// Unread and personal @mention come from aria-labelledby ids, which do not depend on the UI language.
export async function chatList(): Promise<Chat[]> {
  const raw: (Chat & { kind: string })[] = await run(() => [...document.querySelectorAll('[role="treeitem"][data-item-type="chat"]')].map(e => {
    // value is "<folder kind>|<folder id>/<chat kind>|<chat id>" inside a folder, "<chat kind>|<chat id>" outside
    const [kind, id] = (e.getAttribute('data-fui-tree-item-value') || '').split('/').pop()!.split('|');
    const by = (e.getAttribute('aria-labelledby') || '').split(/\s+/);
    // title / time / last message have their own ids: "<part>-chat-list-item_<chat id>"
    const part = (name: string) => ((e.querySelector(`[id^="${name}-chat-list-item_"]`) as HTMLElement | null)?.innerText || '').replace(/\s+/g, ' ').trim();
    return { id, kind, title: part('title'), preview: part('message-preview'), time: part('time'),
      unread: by.some(x => x.includes('chat_list_unread_text')), mention: by.some(x => x.includes('simple_collab_voiceover_personal_mention_string')) };
  }));
  const seen = new Set<string>(); // a favourite chat is listed twice
  return raw.filter(c => c.id && !seen.has(c.id) && seen.add(c.id)).map(({ kind, ...c }) => ({ ...c, type: TYPES[kind] ?? kind }));
}

// Calls onChange soon after the set of unread chats or their last messages changes: Teams updates the page itself,
// a MutationObserver in the page notices and calls a DevTools binding. Install again after every page load.
// Page → app calls go through DevTools bindings; one listener serves them all.
const bindings = new Map<string, (payload: string) => void>();
async function bind(name: string, fn: (payload: string) => void) {
  if (!wc.debugger.listenerCount('message'))
    wc.debugger.on('message', (_e, method, params) => { if (method === 'Runtime.bindingCalled') bindings.get(params.name)?.(params.payload); });
  bindings.set(name, fn);
  await cdp('Runtime.addBinding', { name });
}
const BINDING = 'tbChatsChanged';
export async function watchChats(onChange: () => void) {
  await bind(BINDING, () => onChange());
  await run(name => {
    const w = window as any;
    w.__tbWatch?.disconnect();
    const unread = () => [...document.querySelectorAll('[role="treeitem"][data-item-type="chat"]')]
      .filter(e => (e.getAttribute('aria-labelledby') || '').includes('chat_list_unread_text'))
      .map(e => `${e.getAttribute('data-fui-tree-item-value')} ${e.querySelector('[id^="message-preview-chat-list-item_"]')?.textContent}`).join('\n');
    let last = unread(), timer = 0;
    w.__tbWatch = new MutationObserver(() => { // Teams changes the page constantly: compare at most twice a second
      if (timer) return;
      timer = window.setTimeout(() => { timer = 0; const now = unread(); if (now !== last) { last = now; w[name]?.(''); } }, 500);
    });
    w.__tbWatch.observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['aria-labelledby'] });
  }, BINDING);
}

// ---------- calls: an incoming call, or a meeting that started, as Teams shows it in its page ----------
// Teams' toast for them has a decline button with a fixed id (reject-call-button / reject-meeting-button); the toast's
// other buttons (accept with video or audio, join) are taken as they are, with Teams' own labels, and numbered for ringClick.
// title/message: the toast's text lines that are not button labels (the caller first). outline: its data-tid ids, for the log.
export type Ring = { kind: 'call' | 'meeting'; title: string; message: string; buttons: { label: string; reject: boolean }[]; outline: string };
const RING = 'tbRing';
export async function watchRings(onChange: (ring: Ring | null) => void) {
  await bind(RING, payload => onChange(payload ? JSON.parse(payload) : null));
  await run(name => {
    const w = window as any;
    w.__tbRing?.disconnect();
    const read = () => {
      const reject = document.querySelector('[data-tid="reject-call-button"], [data-tid="reject-meeting-button"]') as HTMLElement | null;
      if (!reject?.getClientRects().length) return '';
      // the text of an element outside its buttons, a line per text node (the caller's name, "Video call")
      const texts = (e: Element) => {
        const out: string[] = [];
        const walk = document.createTreeWalker(e, NodeFilter.SHOW_TEXT);
        for (let n = walk.nextNode(); n; n = walk.nextNode()) if (!n.parentElement?.closest('button') && n.textContent!.trim()) out.push(n.textContent!.trim());
        return out;
      };
      let box = reject.parentElement; // the toast: the nearest ancestor with other buttons and some text of its own
      for (let up = 0; box && up < 6 && (box.querySelectorAll('button').length < 2 || !texts(box).length); up++) box = box.parentElement;
      let buttons = [...(box?.querySelectorAll('button') ?? [])].filter(b => b.getClientRects().length) as HTMLElement[];
      if (!box || buttons.length > 8) { box = reject.parentElement!; buttons = [reject]; } // climbed past the toast
      buttons.forEach((b, i) => b.setAttribute('data-tb-ring', String(i)));
      const labels = buttons.map(b => (b.getAttribute('aria-label') || b.innerText || '').trim());
      const lines = texts(box);
      return JSON.stringify({ kind: reject.getAttribute('data-tid') === 'reject-call-button' ? 'call' : 'meeting', title: lines[0] || '',
        message: lines.slice(1, 3).join(' · '), buttons: buttons.map((b, i) => ({ label: labels[i], reject: b === reject })),
        outline: [...box.querySelectorAll('[data-tid]')].map(e => e.getAttribute('data-tid')).join(' ') });
    };
    let last = read(), timer = 0;
    if (last) w[name]?.(last);
    w.__tbRing = new MutationObserver(() => { // a call must not wait: look three times a second at most
      if (timer) return;
      timer = window.setTimeout(() => { timer = 0; const now = read(); if (now !== last) { last = now; w[name]?.(now); } }, 300);
    });
    w.__tbRing.observe(document.body, { subtree: true, childList: true });
  }, RING);
}
// A button of the ringing toast, by its number in Ring.buttons. ponytail: a real click, so a history veil on the page would take it
export const ringClick = (i: number) => click(`[data-tb-ring="${i}"]`);
// "Join" after Teams' own meeting toast is gone: the meeting chat named in the toast, and its Join button.
export async function joinMeeting(text: string) {
  const c = (await chatList()).find(x => x.type === 'meeting' && x.title && text.includes(x.title));
  if (!c) return false;
  await openChat(c);
  if (!(await exists('[data-tid="chat-join-button"]', 3000))) return false;
  await click('[data-tid="chat-join-button"]');
  return true;
}
// A call or meeting in progress: its hang-up button ("Leave" is #hangup-button in a live call; the compact view's one not seen yet)
export const HANGUP = '[data-tid*="hangup" i], [id*="hangup" i]';
// ...or the screen before joining a meeting (camera and microphone choice, the lobby): its parts are "prejoin-…"
export const CALL_UI = `${HANGUP}, [data-tid^="prejoin-"]`;
export const inCall = (): Promise<boolean> => run(s => !!document.querySelector(s), CALL_UI);
// Told at once when a call or a meeting's join screen comes or goes in the page, as the rings are.
const CALL = 'tbCall';
export async function watchCall(onChange: (on: boolean) => void) {
  await bind(CALL, payload => onChange(payload === '1'));
  await run(([name, sel]) => {
    const w = window as any;
    w.__tbCall?.disconnect();
    const read = () => (document.querySelector(sel) ? '1' : '');
    let last = read(), timer = 0;
    if (last) w[name]?.(last);
    w.__tbCall = new MutationObserver(() => {
      if (timer) return;
      timer = window.setTimeout(() => { timer = 0; const now = read(); if (now !== last) { last = now; w[name]?.(now); } }, 300);
    });
    w.__tbCall.observe(document.body, { subtree: true, childList: true });
  }, [CALL, CALL_UI]);
}
// ---------- Jira: a ticket's status and name in the chat's messages ----------
// In a message's text (not in link previews, not in the compose box):
// - a link to <base>/browse/KEY gets data-tb-jira, which the app's style draws after the link;
// - a bare key ("POS20-13443") Jira knows is underlined and opens in Jira on a click, and the text goes after the run of
//   text it stands in, in an empty element of the app's own, also drawn by the style.
// Teams draws its messages with React, which keeps its own text nodes: they are never cut or changed. The underline is a
// CSS highlight (no element), the click is caught on the document, the app's elements are only added next to Teams' own;
// and since a style draws all the texts, the messages' text as read (innerText: the app's own reading) stays as it was.
// Keys with no text yet, or an expired one, go to onKeys, jiraText brings the answer; a click on a key goes to onOpen.
// base '' takes it all away.
const JIRA = 'tbJira', JIRA_OPEN = 'tbJiraOpen';
export async function watchJira(base: string, onKeys: (keys: string[]) => void, onOpen: (key: string) => void) {
  await bind(JIRA, payload => onKeys(JSON.parse(payload)));
  await bind(JIRA_OPEN, key => onOpen(key));
  await run(([name, openName, base]) => {
    const w = window as any;
    w.__tbJira?.disconnect();
    if (w.__tbJiraClick) document.removeEventListener('click', w.__tbJiraClick, true);
    w.__tbJiraInfo = {}; // key → { text, until, found }
    w.__tbJiraAsked = new Set();
    for (const a of document.querySelectorAll('a[data-tb-jira]')) a.removeAttribute('data-tb-jira');
    for (const s of document.querySelectorAll('span.tb-jira')) s.remove();
    (CSS as any).highlights?.delete('tb-jira');
    if (!base) return;
    if (!document.getElementById('tb-jira')) {
      const s = document.createElement('style');
      s.id = 'tb-jira'; // inline-block: the link's underline stays off the text
      s.textContent = 'a[data-tb-jira]::after,span.tb-jira::after{content:"· " attr(data-tb-jira);display:inline-block;margin-left:5px;color:var(--colorNeutralForeground3,#888);text-decoration:none}'
        + '::highlight(tb-jira){text-decoration:underline dotted;text-underline-offset:3px;color:var(--colorBrandForegroundLink,#5b5fc7)}';
      document.head.append(s);
    }
    const prefix = `${base.toLowerCase()}/browse/`;
    const KEY = /\b[A-Z][A-Z0-9]{1,9}-\d{1,6}\b/g;
    let keysAt = new WeakMap<Text, { start: number; end: number; key: string }[]>(); // the underlined keys, for a click
    const ask = (key: string, want: string[]) => {
      const info = w.__tbJiraInfo[key];
      if ((!info || Date.now() > info.until) && !w.__tbJiraAsked.has(key)) { w.__tbJiraAsked.add(key); want.push(key); }
      return info;
    };
    const scan = () => {
      const want: string[] = [], ranges: Range[] = [], kept = new Set<Element>();
      keysAt = new WeakMap();
      for (const content of document.querySelectorAll('[data-tid="chat-pane-message"] [id^="content-"]')) {
        for (const a of content.querySelectorAll<HTMLAnchorElement>('a[href]')) {
          const key = a.href.toLowerCase().startsWith(prefix) && !a.isContentEditable ? a.href.slice(prefix.length).match(/^([a-z][a-z0-9]{1,9}-\d{1,6})(?:[/?#]|$)/i)?.[1].toUpperCase() : undefined;
          const info = key && ask(key, want);
          if (info?.text && a.getAttribute('data-tb-jira') !== info.text) a.setAttribute('data-tb-jira', info.text);
          if (info && !info.text) a.removeAttribute('data-tb-jira');
        }
        const walk = document.createTreeWalker(content, NodeFilter.SHOW_TEXT);
        for (let node = walk.nextNode() as Text | null; node; node = walk.nextNode() as Text | null) {
          if (!node.parentElement || node.parentElement.isContentEditable || node.parentElement.closest('a, code, pre')) continue; // an edited message: its editor's own
          const text = node.nodeValue ?? '', found: { start: number; end: number; key: string }[] = [];
          for (const m of text.matchAll(KEY)) if (ask(m[0], want)?.found) found.push({ start: m.index!, end: m.index! + m[0].length, key: m[0] });
          if (!found.length) continue;
          keysAt.set(node, found);
          for (const f of found) { const r = new Range(); r.setStart(node, f.start); r.setEnd(node, f.end); ranges.push(r); }
          // the key named again only where the text goes on after it, or there are several
          const label = found.map(f => (found.length > 1 || /[\p{L}\p{N}]/u.test(text.slice(f.end)) ? `${f.key}: ` : '') + w.__tbJiraInfo[f.key].text).join('; ');
          let note = node.nextSibling as Element | null;
          if (!(note instanceof HTMLSpanElement && note.classList.contains('tb-jira'))) { note = Object.assign(document.createElement('span'), { className: 'tb-jira' }); node.after(note); }
          if (note.getAttribute('data-tb-jira') !== label) note.setAttribute('data-tb-jira', label);
          kept.add(note);
        }
      }
      for (const s of document.querySelectorAll('span.tb-jira')) if (!kept.has(s)) s.remove(); // its text changed or went
      (CSS as any).highlights?.set('tb-jira', new (window as any).Highlight(...ranges));
      if (want.length) w[name]?.(JSON.stringify(want));
    };
    w.__tbJiraScan = scan;
    w.__tbJiraClick = (e: MouseEvent) => { // a click on an underlined key, not the end of selecting text
      const at = (document as any).caretPositionFromPoint?.(e.clientX, e.clientY);
      const hit = at && keysAt.get(at.offsetNode)?.find(f => at.offset >= f.start && at.offset < f.end);
      if (!hit || String(getSelection() ?? '')) return;
      e.preventDefault(); e.stopPropagation();
      w[openName]?.(hit.key);
    };
    document.addEventListener('click', w.__tbJiraClick, true);
    let timer = 0;
    w.__tbJira = new MutationObserver(() => { if (!timer) timer = window.setTimeout(() => { timer = 0; scan(); }, 500); });
    w.__tbJira.observe(document.body, { subtree: true, childList: true });
    scan();
  }, [JIRA, JIRA_OPEN, base]);
}
// The text after a ticket's links and keys ('' for none), for how long it holds before the page asks again, and whether
// Jira knows the ticket: only then a bare key is marked (UTF-8 or ISO-8601 look like keys too).
export const jiraText = (key: string, text: string, ttl: number, found: boolean) => run(([key, text, ttl, found]) => {
  const w = window as any;
  if (!w.__tbJiraInfo) return;
  w.__tbJiraInfo[key] = { text, until: Date.now() + ttl, found };
  w.__tbJiraAsked.delete(key);
  w.__tbJiraScan?.();
}, [key, text, ttl, found]);
// Chromium's "memory is short" pass and a full garbage collection: the page drops what it can rebuild.
export const trimMemory = () => cdp('Memory.simulatePressureNotification', { level: 'critical' }).then(() => cdp('HeapProfiler.collectGarbage'));

export async function findChat(part: string) {
  const hits = (await chatList()).filter(c => c.title.includes(part));
  if (!hits.length) throw fail(404, `no chat matches "${part}"`);
  if (hits.length > 1) throw fail(409, `ambiguous "${part}": ${hits.map(c => c.title).join(' | ')}`);
  return hits[0];
}
export async function selfChat() {
  const c = (await chatList()).find(x => x.type === 'self');
  if (!c) throw fail(500, 'self-chat "(вы)" not in the chat list');
  return c;
}
// The user's display name: the self-chat title without its "(you)" suffix in whatever language.
export const myName = async () => (await selfChat()).title.replace(/\s*\([^)]*\)\s*$/, '');

// The user's own status as their avatar button names it ("Ваш профиль, состояние Не беспокоить"), in Teams' language.
export const myStatus = (): Promise<string> => run(() => document.getElementById('idna-me-control-avatar-trigger')?.getAttribute('aria-label') ?? '');
// ---------- theme ----------
// Teams draws itself in the scheme the app set: its page background is dark exactly when prefers-color-scheme says dark.
export const themeFollowed = (): Promise<boolean> => run(() => {
  const [r, g, b, a = 1] = (getComputedStyle(document.body).backgroundColor.match(/[\d.]+/g) ?? []).map(Number);
  return !a || (r + g + b < 384) === matchMedia('(prefers-color-scheme: dark)').matches; // a transparent body tells nothing
});
// Teams' own theme set to "as the operating system" (… → Settings → Appearance), so it follows the app's theme from then on.
export async function followSystemTheme() {
  await click('[data-tid="more-options-header"]');
  try {
    const step = async (sel: string, what: string, timeout = 8000) => {
      if (!(await exists(sel, timeout))) throw fail(502, `${what} not found in Teams`);
      await click(sel);
    };
    await step('[data-tid="settings-button-menu"]', 'Settings', 3000);
    await step('[data-tid="appearance"]', 'Appearance');
    await step('[data-tid="appearance-settings-theme-selector-dropdown"]', 'the theme choice');
    await step('[data-tid="appearance-settings-theme-selector-dropdown-followOsTheme-dropdown-option"]', 'the "as the system" theme', 3000);
  } finally { await showChats(); }
}

// ---------- Activity: mentions in channels ----------
// kind: the icon Teams gives the item — iconsMention (you), everyoneMentionIcon, tagIcon, followChannelIcon (a new post)…
// title: who did what, in Teams' words; where: the channel or chat.
export type Activity = { id: string; kind: string; title: string; text: string; where: string };
// A number on the Activity button: something new there. ponytail: its badge was never seen unread; the caller also reads on a timer
export const activityBadge = (): Promise<boolean> => run(s => {
  const b = document.querySelector(s) as HTMLElement | null;
  return !!b && (/\d/.test(b.innerText) || !!b.querySelector('[class*="adge"]'));
}, ACTIVITY);
// The latest items of the Activity feed; Teams is back on the chat list afterwards.
export async function activity(): Promise<Activity[]> {
  await click(ACTIVITY);
  try {
    if (!(await exists('[data-tid="activity-feed-list-item"]', 10_000))) return [];
    return await run(() => [...document.querySelectorAll('[data-tid="activity-feed-list-item"]')].slice(0, 20).map(e => {
      const id = e.querySelector('[data-tid^="activity-feed-item-media-"]')?.getAttribute('data-tid')?.slice('activity-feed-item-media-'.length) ?? '';
      const part = (name: string) => (document.getElementById(`activity-feed-item-${name}-${id}`) as HTMLElement | null)?.innerText.replace(/\s+/g, ' ').trim() ?? '';
      return { id, kind: e.querySelector('[data-tid^="activity-feed-item-activity-icon-"] svg[data-testid]')?.getAttribute('data-testid') ?? '',
        title: part('title'), text: part('message-preview'), where: part('location') };
    }).filter(a => a.id));
  } finally { await showChats(); }
}
// Opens an Activity item where Teams shows it (the channel post), for the user: the page stays there.
export async function openActivity(id: string) {
  await click(ACTIVITY);
  const item = `[data-tid="activity-feed-item-media-${id}"]`;
  if (!(await exists(item, 10_000))) throw fail(404, 'the item is no longer in Activity');
  await click(item);
}
export const currentChatId =(): Promise<string | undefined> => run(() => document.querySelector('[id^="chat-header-"]')?.id.replace('chat-header-', ''));
const isOpen = (c: ChatRef): Promise<boolean> => run(id => !!document.getElementById(`chat-header-${id}`), c.id);
export async function openChat<T extends ChatRef>(c: T) {
  if (await isOpen(c)) return c;
  await showChats();
  // right after load the list still shifts and a click can miss, so retry
  for (let attempt = 1; ; attempt++) {
    await click(`[role="treeitem"][data-fui-tree-item-value$="|${c.id}"]`);
    if (await waitFor(id => !!document.getElementById(`chat-header-${id}`), c.id, 5000)) break;
    if (attempt === 3) throw fail(502, `chat "${c.title}" did not open`);
  }
  await exists('[id^="message-body-"]');
  return c;
}

// Own messages carry Fluent's ChatMyMessage class; consecutive messages may show the author only once.
export const lastMessages = (n: number): Promise<Message[]> => run(n => {
  let author = '';
  return [...document.querySelectorAll('[data-tid="chat-pane-message"]')].map(m => {
    author = (m.closest('[data-tid="chat-pane-item"]')?.querySelector('[data-tid="message-author-name"]') as HTMLElement | null)?.innerText.trim() || author;
    const mid = m.getAttribute('data-mid') || '';
    // the message element itself also holds reactions ("1 реакция"); content-<id> is the text alone
    const content = (document.getElementById(`content-${mid}`) ?? m) as HTMLElement;
    // "Reply with quote" draws the quoted message as a card inside the content: its author and text are not this message's
    const card = content.querySelector('[data-tid="quoted-reply-card"]') as HTMLElement | null;
    const box = (card?.closest('[data-track-module-name="messageQuotedReply"]') ?? card) as HTMLElement | null;
    const flat = (s: string) => s.replace(/\s+/g, ' ').trim();
    const quote = card && { replyTo: flat((card.querySelector('[data-tid="quoted-reply-timestamp"]')?.previousElementSibling as HTMLElement | null)?.innerText ?? ''),
      quote: flat((card.querySelector('[data-tid="quoted-reply-preview-content"]') as HTMLElement | null)?.innerText ?? '') };
    return { id: mid, author, mine: m.className.includes('ChatMyMessage'), text: flat(box ? content.innerText.replace(box.innerText, '') : content.innerText), ...quote };
  }).filter(m => m.id).slice(-n);
}, n);

// Up to n latest messages, newer than `since` (epoch ms) or, with afterMine, after the user's own last message:
// Teams keeps only a window of them in the page, so scroll up and collect until enough, no more, or stopped.
// veil: a text shown over the page while it reads; the veil also takes the clicks, since a click on another chat
// would mix that chat into the history.
type HistoryOptions = { since?: number; afterMine?: boolean; signal?: AbortSignal; veil?: string };
const VEIL = (text: string) => `body::before { content: ''; position: fixed; inset: 0; z-index: 2147483646; background: rgba(0, 0, 0, .35); cursor: progress; }
  body::after { content: ${JSON.stringify(text)}; position: fixed; left: 50%; top: 40%; transform: translateX(-50%); z-index: 2147483647; max-width: 60vw;
    padding: 12px 18px; border-radius: 8px; background: #292929; color: #fff; font: 14px "Segoe UI", system-ui, sans-serif; box-shadow: 0 4px 16px rgba(0, 0, 0, .4); }`;
export async function history(c: ChatRef, n: number, { since = 0, afterMine, signal, veil }: HistoryOptions = {}): Promise<Message[]> {
  await openChat(c);
  const veiled = veil ? await wc.insertCSS(VEIL(veil)) : undefined; // a JSON string is a valid CSS string: the title cannot break out of it
  const seen = new Map<string, Message>();
  const collect = async () => { for (const m of await lastMessages(1000)) seen.set(m.id, m); };
  const scroll = (top: boolean) => run(top => { const v = document.querySelector('[data-tid="message-pane-list-viewport"]'); if (v) v.scrollTop = top ? 0 : v.scrollHeight; }, top);
  const lastMine = () => Math.max(0, ...[...seen.values()].filter(m => m.mine).map(m => Number(m.id))); // ids are epoch ms
  const enough = () => seen.size >= n || Math.min(...[...seen.keys()].map(Number)) < since || (!!afterMine && lastMine() > 0);
  try {
    await collect();
    for (let idle = 0, rounds = 0; !enough() && idle < 3 && rounds < 60 && !signal?.aborted; rounds++) { // ponytail: ~60 screens max, enough for a working day
      const before = seen.size;
      await scroll(true);
      await sleep(1200); // older messages load lazily
      await collect();
      idle = seen.size === before ? idle + 1 : 0;
    }
    await scroll(false);
    // another chat opened meanwhile (from the keyboard, say): what was collected may belong to it
    if (!(await isOpen(c))) throw Object.assign(new Error('another chat was opened while reading'), { switched: true });
  } finally { if (veiled) await wc.removeInsertedCSS(veiled).catch(() => {}); }
  const from = Math.max(since, afterMine ? lastMine() + 1 : 0);
  return [...seen.values()].filter(m => Number(m.id) >= from).sort((a, b) => Number(a.id) - Number(b.id)).slice(-n);
}

const BOX = '[id^="new-message-"]';
const focusBox = () => run(b => (document.querySelector(b) as HTMLElement | null)?.focus(), BOX);
const boxFocused = (): Promise<boolean> => run(b => !!document.activeElement?.closest(b), BOX);
// Keys go out only while the compose box has the focus: Delete or Enter on a focused message could act on that message.
async function ensureBoxFocus() {
  if (await boxFocused()) return true;
  await focusBox();
  return boxFocused();
}
// Leaves nothing half-typed in the user's chat when sending is aborted.
const clearBox = async () => { if (await ensureBoxFocus()) { await press('a', CTRL); await press('Delete'); } };
// What the user has typed but not sent in the open chat ("Improve my text" reads it), without a quoted message's card.
export const boxText = (): Promise<string> => run(b => {
  const box = document.querySelector(b) as HTMLElement | null;
  return (box?.innerText ?? '').replace((box?.querySelector('card') as HTMLElement | null)?.innerText || '\0', '');
}, BOX);
const boxQuote = () => exists(`${BOX} card`, 0); // "Reply with quote" puts the quoted message in the box as a card

// ---------- a message's hover bar: reactions and quoted replies ----------
// The other side's latest message in the open chat: the one a toast announced, or a newer one.
const theirLast = async () => (await lastMessages(30)).reverse().find(m => !m.mine);
// Rests the mouse on a message, so Teams shows its hover bar, and clicks a button of that bar.
async function hoverBar(mid: string, sel: string) {
  for (let attempt = 1; ; attempt++) {
    const at = await run(id => {
      const e = document.getElementById(`message-body-${id}`);
      if (!e) return null;
      e.scrollIntoView({ block: 'center' });
      const r = e.getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + Math.min(r.height / 2, 16) };
    }, mid);
    if (!at) throw fail(404, 'message not found in the chat');
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: 1, y: 1 }); // the bar shows on entering the message
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: at.x, y: at.y });
    if (await exists(sel, 3000)) return click(sel);
    if (attempt === 3) throw fail(502, 'Teams did not show the message actions');
  }
}
// A reaction on the other side's latest message: Teams' own quick one from the hover bar when it has it, else from the full
// picker (tab: the picker's category with that reaction).
// emoticon id → the hover bar's message-actions-<name>. Teams' quick 👍 "Like" is the emoticon "yes"; the picker's "like"
// ("Мне нравится") is another thumb, the one that went out instead.
const QUICK: Record<string, string> = { yes: 'like', heart: 'heart', laugh: 'laugh', surprised: 'surprised' };
export async function react(c: ChatRef, emoticon: string, tab: number) {
  await openChat(c);
  const m = await theirLast();
  if (!m) throw fail(404, 'no message from the other side in this chat');
  if (QUICK[emoticon]) { await hoverBar(m.id, `[data-tid="message-actions-${QUICK[emoticon]}"]`); return { chat: c.title, reaction: emoticon }; }
  await hoverBar(m.id, '[data-tid="expanded-reactions-picker-entry"]');
  const button = `[data-tid="emoticon-button-${emoticon}"]`;
  if (!(await exists('[data-tid^="emoji-categories-tab-"]', 5000))) throw fail(502, 'the reactions picker did not open');
  if (!(await exists(button, 300))) await click(`[data-tid="emoji-categories-tab-${tab}"]`); // a category shows its reactions once opened
  if (!(await exists(button, 5000))) { await press('Escape'); throw fail(404, `Teams has no reaction "${emoticon}" now`); }
  await click(button);
  return { chat: c.title, reaction: emoticon };
}

// The user's unsent text in the compose box is never sent along or wiped: with `replace` (that text, as "Improve my text"
// read it) it gives way to the new one, otherwise sending stops. quote: as "Reply with quote" to the other side's latest message.
export async function send(c: ChatRef, text: string, { replace, quote }: { replace?: string; quote?: boolean } = {}) {
  await openChat(c);
  await click(BOX);
  const had = plain(await boxText());
  if (had !== plain(replace ?? '') || (!replace && await boxQuote()))
    throw Object.assign(fail(409, replace ? 'the text in the compose box has changed, not sent' : 'the compose box has your unsent text, not sent'), { boxTaken: !replace });
  if (replace) await clearBox();
  if (quote) {
    const m = await theirLast();
    if (!m) throw fail(404, 'no message from the other side to quote, not sent');
    await hoverBar(m.id, '[data-tid="message-actions-more"]');
    if (!(await exists('[data-tid="message-actions-quoted-reply"]', 3000))) { await press('Escape'); throw fail(502, 'no "Reply with quote" in Teams, not sent'); }
    await click('[data-tid="message-actions-quoted-reply"]');
    if (!(await waitFor(s => !!document.querySelector(s), `${BOX} card`, 5000))) throw fail(502, 'Teams did not add the quote, not sent');
    if (!(await ensureBoxFocus())) { await clearBox(); throw fail(500, 'compose box lost the focus, not sent'); }
  }
  const lines = text.split('\n');
  try {
    for (let i = 0; i < lines.length; i++) {
      if (i) await press('Enter', SHIFT);
      if (lines[i]) await typeLine(lines[i]);
    }
  } catch (e) { await press('Escape'); await clearBox(); throw e; }
  const typed = plain(await boxText());
  if (typed !== plain(text)) {
    await clearBox();
    throw fail(500, `compose box mismatch, not sent: ${typed}`);
  }
  if (!(await isOpen(c))) throw fail(409, 'chat switched before send, not sent');
  if (!(await ensureBoxFocus())) throw fail(500, 'compose box lost the focus, not sent');
  await press('Enter');
  const probe = plain(lines.find(l => plain(l)) ?? '').slice(0, 40);
  const seen = await waitFor(t => [...document.querySelectorAll('[id^="message-body-"]')].slice(-3).some(e =>
    (e.getAttribute('aria-label') || (e as HTMLElement).innerText).replace(/[\p{Extended_Pictographic}️]/gu, '').replace(/\s+/g, ' ').includes(t)), probe);
  if (!seen) throw fail(504, 'sent, but not seen in chat yet: check Teams before retrying');
  return { chat: c.title, sent: text };
}

// Opens the GIF tab of the compose-area picker and searches; leaves the picker open.
const RESULTS = '[role="menuitem"]:has(img)';
async function searchGifs(query: string): Promise<{ label: string; src: string }[]> {
  if (!(await exists(BOX, 0))) await openChat(await selfChat()); // the picker lives in a compose area
  const gifTab = () => !!document.querySelector('[data-tid="unified-picker-tab-gifs"]');
  // the button toggles, so click only while the picker is closed; right after load a click can be ignored, so retry.
  // A multi-line draft moves it from "sendMessageCommands-…" to "newMessageCommands-…".
  const FUN = '[data-tid$="MessageCommands-UnifiedFunPicker"]';
  for (let attempt = 1; !(await run(gifTab)); attempt++) {
    if (attempt > 3) throw fail(502, 'GIF picker did not open');
    await exists(FUN, 5000); // the toolbar re-renders while typing
    await click(FUN);
    await waitFor(gifTab, null, 3000);
  }
  await click('[data-tid="unified-picker-tab-gifs"]');
  const input = '[data-tid="unified-picker-content"] input[type="text"]';
  await exists(input, 5000);
  await click(input, -1);
  await press('a', CTRL);
  await insertText(query);
  if (!(await exists(RESULTS, 10_000))) { await press('Escape'); throw fail(404, `no GIF for "${query}"`); }
  await sleep(1500); // results re-render while the search settles
  return run(s => [...document.querySelectorAll(s)].map(x => ({ label: x.getAttribute('aria-label') || '', src: x.querySelector('img')?.src || '' })), RESULTS);
}

export async function gifSearch(query: string, n: number) {
  const list = await searchGifs(query);
  await press('Escape');
  return list.slice(0, n);
}

const boxHasImage = (): Promise<boolean> => run(b => !!document.querySelector(b)?.querySelector('img'), BOX);
// With `text` the GIF goes in after it, as one message.
export async function sendGif(c: ChatRef, query: string, src?: string, text?: string) {
  await openChat(c);
  if (await boxHasImage()) throw fail(409, 'compose box already has an image, not sent');
  let label: string;
  try {
    if (text) { await click(BOX); await typeLine(text); }
    const list = await searchGifs(query);
    const i = src ? list.findIndex(g => g.src === src) : Math.max(0, list.findIndex(g => g.label.toLowerCase().includes(query.toLowerCase())));
    if (i < 0) throw fail(404, 'chosen GIF is no longer in the results, not sent');
    label = list[i].label;
    await click(RESULTS, i);
    if (!(await waitFor(b => !!document.querySelector(b)?.querySelector('img[src*="giphy"]'), BOX, 10_000))) throw fail(502, 'GIF did not reach the compose box, not sent');
    const typed = plain(await run(b => (document.querySelector(b) as HTMLElement | null)?.innerText ?? '', BOX));
    if (typed !== plain(text ?? '')) throw fail(500, `compose box mismatch, not sent: ${typed}`);
    if (!(await isOpen(c))) throw fail(409, 'chat switched before send, not sent');
  } catch (e) { await press('Escape'); await clearBox(); throw e; }
  if (!(await ensureBoxFocus())) throw fail(500, 'compose box lost the focus, not sent');
  await press('End', CTRL);
  await press('Enter');
  if (!(await waitFor(b => !document.querySelector(b)?.querySelector('img'), BOX))) throw fail(504, 'GIF still in compose box: check Teams before retrying');
  return { chat: c.title, gif: label };
}


