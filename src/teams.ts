// Drives the Teams web page in the hidden window: DOM reads via executeJavaScript, real input via the DevTools protocol.
// Selectors use Teams' own data-tid / ids and image names, not UI texts, so any Teams language works.
import type { WebContents } from 'electron';

export const TEAMS_URL = 'https://teams.cloud.microsoft/';
// preview: the last message as the chat list shows it ("Author: text" in group chats, cut by Teams); time: "16:33" today, a date before
export type Chat = { id: string; title: string; preview: string; time: string; type: string; unread: boolean; mention: boolean };
// replyTo, quote: a quoted reply's author and text, kept out of the message's own text
// images: the srcs of the message's pictures (not its emoji), for imageData
export type Message = { id: string; author: string; mine: boolean; text: string; replyTo?: string; quote?: string; images?: string[] };
type ChatRef = Pick<Chat, 'id' | 'title'>;

export const fail = (status: number, message: string) => Object.assign(new Error(message), { status });
const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));
// Teams may render emoji as images, so text checks compare without them.
export const plain = (t: string) => t.replace(/[\p{Extended_Pictographic}️​-‍⁠﻿]/gu, '').replace(/\s+/g, ' ').trim();

// Teams' own pages: the app talks to these only.
export const TEAMS_PAGE = /^https:\/\/([\w-]+\.)*(teams\.cloud\.microsoft|teams\.microsoft\.com|teams\.live\.com)\//;
// A page that stopped answering (a hung renderer) fails the step instead of holding the app's queue for ever.
const within = <T>(p: Promise<T>, ms: number, what: string) => new Promise<T>((ok, no) => {
  const timer = setTimeout(() => no(fail(504, `the Teams page did not answer in ${ms / 1000} s (${what})`)), ms);
  p.then(v => { clearTimeout(timer); ok(v); }, e => { clearTimeout(timer); no(e); });
});

let wc: WebContents;
const cdp = (method: string, params: object = {}) => within(wc.debugger.sendCommand(method, params), 20_000, method);

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
  : within(wc.executeJavaScript(`(${fn})(${JSON.stringify(arg)})`, true), 30_000, 'a script');

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
  const items = [...document.querySelectorAll('[role="treeitem"][data-item-type="chat"]')];
  return {
    rail: has('[data-tid="app-bar-wrapper"]'),
    chat: has('[data-tid="86fcd49b-61a2-4701-b771-54728cd291fb"]'),
    activity: has('[data-tid="14d6962d-6eeb-4f48-8890-de55454bb136"]'),
    list: listShown ? has('[role="treeitem"][data-item-type="chat"] [id^="title-chat-list-item_"]') : null,
    // a chat's id ("<kind>|<id>") and the ids its unread and @mention marks are named by
    ids: listShown && items.length ? items.every(e => (e.getAttribute('data-fui-tree-item-value') || '').includes('|')) : null,
    marks: listShown && items.length ? items.some(e => (e.getAttribute('aria-labelledby') || '').includes('chat-list-item_')) : null,
    header: chatOpen ? has('[id^="chat-header-"]') : null,
    messages: chatOpen ? has('[data-tid="chat-pane-message"][data-mid]') && has('[id^="content-"]') : null,
    compose: chatOpen ? has('[id^="new-message-"]') : null,
    chips: chatOpen && has('[id^="chat-header-"]') ? has('[data-tb-chips]') : null, // Soroka's buttons above the box (watchUi)
    menu: has('[data-tid="more-options-header"]'),
    bar: has('[data-tb-group]'), // Soroka's ▤ ⚙ in Teams' header (watchUi)
  };
});

// Back to the chat list from Activity or Teams (the user's, or the app's own visit to Activity).
export async function showChats() {
  if (await exists('[role="tree"]', 0)) return;
  await click(CHAT_APP);
  if (!(await exists('[role="tree"]', 10_000))) throw fail(502, 'the chat list did not open');
}

// Real mouse click in the middle of the index-th match (negative index counts from the end).
async function click(sel: string, index = 0) {
  const at = await run(({ sel, index }) => {
    const e = [...document.querySelectorAll(sel)].at(index);
    if (!e) return null;
    e.scrollIntoView({ block: 'center' });
    const r = e.getBoundingClientRect();
    return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
  }, { sel, index });
  if (!at) throw fail(404, `not found: ${sel}`);
  for (const type of ['mouseMoved', 'mousePressed', 'mouseReleased'])
    await cdp('Input.dispatchMouseEvent', { type, x: at.x, y: at.y, button: 'left', clickCount: 1 });
}

const KEYS: Record<string, { key: string; code: string; windowsVirtualKeyCode: number; text?: string }> = {
  Enter: { key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r' },
  Escape: { key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 },
  Tab: { key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9 },
  End: { key: 'End', code: 'End', windowsVirtualKeyCode: 35 },
  Backspace: { key: 'Backspace', code: 'Backspace', windowsVirtualKeyCode: 8 },
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
    if (!(await waitFor(([s, id]) => [...document.querySelectorAll(s)].some(o => (o.querySelector('img')?.getAttribute('src') || '').includes(`/emoticons/${id}/`)), [SUGGESTION, id], 6000))) {
      for (let i = 0; i <= code.length; i++) await press('Backspace'); // Teams no longer offers its picture: the plain emoji instead
      await insertText(part);
      continue;
    }
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
// Only Teams' own page may call them: a binding stays in the page across navigations, to a sign-in page or wherever a link leads.
const bindings = new Map<string, (payload: string) => void>();
async function bind(name: string, fn: (payload: string) => void) {
  const page = wc;
  if (!page.debugger.listenerCount('message'))
    page.debugger.on('message', (_e, method, params) => {
      if (method !== 'Runtime.bindingCalled' || !TEAMS_PAGE.test(page.getURL())) return;
      try { bindings.get(params.name)?.(String(params.payload)); } catch { /* a malformed call: nothing to do */ }
    });
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

// Teams' own parts, never the inside of a message: a card or formatted text in a message might carry an id or a data-tid
// that looks like one of them (not seen; nothing known keeps a message's markup from bringing one).
const MSG = '[data-tid="chat-pane-message"]';
const OUT = `:not(${MSG} *)`;

// ---------- Soroka's own buttons in the Teams page ----------
// Above the compose box of an open chat: "Suggest a reply", "Improve my text", "Chat summary" — a click runs the same as the
// hotkey, nothing is sent. Teams re-renders its page at will: an observer puts the row back, always one.
// The row goes in Teams' compose footer, right before the card that holds the box (seen live 2026-11-23: the footer's
// children are a gutter for "typing…", the box's card, a 16 px line under it). A chat, not a channel: its header
// [id^="chat-header-"] is there. The colours are Teams' own Fluent tokens, so they follow Teams' theme, not Windows'.
// And one title bar: Teams' own header ([data-tid="title-bar"], 48 px, seen live 2026-11-23: a flex row; "···" and the avatar
// in its titlebar-end-slot) moves the window, Soroka's ▤ and ⚙ stand before its "···", Soroka's icon at its left end,
// and Windows' — ☐ ✕ are drawn over its right end (a padding keeps Teams' own buttons clear of them).
export type UiAct = 'suggest' | 'improve' | 'summary' | 'drafts' | 'settings';
const UI_ACTS: string[] = ['suggest', 'improve', 'summary', 'drafts', 'settings'] satisfies UiAct[];
// labels: "✨ Suggest reply" — the first word is the icon; ai: the AI ones are shown; running: a spinner, clicks wait
export type Chips = { suggest: string; improve: string; summary: string; tips: Record<'suggest' | 'improve' | 'summary', string>; empty: string; ai: boolean; running: UiAct[] };
// badge: the mark on ▤ (src/badge.ts); busy: what Soroka does now (a spinner, the text in ▤'s tooltip); panel, settings: open;
// controls: the width Windows' buttons take over the header's right end, in px
export type BarUi = { badge: { red: number; dot: boolean }; busy: string; panel: boolean; settings: boolean; controls: number; tips: { drafts: string; settings: string } };
export type PageUi = { chips: Chips; bar: BarUi };
// Teams' header as found: its height and colours (#rrggbb) for the window's caption buttons
export type Bar = { height: number; bg: string; fg: string };
// A user style sheet (insertCSS) loses to Teams' own rules unless !important: every rule here is. Its colours do not turn
// a button's native look off as an author's would (seen live): appearance: none does.
export const UI_CSS = `
[data-tb-chips] { display: flex !important; align-items: center !important; gap: 6px !important; padding: 0 0 6px 2px !important; }
[data-tb-chips] > img { width: 16px !important; height: 16px !important; margin-right: 2px !important; }
.tb-chip { appearance: none !important; display: inline-flex !important; align-items: center !important; gap: 6px !important; height: 28px !important; padding: 0 12px !important; margin: 0 !important;
  border-radius: 14px !important; box-sizing: border-box !important; cursor: pointer !important; border: 1px solid var(--colorBrandStroke2, #c5cbfa) !important;
  background: transparent !important; color: var(--colorBrandForeground1, #5b5fc7) !important; font: inherit !important; font-size: 12px !important; white-space: nowrap !important; }
.tb-chip:hover { background: var(--colorBrandBackground2, #ebf3fc) !important; }
.tb-chip[aria-disabled="true"] { opacity: .45 !important; cursor: default !important; background: transparent !important; }
.tb-spin { width: 10px !important; height: 10px !important; border-radius: 50% !important; box-sizing: border-box !important;
  border: 2px solid color-mix(in srgb, currentColor 25%, transparent) !important; border-top-color: currentColor !important; animation: tb-spin .8s linear infinite !important; }
@keyframes tb-spin { to { transform: rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { .tb-spin { animation-duration: 3s !important; } }
[data-tb-bar], [data-tb-bar] * { app-region: drag !important; }
[data-tb-bar] :is(button, a, input, textarea, select, [role="button"], [role="search"], [role="combobox"], [role="menuitem"], [role="tab"], [contenteditable="true"]),
[data-tb-bar] :is(button, a, input, textarea, select, [role="button"], [role="search"], [role="combobox"], [role="menuitem"], [role="tab"], [contenteditable="true"]) * { app-region: no-drag !important; }
[data-tb-bar] { padding-right: calc(var(--tb-controls, 0px) + 8px) !important; box-sizing: border-box !important; }
[data-tb-logo] { width: 18px !important; height: 18px !important; margin: 0 4px 0 14px !important; align-self: center !important; flex: none !important; }
[data-tb-group] { display: flex !important; align-items: center !important; gap: 2px !important; padding: 0 4px !important; height: 32px !important; margin-right: 4px !important;
  border-radius: 6px !important; background: var(--colorBrandBackground2, #e2e3f7) !important; }
.tb-btn { appearance: none !important; position: relative !important; width: 32px !important; height: 28px !important; padding: 0 !important; margin: 0 !important; border: 0 !important;
  border-radius: 4px !important; display: flex !important; align-items: center !important; justify-content: center !important; cursor: pointer !important; font: inherit !important;
  font-size: 15px !important; background: transparent !important; color: var(--colorBrandForeground1, #4f52b2) !important; }
.tb-btn:hover, .tb-btn[aria-pressed="true"] { background: var(--colorBrandBackground2Hover, rgba(0,0,0,.07)) !important; }
.tb-btn .tb-badge { position: absolute !important; top: 0 !important; right: -3px !important; background: #c4314b !important; color: #fff !important; border-radius: 9px !important;
  font-size: 10px !important; font-weight: 600 !important; line-height: 14px !important; padding: 0 5px !important; }
.tb-btn .tb-dot { position: absolute !important; top: 3px !important; right: 3px !important; width: 8px !important; height: 8px !important; border-radius: 50% !important;
  background: var(--colorBrandForeground1, #5b5fc7) !important; }
.tb-btn .tb-spin { position: absolute !important; top: 2px !important; right: 2px !important; width: 9px !important; height: 9px !important; }`;
const ACT = 'tbAct', BAR = 'tbBar';
export async function watchUi(icon: string, onAct: (act: UiAct) => void, onBar: (bar: Bar | null) => void) {
  await bind(ACT, payload => { if (UI_ACTS.includes(payload)) onAct(payload as UiAct); });
  await bind(BAR, payload => {
    if (!payload) return onBar(null);
    const b = JSON.parse(payload), color = (c: unknown) => (/^#[0-9a-f]{6}$/.test(String(c)) ? String(c) : '#ebebeb');
    onBar({ height: Math.min(64, Math.max(24, Math.round(Number(b.height)) || 48)), bg: color(b.bg), fg: color(b.fg) });
  });
  await run(([name, barName, box, icon]) => {
    const w = window as any;
    w.__tbUiWatch?.disconnect();
    clearTimeout(w.__tbBarLost);
    w.__tbBarSent = undefined; // a new watcher: the app hears where things are again
    // a header back within its second cancels the loss: the next one counts its own second
    const tell = (geo: string) => { clearTimeout(w.__tbBarLost); w.__tbBarLost = 0; if (geo !== w.__tbBarSent) { w.__tbBarSent = geo; w[barName]?.(geo); } };
    const hex = (c: string) => '#' + (c.match(/[\d.]+/g) ?? []).slice(0, 3).map(n => Math.round(Number(n)).toString(16).padStart(2, '0')).join('');
    // the header's own background, or the first one behind it that is not transparent
    const bgOf = (e: Element | null) => {
      for (; e; e = e.parentElement) { const c = getComputedStyle(e).backgroundColor, a = c.match(/[\d.]+/g) ?? []; if (a.length && (a.length < 4 || Number(a[3]) > 0)) return c; }
      return 'rgb(235, 235, 235)';
    };
    const bar = () => {
      const ui = w.__tbUi, head = document.querySelector('[data-tid="title-bar"]'), more = head?.querySelector('#more-options-header')?.parentElement;
      // the icon in Teams' start slot (its menu button): one more item in the header itself would shift its space-between layout
      const start = head?.querySelector('[data-tid="titlebar-start-slot"]') ?? head;
      const ok = !!ui && !!head && !!more;
      for (const e of document.querySelectorAll('[data-tb-bar]')) if (!ok || e !== head) e.removeAttribute('data-tb-bar');
      for (const e of document.querySelectorAll('[data-tb-logo]')) if (!ok || e.parentElement !== start) e.remove();
      for (const e of document.querySelectorAll('[data-tb-group]')) if (!ok || e.nextElementSibling !== more) e.remove();
      if (!ok) { // gone for a second, not for a frame of Teams' re-render: then Soroka's own strip comes back
        if (w.__tbBarSent && !w.__tbBarLost) w.__tbBarLost = window.setTimeout(() => { w.__tbBarLost = 0; if (!document.querySelector('[data-tb-bar]')) tell(''); }, 1000);
        return;
      }
      head!.setAttribute('data-tb-bar', '');
      document.documentElement.style.setProperty('--tb-controls', `${ui.bar.controls}px`);
      if (!start!.querySelector(':scope > [data-tb-logo]')) {
        const logo = document.createElement('img');
        logo.src = icon; logo.alt = ''; logo.setAttribute('data-tb-logo', '');
        start!.prepend(logo);
      }
      let group = more!.previousElementSibling as HTMLElement | null;
      if (!group?.hasAttribute('data-tb-group')) {
        group = document.createElement('span');
        group.setAttribute('data-tb-group', '');
        group.addEventListener('mousedown', e => e.preventDefault()); // the caret stays where the user types, as with the strip's ▤
        group.addEventListener('click', e => { const b = (e.target as Element).closest('button'); if (b) w[name]?.(b.dataset.tbAct); });
        more!.before(group);
      }
      const b = ui.bar, key = JSON.stringify(b);
      if ((group as any).__tbKey !== key) {
        (group as any).__tbKey = key;
        const button = (act: string, glyph: string, tip: string, pressed: boolean, ...marks: Element[]) => {
          const e = document.createElement('button');
          e.type = 'button'; e.className = 'tb-btn'; e.dataset.tbAct = act; e.title = tip;
          e.setAttribute('aria-label', tip); e.setAttribute('aria-pressed', String(pressed));
          e.append(glyph, ...marks);
          return e;
        };
        // red wins; then the spinner while Soroka works; then the dot of a result asked for
        const mark = document.createElement('span');
        if (b.badge.red) { mark.className = 'tb-badge'; mark.textContent = b.badge.red > 9 ? '9+' : String(b.badge.red); }
        else mark.className = b.busy ? 'tb-spin' : 'tb-dot';
        group.replaceChildren(button('drafts', '▤', b.busy || b.tips.drafts, b.panel, ...(b.badge.red || b.busy || b.badge.dot ? [mark] : [])),
          button('settings', '⚙', b.tips.settings, b.settings));
      }
      const bg = bgOf(head!), [r, g, bl] = (bg.match(/[\d.]+/g) ?? []).map(Number);
      tell(JSON.stringify({ height: Math.round(head!.getBoundingClientRect().height), bg: hex(bg), fg: r + g + bl < 384 ? '#ffffff' : '#242424' }));
    };
    const render = () => {
      bar();
      const ui = w.__tbUi, input = document.querySelector(box) as HTMLElement | null;
      const card = input?.closest('[data-tid="chat-pane-compose-message-footer"] > *');
      const here = !!ui && !!card && !!document.querySelector('[id^="chat-header-"]');
      for (const r of document.querySelectorAll('[data-tb-chips]')) if (!here || r.nextElementSibling !== card) r.remove();
      if (!here) return;
      let row = card!.previousElementSibling as HTMLElement | null;
      if (!row?.hasAttribute('data-tb-chips')) {
        row = document.createElement('div');
        row.setAttribute('data-tb-chips', '');
        row.addEventListener('mousedown', e => e.preventDefault()); // the caret stays in the box
        row.addEventListener('click', e => {
          const b = (e.target as Element).closest('button');
          if (b && !b.getAttribute('aria-disabled') && !b.querySelector('.tb-spin')) w[name]?.(b.dataset.tbAct);
        });
        card!.before(row);
      }
      // the box's text without a quoted message's card, as boxText reads it
      const empty = !(input!.innerText ?? '').replace((input!.querySelector('card') as HTMLElement | null)?.innerText || '\0', '').trim();
      const c = ui.chips, key = JSON.stringify([c, empty]);
      if ((row as any).__tbKey === key) return; // the observer sees the row's own changes too
      (row as any).__tbKey = key;
      const chip = (act: string, label: string, tip: string, off = false) => {
        const b = document.createElement('button');
        b.type = 'button'; // never a submit of Teams' form
        b.className = 'tb-chip'; b.dataset.tbAct = act; b.title = tip;
        if (off) b.setAttribute('aria-disabled', 'true');
        const [mark, ...words] = label.split(' '), ic = document.createElement('span');
        if (c.running.includes(act)) ic.className = 'tb-spin'; else ic.textContent = mark;
        b.append(ic, words.join(' '));
        return b;
      };
      const logo = document.createElement('img');
      logo.src = icon; logo.alt = '';
      row.replaceChildren(logo, ...(c.ai ? [chip('suggest', c.suggest, c.tips.suggest), chip('improve', c.improve, empty ? c.empty : c.tips.improve, empty)] : []),
        chip('summary', c.summary, c.tips.summary));
    };
    w.__tbRender = render;
    let timer = 0;
    w.__tbUiWatch = new MutationObserver(() => { // Teams changes the page constantly: at most three times a second
      if (!timer) timer = window.setTimeout(() => { timer = 0; render(); }, 300);
    });
    // class and style: a change of Teams' theme recolours its header
    w.__tbUiWatch.observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['class', 'style'] });
    render();
  }, [ACT, BAR, BOX, icon]);
}
// The app's state for its buttons in the page: what they say, which are on, what runs
export const showUi = (ui: PageUi): Promise<void> => run(ui => { (window as any).__tbUi = ui; (window as any).__tbRender?.(); }, ui);
// The page leaves for a call window (calls.ts popOutCall): its watcher stops and Soroka's buttons go with it — that page's
// header is hidden there, and its reports would move the main window's new page.
export const stopUi = (): Promise<void> => run(() => {
  const w = window as any;
  w.__tbUiWatch?.disconnect(); clearTimeout(w.__tbBarLost); w.__tbRender = null; w.__tbUi = null;
  document.querySelectorAll('[data-tb-chips], [data-tb-group], [data-tb-logo]').forEach(e => e.remove());
  document.querySelectorAll('[data-tb-bar]').forEach(e => e.removeAttribute('data-tb-bar'));
});

// ---------- calls: an incoming call, or a meeting that started, as Teams shows it in its page ----------
// Teams' toast for them, at the bottom right, marks its parts with data-testid (its notifications app, read in its bundle):
// a call is calling-notification (or its compact layout), its round buttons in calling-actions, labelled by aria-label only;
// a started meeting is meeting-notification, its buttons in msn-actions, and the toast's ✕ (dismiss-toast-button) declines it.
// (reject-call-button / reject-meeting-button, looked for before, belong to a meeting room console: a web Teams never shows them.)
// The buttons are taken as they are, with Teams' own labels, and numbered for ringClick.
// title/message: the toast's text lines that are not button labels (the caller first). outline: its data-testid ids, for the log.
export type Ring = { kind: 'call' | 'meeting'; title: string; message: string; buttons: { label: string; reject: boolean }[]; outline: string };
const RING = 'tbRing';
const TOAST = ['calling-notification', 'calling-notification-compact-layout', 'meeting-notification'].map(id => `[data-testid="${id}"]${OUT}`).join(', ');
export async function watchRings(onChange: (ring: Ring | null) => void) {
  await bind(RING, payload => {
    const r = payload ? JSON.parse(payload) : null;
    if (!r) return onChange(null);
    if (Array.isArray(r.buttons)) onChange({ kind: r.kind === 'call' ? 'call' : 'meeting', title: String(r.title ?? ''), message: String(r.message ?? ''),
      buttons: r.buttons.map((b: { label?: unknown; reject?: unknown }) => ({ label: String(b?.label ?? ''), reject: b?.reject === true })), outline: String(r.outline ?? '') });
  });
  await run(([name, toastSel]) => {
    const w = window as any;
    w.__tbRing?.disconnect();
    const read = () => {
      const toast = document.querySelector(toastSel) as HTMLElement | null;
      if (!toast?.getClientRects().length) return '';
      const kind = toast.getAttribute('data-testid') === 'meeting-notification' ? 'meeting' : 'call';
      // the text of an element outside its buttons, a line per text node (the caller's name, "Video call")
      const texts = (e: Element) => {
        const out: string[] = [];
        const walk = document.createTreeWalker(e, NodeFilter.SHOW_TEXT);
        for (let n = walk.nextNode(); n; n = walk.nextNode()) if (!n.parentElement?.closest('button') && n.textContent!.trim()) out.push(n.textContent!.trim());
        return out;
      };
      const seen = (b: Element | null | undefined) => !!b?.getClientRects().length;
      const buttons = [...toast.querySelectorAll('[data-testid="calling-actions"] button, [data-testid="msn-actions"] button')].filter(seen) as HTMLElement[];
      const labels = buttons.map(b => (b.getAttribute('aria-label') || b.innerText || '').trim());
      // a call's decline button is told only by its label, in the uk/ru/en Teams the app keeps ("Відхилити виклик", "Отклонить звонок", "Decline call");
      // a meeting has none in the toast: its ✕ closes the toast, which is how Teams declines it
      let reject = buttons.findIndex((_b, i) => /^(відхилити|отклонить|decline)/i.test(labels[i]));
      const close = toast.closest('[data-testid="notification-wrapper"]')?.querySelector('[data-testid="dismiss-toast-button"]') as HTMLElement | null;
      if (kind === 'meeting' && seen(close)) { reject = buttons.push(close!) - 1; labels.push((close!.getAttribute('aria-label') || '').trim()); }
      buttons.forEach((b, i) => b.setAttribute('data-tb-ring', String(i)));
      const lines = texts(toast);
      return JSON.stringify({ kind, title: lines[0] || '', message: lines.slice(1, 3).join(' · '),
        buttons: buttons.map((_b, i) => ({ label: labels[i], reject: i === reject })),
        outline: [toast, ...toast.querySelectorAll('[data-testid]')].map(e => e.getAttribute('data-testid')).join(' ') });
    };
    let last = read(), timer = 0;
    if (last) w[name]?.(last);
    w.__tbRing = new MutationObserver(() => { // a call must not wait: look three times a second at most
      if (timer) return;
      timer = window.setTimeout(() => { timer = 0; const now = read(); if (now !== last) { last = now; w[name]?.(now); } }, 300);
    });
    w.__tbRing.observe(document.body, { subtree: true, childList: true });
  }, [RING, TOAST]);
}
// A button of the ringing toast, by its number in Ring.buttons: a real click, so the veil of a summary reading goes first.
export async function ringClick(i: number) {
  if (veiling) { await veiling.page.removeInsertedCSS(veiling.key).catch(() => {}); veiling = null; }
  await click(`[data-tb-ring="${i}"]${OUT}`); // a message could carry one too
}
// Teams' own toasts in this page, of any kind, reported once each — only logged for now, to learn what Teams shows here.
// Each sits in a notification-wrapper; its content's data-testid is the kind (normal-notification for a message,
// calling-notification, meeting-notification, calling-notification-compact-layout). title: the content's first text line.
export type TeamsToast = { kind: string; title: string; outline: string };
const TOASTS = 'tbToasts';
export async function watchToasts(onToast: (toast: TeamsToast) => void) {
  await bind(TOASTS, payload => {
    const x = JSON.parse(payload);
    onToast({ kind: String(x.kind ?? ''), title: String(x.title ?? ''), outline: String(x.outline ?? '') });
  });
  await run(([name, sel]) => {
    const w = window as any;
    w.__tbToasts?.disconnect();
    const seen = new WeakSet<Element>();
    let timer = 0;
    const look = () => {
      timer = 0;
      for (const e of document.querySelectorAll(sel)) {
        if (seen.has(e) || !e.getClientRects().length) continue;
        seen.add(e);
        const content = e.querySelector('[data-testid$="-notification"], [data-testid$="-compact-layout"]') ?? e;
        let title = '';
        const walk = document.createTreeWalker(content, NodeFilter.SHOW_TEXT);
        for (let n = walk.nextNode(); n && !title; n = walk.nextNode()) if (!n.parentElement?.closest('button')) title = n.textContent!.trim();
        w[name]?.(JSON.stringify({ kind: content.getAttribute('data-testid') || '', title: title.slice(0, 80),
          outline: [...e.querySelectorAll('[data-testid]')].map(x => x.getAttribute('data-testid')).join(' ').slice(0, 300) }));
      }
    };
    w.__tbToasts = new MutationObserver(() => { if (!timer) timer = window.setTimeout(look, 500); });
    w.__tbToasts.observe(document.body, { subtree: true, childList: true });
    look();
  }, [TOASTS, `[data-testid="notification-wrapper"]${OUT}`]);
}

// "Join" after Teams' own meeting toast is gone: the meeting chat named in the toast, and its Join button.
// The longest title in it: a meeting named with a part of another's title ("Daily") does not take its place.
export async function joinMeeting(text: string) {
  const c = (await chatList()).filter(x => x.type === 'meeting' && x.title && text.includes(x.title)).sort((a, b) => b.title.length - a.title.length)[0];
  if (!c) return false;
  await openChat(c);
  const join = `[data-tid="chat-join-button"]${OUT}`; // the chat header's, not a button-like thing in a message
  if (!(await exists(join, 3000))) return false;
  await click(join);
  return true;
}
// A call or meeting in progress: its hang-up button ("Leave" is #hangup-button in a live call; the compact view's one not seen yet)
export const HANGUP = `[data-tid*="hangup" i]${OUT}, [id*="hangup" i]${OUT}`;
export const LEAVE = `#hangup-button${OUT}`;
// ...or the screen before joining a meeting (camera and microphone choice, the lobby): its parts are "prejoin-…"
export const CALL_UI = `${HANGUP}, [data-tid^="prejoin-"]${OUT}`;
export const inCall = (): Promise<boolean> => run(s => !!document.querySelector(s), CALL_UI);
// Told at once when a call or a meeting's join screen comes or goes in the page, as the rings are:
// with what matched (its data-tid or id, for the log), or '' when it is gone.
const CALL = 'tbCall';
export async function watchCall(onChange: (on: string) => void) {
  await bind(CALL, onChange);
  await run(([name, sel]) => {
    const w = window as any;
    w.__tbCall?.disconnect();
    const read = () => { const e = document.querySelector(sel); return e ? e.getAttribute('data-tid') || e.id || '?' : ''; };
    let last = read(), timer = 0;
    if (last) w[name]?.(last);
    w.__tbCall = new MutationObserver(() => {
      if (timer) return;
      timer = window.setTimeout(() => { timer = 0; const now = read(); if (now !== last) { last = now; w[name]?.(now); } }, 300);
    });
    w.__tbCall.observe(document.body, { subtree: true, childList: true });
  }, [CALL, CALL_UI]);
}

// ---------- Teams asking to sign in again ----------
// Its banner ("Sign in again… a request from your IT department or Teams") has a Sign in button that here only retries
// Teams' quiet sign-in in the page, which Microsoft refuses when it wants the password again: nothing happens. Told when
// the button comes ('shown') and goes ('gone'), the state as it is at every (re)attach, and its click ('click'), so the
// app says so even with the window hidden and opens Microsoft's sign-in itself.
// ponytail: the button is told by its label in a ru/uk/en Teams (outside messages); another language: the banner as it is.
const SIGNIN = 'tbSignIn';
export async function watchSignIn(onChange: (what: 'shown' | 'gone' | 'click') => void) {
  await bind(SIGNIN, what => { if (what === 'shown' || what === 'gone' || what === 'click') onChange(what); });
  await run(([name, msg]) => {
    const w = window as any;
    const LABEL = /^(войти|увійти|sign in)$/i;
    const ours = (b: Element | null) => !!b && !b.closest(msg) && LABEL.test((b.textContent ?? '').trim());
    if (w.__tbSignIn) document.removeEventListener('click', w.__tbSignIn, true);
    w.__tbSignIn = (e: Event) => { if (ours((e.target as Element | null)?.closest?.('button') ?? null)) w[name]?.('click'); };
    document.addEventListener('click', w.__tbSignIn, true);
    const read = () => [...document.querySelectorAll('button')].some(b => ours(b) && b.getClientRects().length) ? 'shown' : 'gone';
    w.__tbSignInSeen?.disconnect();
    let last = read(), timer = 0;
    w[name]?.(last);
    w.__tbSignInSeen = new MutationObserver(() => { // once a second at most: the banner is no call
      if (timer) return;
      timer = window.setTimeout(() => { timer = 0; const now = read(); if (now !== last) { last = now; w[name]?.(now); } }, 1000);
    });
    w.__tbSignInSeen.observe(document.body, { subtree: true, childList: true });
  }, [SIGNIN, MSG]);
}

// ---------- the compact view of a call: what the call window's page shows, and its buttons ----------
// The call window's page is not the one this module is attached to: its scripts run through its own webContents.
// Teams' meeting bar has no raise-hand button of its own (only its compact view has): it is in the reactions menu.
// ponytail: on and off are told by Teams' labels in a ru/uk/en Teams (as call.html tells a video button), null in another;
// the speaker by a "… говорит / is speaking" label and the people by Teams' avatars, both seen only in Teams' compact view:
// in the full meeting view either may stay empty. The log line of popOutCall's first look says which were found.
export type CallView = { mic: boolean | null; video: boolean | null; hand: boolean | null; count: number; speaker: string; people: string[]; probe: string };
const CALL_SEL = { mic: `#microphone-button${OUT}`, video: `#video-button${OUT}`, hand: `#raisehands-button${OUT}`, react: `#reaction-menu-button${OUT}`,
  roster: `#roster-button${OUT}`, msg: MSG };
const runOn = (page: WebContents, fn: (arg: any) => unknown, arg: unknown): Promise<any> =>
  within(page.executeJavaScript(`(${fn})(${JSON.stringify(arg)})`, true), 5_000, 'a script in the call window');
export const callView = (page: WebContents): Promise<CallView> => runOn(page, sel => {
  const seen = (e: Element | null) => e && e.getClientRects().length ? e : null;
  const label = (e: Element | null) => (e?.getAttribute('aria-label') ?? '').trim();
  // "Включить микрофон" while it is off, "Отключить микрофон" while on (uk Увімкнути/Вимкнути, en Unmute/Mute, Turn … on/off)
  const on = (e: Element | null) => !e ? null : /включить|увімкн|unmute|turn( \w+)? on\b/i.test(label(e)) ? false
    : /отключить|выключить|вимкн|\bmute\b|turn( \w+)? off\b/i.test(label(e)) ? true : null;
  const mic = seen(document.querySelector(sel.mic)), video = seen(document.querySelector(sel.video)), hand = seen(document.querySelector(sel.hand));
  const roster = document.querySelector(sel.roster);
  let speaker = '';
  for (const e of document.querySelectorAll('[aria-label]')) {
    const m = /^(.+?)[,.]?\s+(говорит|говорить|is speaking)$/i.exec(label(e));
    if (m && seen(e) && !e.closest(sel.msg)) { speaker = m[1]; break; }
  }
  const people = new Map<string, string>(); // name → initials, as Teams shows them
  for (const a of document.querySelectorAll('.fui-Avatar')) {
    if (!seen(a) || a.closest(sel.msg)) continue;
    const name = label(a) || a.querySelector('img')?.getAttribute('alt') || '';
    const initials = a.querySelector('.fui-Avatar__initials')?.textContent?.trim()
      || name.split(/\s+/).filter(w => /^\p{L}/u.test(w)).slice(0, 2).map(w => w[0].toUpperCase()).join('');
    if (name && initials && !people.has(name)) people.set(name, initials);
  }
  return { mic: on(mic), video: on(video), hand: hand ? /опуст|lower/i.test(label(hand)) : null, count: Number(/\d+/.exec(roster?.textContent ?? '')?.[0] ?? 0),
    speaker, people: [...people.values()],
    probe: `mic ${mic ? [...mic.attributes].map(a => a.name === 'data-state' ? `data-state=${a.value}` : a.name).join(' ') : 'none'}; video ${video ? 'yes' : 'none'}; `
      + `hand ${hand ? 'in the bar' : document.querySelector(sel.react) ? 'in the reactions menu' : 'none'}; roster ${roster ? 'yes' : 'none'}` };
}, CALL_SEL);
// Presses Teams' own button: mic, video, hand (through the reactions menu when the bar has none), roster (opens the people
// pane, never closes it). False when Teams has no such button.
export const callPress = (page: WebContents, what: 'mic' | 'video' | 'hand' | 'roster'): Promise<boolean> => runOn(page, async ([what, sel]) => {
  const seen = (s: string) => [...document.querySelectorAll(s)].find(e => e.getClientRects().length) as HTMLElement | undefined;
  if (what !== 'hand') {
    const b = seen(sel[what]);
    if (b && !(what === 'roster' && b.getAttribute('aria-expanded') === 'true')) b.click();
    return !!b;
  }
  let hand = seen(sel.hand);
  const menu = hand ? undefined : seen(sel.react);
  if (menu) { menu.click(); for (let i = 0; i < 20 && !(hand = seen(sel.hand)); i++) await new Promise(r => setTimeout(r, 100)); }
  hand?.click();
  if (menu?.getAttribute('aria-expanded') === 'true') menu.click(); // the menu it opened
  return !!hand;
}, [what, CALL_SEL]);
// ---------- Jira: a ticket's status and name in the chat's messages ----------
// In a message's text (not in link previews, not in the compose box):
// - a link to <base>/browse/KEY gets data-tb-jira, which the app's style draws after the link;
// - a bare key ("PROJ-123") Jira knows is underlined and opens in Jira on a click, and the text goes after the run of
//   text it stands in, in an empty element of the app's own, also drawn by the style.
// Teams draws its messages with React, which keeps its own text nodes: they are never cut or changed. The underline is a
// CSS highlight (no element), the click is caught on the document, the app's elements are only added next to Teams' own;
// and since a style draws all the texts, the messages' text as read (innerText: the app's own reading) stays as it was.
// Keys with no text yet, or an expired one, go to onKeys, jiraText brings the answer; a click on a key goes to onOpen.
// base '' takes it all away.
const JIRA = 'tbJira', JIRA_OPEN = 'tbJiraOpen';
export const JIRA_KEY = /^[A-Z][A-Z0-9]{1,9}-\d{1,6}$/;
export async function watchJira(base: string, onKeys: (keys: string[]) => void, onOpen: (key: string) => void) {
  await bind(JIRA, payload => {
    const keys: unknown = JSON.parse(payload);
    if (Array.isArray(keys)) onKeys(keys.filter(k => typeof k === 'string' && JIRA_KEY.test(k)).slice(0, 50));
  });
  let opened = 0; // one ticket a second at most, whatever the page sends
  await bind(JIRA_OPEN, key => { if (Date.now() - opened > 1000) { opened = Date.now(); onOpen(key); } });
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

// The user's own status as their avatar button names it ("Ваш профиль, состояние Не беспокоить"), in Teams' language.
export const myStatus = (): Promise<string> => run(() => document.getElementById('idna-me-control-avatar-trigger')?.getAttribute('aria-label') ?? '');
// ---------- theme ----------
// Teams draws itself in the scheme the app set: its page background is dark exactly when prefers-color-scheme says dark.
export const themeFollowed = (): Promise<boolean> => run(() => {
  const [r, g, b, a = 1] = (getComputedStyle(document.body).backgroundColor.match(/[\d.]+/g) ?? []).map(Number);
  return !a || (r + g + b < 384) === matchMedia('(prefers-color-scheme: dark)').matches; // a transparent body tells nothing
});
const step = async (sel: string, what: string, timeout = 8000) => {
  if (!(await exists(sel, timeout))) throw fail(502, `${what} not found in Teams`);
  await click(sel);
};
// A section of Teams' settings (… → Settings → section); the caller goes back to the chats.
async function settingsSection(section: string) {
  await click('[data-tid="more-options-header"]');
  await step('[data-tid="settings-button-menu"]', 'Settings', 3000);
  await step(`[data-tid="${section}"]`, `the "${section}" settings`);
}
// Teams' own theme set to "as the operating system" (… → Settings → Appearance), so it follows the app's theme from then on.
export async function followSystemTheme() {
  try {
    await settingsSection('appearance');
    await step('[data-tid="appearance-settings-theme-selector-dropdown"]', 'the theme choice');
    await step('[data-tid="appearance-settings-theme-selector-dropdown-followOsTheme-dropdown-option"]', 'the "as the system" theme', 3000);
  } finally { await showChats(); }
}
// Teams' app language (Settings → General → "Language and regional formats"): the locale's option ("uk-UA"), then "Save and
// restart", and Teams reloads its page in it. Its switch "Use the Microsoft 365 account's language" goes off first: on, it keeps
// the account's language. false: Teams has that language already, nothing changed.
export async function setLanguage(locale: string) {
  let applied = false;
  try {
    await settingsSection('general');
    if (!(await exists('[data-tid="language-dropdown-combobox"]', 8000))) throw fail(502, 'the language choice not found in Teams');
    const SWITCH = '[data-tid="language-switch-toggle-switch"]';
    if (await run(s => document.querySelector(s)?.getAttribute('data-checked') === 'true', SWITCH)) await click(SWITCH);
    await click('[data-tid="language-dropdown-combobox"]');
    const option = `[data-tid="language-dropdown-${locale.toLowerCase()}-combobox-option"]`;
    if (!(await exists(option, 4000))) { await press('Escape'); throw fail(404, `Teams has no language ${locale}`); }
    if (await run(s => document.querySelector(s)?.getAttribute('aria-selected') === 'true', option)) { await press('Escape'); return false; }
    await click(option);
    await step('[data-tid="apply-language-change-button"]', '"Save and restart"', 4000);
    applied = true;
    return true;
  } finally { if (!applied) await showChats().catch(() => {}); } // applied: the page reloads by itself
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
const HEADER = `[id^="chat-header-"]${OUT}`;
export const currentChatId =(): Promise<string | undefined> => run(h => document.querySelector(h)?.id.replace('chat-header-', ''), HEADER);
const header = ([h, id]: string[]) => [...document.querySelectorAll(h)].some(e => e.id === `chat-header-${id}`);
const isOpen = (c: ChatRef): Promise<boolean> => run(header, [HEADER, c.id]);
export async function openChat<T extends ChatRef>(c: T) {
  if (await isOpen(c)) return c;
  await showChats();
  // right after load the list still shifts and a click can miss, so retry
  for (let attempt = 1; ; attempt++) {
    await click(`[role="treeitem"][data-fui-tree-item-value$="|${c.id}"]`);
    if (await waitFor(header, [HEADER, c.id], 5000)) break;
    if (attempt === 3) throw fail(502, `chat "${c.title}" did not open`);
  }
  await exists('[id^="message-body-"]');
  return c;
}
// The chat in the user's sight and where in it: the message at the top of the view and how far down it sat (no spot at the
// newest messages, where Teams opens a chat itself). None under Soroka's "select a chat" (pickHint).
export type LeftOff = { id: string; spot?: { mid: string; at: number } };
export const leftOff = (): Promise<LeftOff | undefined> => run(h => {
  const id = !document.documentElement.hasAttribute('data-soroka-pick') && document.querySelector(h)?.id.replace('chat-header-', '');
  if (!id) return undefined;
  const v = document.querySelector('[data-tid="message-pane-list-viewport"]');
  if (!v || v.scrollHeight - v.scrollTop - v.clientHeight < 50) return { id };
  const top = v.getBoundingClientRect().top;
  const m = [...v.querySelectorAll<HTMLElement>('[data-tid="chat-pane-message"][data-mid]')].find(e => e.getBoundingClientRect().top >= top);
  return { id, spot: m && { mid: m.dataset.mid, at: Math.round(m.getBoundingClientRect().top - top) } };
}, HEADER);
// Back there: Teams opens a chat at its newest messages and loads older ones as the view nears the top, so it goes up
// until the message is there (a few loads), then puts it where it was.
export async function backTo(c: LeftOff) {
  await openChat({ id: c.id, title: c.id });
  for (let i = 0; c.spot && i < 8; i++, await sleep(700)) if (await run(({ mid, at }) => {
    const v = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (!v) return false;
    const m = v.querySelector(`[data-mid="${CSS.escape(mid)}"]`);
    if (!m) { v.scrollTop = 0; return false; }
    v.scrollTop += m.getBoundingClientRect().top - v.getBoundingClientRect().top - at;
    return true;
  }, c.spot)) return;
}
// Soroka's own "nothing open" (a picture, a title, a line) over the self-chat a hidden window parks on, as Teams shows on
// start; the parked chat's mark in the list goes too: its background, its brighter text and its "⋯", in the colour of a chat
// at rest. Empty texts: a blank cover. The first click outside it (the user's, or Soroka's own in the list) or key takes it
// away; null, a chat opened for the user.
export const pickHint = (hint: { picture: string; title: string; line: string } | null) => run(hint => {
  const root = document.documentElement;
  if (!hint) return root.removeAttribute('data-soroka-pick');
  let s = document.getElementById('soroka-pick');
  if (!s) {
    s = document.head.appendChild(document.createElement('style'));
    s.id = 'soroka-pick';
    const off = (e: Event) => { if (e.type === 'keydown' || !(e.target as Element).closest?.('[data-tid="app-layout-area--main"]')) root.removeAttribute('data-soroka-pick'); };
    addEventListener('pointerdown', off, true);
    addEventListener('keydown', off, true);
  }
  const rest = [...document.querySelectorAll('[role="treeitem"][data-item-type="chat"][tabindex="-1"]')]
    .find(e => !(e.getAttribute('aria-labelledby') || '').includes('chat_list_unread_text'))?.querySelector('[id^="title-chat-list-item_"]');
  const on = 'html[data-soroka-pick] [data-tid="app-layout-area--main"]', item = 'html[data-soroka-pick] [role="treeitem"][data-item-type="chat"][tabindex="0"]';
  s.textContent = `html[data-soroka-pick] [data-tid="app-layout-area--header"],${on}{visibility:hidden}
    ${on}::before{visibility:visible;content:${hint.picture ? `url(${hint.picture}) "\\A" ` : ''}${JSON.stringify(hint.title)};white-space:pre-line;position:absolute;left:0;right:0;top:calc(50% - 130px);text-align:center;font-size:16px;font-weight:600;line-height:1.6}
    ${on}::after{visibility:visible;content:${JSON.stringify(hint.line)};position:absolute;left:0;right:0;top:calc(50% + 72px);text-align:center;font-size:14px;opacity:.75}
    ${item} .fui-TreeItemLayout{background:transparent!important}
    ${item} .fui-TreeItemLayout *{color:${rest ? getComputedStyle(rest).color : 'inherit'}!important}
    ${item} button{visibility:hidden}`;
  root.setAttribute('data-soroka-pick', '');
}, hint);

// Own messages carry Fluent's ChatMyMessage class; consecutive messages may show the author only once.
export const lastMessages = (n: number): Promise<Message[]> => run(n => {
  let author = '';
  return [...document.querySelectorAll('[data-tid="chat-pane-message"]')].map(m => {
    author = (m.closest('[data-tid="chat-pane-item"]')?.querySelector('[data-tid="message-author-name"]') as HTMLElement | null)?.innerText.trim() || author;
    const mid = m.getAttribute('data-mid') || '';
    // the message element itself also holds reactions ("1 реакция"); content-<id> is the text alone
    // looked for inside this message: another message cannot lend it its text
    const content = ([...m.querySelectorAll('[id^="content-"]')].find(e => e.id === `content-${mid}`) ?? m) as HTMLElement;
    // "Reply with quote" draws the quoted message as a card inside the content: its author and text are not this message's
    const card = content.querySelector('[data-tid="quoted-reply-card"]') as HTMLElement | null;
    const box = (card?.closest('[data-track-module-name="messageQuotedReply"]') ?? card) as HTMLElement | null;
    const flat = (s: string) => s.replace(/\s+/g, ' ').trim();
    const quote = card && { replyTo: flat((card.querySelector('[data-tid="quoted-reply-timestamp"]')?.previousElementSibling as HTMLElement | null)?.innerText ?? ''),
      quote: flat((card.querySelector('[data-tid="quoted-reply-preview-content"]') as HTMLElement | null)?.innerText ?? '') };
    // pictures pasted in a chat: AMSImage, a blob: src (seen live 2026-11-13); emoji are pictures too and stay out
    const images = [...content.querySelectorAll('img')].filter(i => i.src && !box?.contains(i) && !/\/emoticons\//.test(i.src) && !/Emoji/i.test(i.getAttribute('itemtype') || ''))
      .map(i => i.src);
    return { id: mid, author, mine: m.className.includes('ChatMyMessage'), text: flat(box ? content.innerText.replace(box.innerText, '') : content.innerText), ...quote,
      ...(images.length ? { images } : {}) };
  }).filter(m => m.id).slice(-n);
}, n);
// The pictures as data URLs, read by the Teams page itself: its own blob: copy, else the picture it shows already, drawn
// out of it; null for one it cannot give (gone from the page, or another site's without CORS). No token leaves Teams.
export const imageData = (srcs: string[]): Promise<(string | null)[]> => run(srcs => Promise.all((srcs as string[]).map(async src => {
  try {
    const b = await (await fetch(src)).blob();
    if (!b.type.startsWith('image/') || b.size > 10_000_000) return null; // ponytail: 10 MB a picture
    return await new Promise<string>((ok, no) => { const f = new FileReader(); f.onload = () => ok(String(f.result)); f.onerror = no; f.readAsDataURL(b); });
  } catch {
    const img = [...document.querySelectorAll('img')].find(i => i.src === src && i.complete && i.naturalWidth);
    if (!img) return null;
    try {
      const c = document.createElement('canvas');
      c.width = img.naturalWidth; c.height = img.naturalHeight;
      c.getContext('2d')!.drawImage(img, 0, 0);
      return c.toDataURL('image/png');
    } catch { return null; }
  }
})), srcs);

// Up to n latest messages, newer than `since` (epoch ms) or, with afterMine, after the user's own last message:
// Teams keeps only a window of them in the page, so scroll up and collect until enough, no more, or stopped.
// veil: a text shown over the page while it reads; the veil also takes the clicks, since a click on another chat
// would mix that chat into the history.
type HistoryOptions = { since?: number; afterMine?: boolean; signal?: AbortSignal; veil?: string };
let veiling: { page: WebContents; key: string } | null = null; // the veil on screen: a call's buttons take it away (ringClick)
const VEIL = (text: string) => `body::before { content: ''; position: fixed; inset: 0; z-index: 2147483646; background: rgba(0, 0, 0, .35); cursor: progress; }
  body::after { content: ${JSON.stringify(text)}; position: fixed; left: 50%; top: 40%; transform: translateX(-50%); z-index: 2147483647; max-width: 60vw;
    padding: 12px 18px; border-radius: 8px; background: #292929; color: #fff; font: 14px "Segoe UI", system-ui, sans-serif; box-shadow: 0 4px 16px rgba(0, 0, 0, .4); }`;
export async function history(c: ChatRef, n: number, { since = 0, afterMine, signal, veil }: HistoryOptions = {}): Promise<Message[]> {
  await openChat(c);
  // the page it went on: a call may take this page to its own window meanwhile (popOutCall), the veil must leave with it
  const page = wc, veiled = veil ? await page.insertCSS(VEIL(veil)) : undefined; // a JSON string is a valid CSS string: the title cannot break out of it
  if (veiled) veiling = { page, key: veiled };
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
  } finally {
    if (veiled) { await page.removeInsertedCSS(veiled).catch(() => {}); if (veiling?.key === veiled) veiling = null; }
  }
  const from = Math.max(since, afterMine ? lastMine() + 1 : 0);
  return [...seen.values()].filter(m => Number(m.id) >= from).sort((a, b) => Number(a.id) - Number(b.id)).slice(-n);
}

const BOX = `[id^="new-message-"]${OUT}`;
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
// A picture (or an emoji: Teams draws them as pictures too) in the box, or a file attached under it: the user's, not to go along.
// emojiOk: the box's text gives way to a new one (send's replace), and its emoji with it.
const boxHasFiles = (emojiOk = false): Promise<boolean> => run(([b, files, emojiOk]) => [...document.querySelector(b)?.querySelectorAll('img') ?? []]
  .some(i => !emojiOk || !(/\/emoticons\//.test(i.src) || /Emoji/i.test(i.getAttribute('itemtype') || ''))) || !!document.querySelector(files), [BOX, ATTACHED, emojiOk]);
// ponytail: Teams' bundle names its file cards "file-chiclet" and "attachments-…"; not yet seen with a file attached live
const ATTACHED = '[data-tid="chat-pane-compose-message-footer"] :is([data-tid*="chiclet" i], [data-tid*="attachment" i])';

// ---------- a message's hover bar: reactions and quoted replies ----------
// The other side's latest message in the open chat: the one a toast announced, or a newer one.
const theirLast = async () => (await lastMessages(30)).reverse().find(m => !m.mine);
// Rests the mouse on a message, so Teams shows its hover bar, and clicks a button of that bar. keepOn: a toggle (a quick
// reaction) already on is left alone — a click would take the user's reaction off.
// ponytail: "on" read from aria-pressed / aria-checked, as Fluent's toggle buttons mark it; not yet seen on a reacted message
async function hoverBar(mid: string, sel: string, keepOn = false) {
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
    if (await exists(sel, 3000)) {
      if (keepOn && await run(s => { const b = document.querySelector(s); return b?.getAttribute('aria-pressed') === 'true' || b?.getAttribute('aria-checked') === 'true'; }, sel)) return;
      return click(sel);
    }
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
  if (QUICK[emoticon]) { await hoverBar(m.id, `[data-tid="message-actions-${QUICK[emoticon]}"]`, true); return { chat: c.title, reaction: emoticon }; }
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
  if (had !== plain(replace ?? '') || (!replace && await boxQuote()) || await boxHasFiles(!!replace)) // a picture or a file would go along
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
  if (plain(await boxText()) !== plain(text)) { // never the text itself in the error: errors go to the log
    await clearBox();
    throw fail(500, 'the compose box does not hold the text as typed, not sent');
  }
  if (!(await isOpen(c))) throw fail(409, 'chat switched before send, not sent');
  if (!(await ensureBoxFocus())) { await clearBox(); throw fail(500, 'compose box lost the focus, not sent'); } // not left there as "the user's"
  await press('Enter');
  // Teams empties the box as the message goes: then it went, however long the chat takes to show it. Still there: Teams did not
  // take the Enter, and the text (the app's own) is wiped, so sending again cannot double it.
  if (!(await waitFor(b => !(document.querySelector(b) as HTMLElement | null)?.innerText.trim(), BOX, 15_000))) {
    await clearBox();
    throw fail(500, 'Teams did not take the message, not sent');
  }
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

// With `text` the GIF goes in after it, as one message. The user's own unsent text, picture or file is never sent along or wiped.
export async function sendGif(c: ChatRef, query: string, src?: string, text?: string) {
  await openChat(c);
  if (plain(await boxText()) || await boxQuote() || await boxHasFiles())
    throw Object.assign(fail(409, 'the compose box has your unsent text, not sent'), { boxTaken: true });
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
    if (typed !== plain(text ?? '')) throw fail(500, 'the compose box does not hold the text as typed, not sent');
    if (!(await isOpen(c))) throw fail(409, 'chat switched before send, not sent');
  } catch (e) { await press('Escape'); await clearBox(); throw e; }
  if (!(await ensureBoxFocus())) throw fail(500, 'compose box lost the focus, not sent');
  await press('End', CTRL);
  await press('Enter');
  if (!(await waitFor(b => !document.querySelector(b)?.querySelector('img'), BOX))) { await clearBox(); throw fail(500, 'Teams did not take the GIF, not sent'); }
  return { chat: c.title, gif: label };
}


