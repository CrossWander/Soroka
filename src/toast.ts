// Toasts: a card in the corner above the clock (as in AgentVitals), never takes focus
import { BrowserWindow, screen } from 'electron';
import { ui, handle, type Msg } from './base';
import { t } from './i18n';
import { signInCard } from './signin';
import { quitting, refreshTray, showDrafts, APP_ICON, APP_ICON_URL, preload, marksOf } from './main';

// reply: a message card (its draft n): reply in the card, quick reactions, the actions go to its "⋯" menu
// stay: no timeout, the card waits for the user's answer; gone: it closed without one of its actions
// stack: the mentions of one channel, in one card without a reply (key: 'chan:' + the channel)
// full: the whole message, lines kept (a text to confirm)
type Toast = { title: string; message: string; hint?: string; tone?: 'accent' | 'error'; long?: boolean; stay?: boolean; full?: boolean; click?: () => void; gone?: () => void;
  actions?: { label: string; run: () => void }[]; reply?: { n: string; variants: string[]; ai: boolean; quick: string[]; msgs: Msg[]; brief: boolean; agent: boolean; marks: string[] };
  stack?: { key: string; msgs: Msg[] } };
// The cards (a message, the sign-in, a call): transparent, the page draws the card with its shadow and asks for its height.
const CARD = 392;
export const cardWindow = (o: Electron.BrowserWindowConstructorOptions = {}) => new BrowserWindow({ width: CARD, height: 200, show: false, frame: false,
  transparent: true, hasShadow: false, resizable: false, alwaysOnTop: true, skipTaskbar: true, focusable: false, icon: APP_ICON, title: 'Soroka', webPreferences: { preload }, ...o });
// On the primary screen: bottom right, `above` px higher, or top right.
export function corner(win: BrowserWindow, height: number, above: number | 'top' = 0) {
  const { workArea: wa } = screen.getPrimaryDisplay(), h = Math.ceil(height);
  win.setBounds({ width: CARD, height: h, x: wa.x + wa.width - CARD, y: above === 'top' ? wa.y : wa.y + wa.height - h - above });
}
export let toastWin: BrowserWindow | null = null;
export let toastDraft: string | undefined; // the message card on screen: its draft, or its channel's stack key
export const carded = (n: string) => toastDraft === n || toastQueue.some(x => (x.reply?.n ?? x.stack?.key) === n); // its card is up, or waits its turn
let toastHandlers: ((() => void) | undefined)[] = [], toastGone: (() => void) | undefined;
// One card at a time: a newer one waits its turn instead of wiping out the one on screen. A burst past 5 waiting comes as
// one card counting the rest: they are in the tray's recent ones, and the messages in "Replies".
export const toastQueue: Toast[] = [];
let toastOn = false, toastGuard: NodeJS.Timeout | undefined, toastMissed = 0, toastNow: Toast | undefined;
export const keepCard = () => { clearTimeout(toastGuard); toastGuard = undefined; }; // the user is busy with it: no minute closes it
function nextToast() {
  toastOn = false; clearTimeout(toastGuard); toastDraft = undefined; toastNow = undefined;
  const gone = toastGone; toastGone = undefined; gone?.();
  let next = toastQueue.shift();
  if (!next && toastMissed) { next = { title: t('toast.more', { n: toastMissed }), message: t('toast.moreText'), long: true, click: () => showDrafts(true) }; toastMissed = 0; }
  if (next) toast(next, true);
}
export const recent: { at: number; title: string; click?: () => void }[] = []; // the last 10 toasts, for the tray menu
export function toast(item: Toast, fromQueue = false) {
  if (quitting) return; // a card now would open a new window, the quit would wait for it: closing the old card showed the next queued one
  if (!fromQueue) { // a queued one was recorded when it came
    recent.unshift({ at: Date.now(), title: item.title, click: item.click });
    recent.splice(10);
    refreshTray();
  }
  if (toastOn) { if (toastQueue.length < 5 || item.stay) toastQueue.push(item); else toastMissed++; return; } // a question to the user is never dropped
  toastOn = true; toastDraft = item.reply?.n ?? item.stack?.key; toastNow = item;
  if (!item.stay) toastGuard = setTimeout(nextToast, 60_000); // the card never reported back (its page crashed): don't block the rest
  const { click, long, stay, gone, actions = [], ...card } = item;
  toastHandlers = [click, ...actions.map(a => a.run)];
  toastGone = gone;
  const payload = { ...card, actions: actions.map(a => a.label), icon: APP_ICON_URL, life: stay ? 0 : long ? 10_000 : 4_000, // 10 s for what the user should act on
    closeLabel: t('toast.close') };
  if (toastWin) return toastWin.webContents.send('toast', payload);
  toastWin = cardWindow({ height: 160 });
  toastWin.loadFile(ui('toast.html'));
  toastWin.webContents.once('did-finish-load', () => toastWin?.webContents.send('toast', payload));
  toastWin.on('closed', () => { toastWin = null; nextToast(); });
}
handle('toast:size', (_e, height: number) => { // the card measured itself: sit it in the corner and show
  if (!toastWin) return;
  corner(toastWin, height, signInCard?.isVisible() ? signInCard.getBounds().height : 0); // the sign-in card stays below
  toastWin.setAlwaysOnTop(true, 'pop-up-menu');
  if (!toastWin.isVisible()) toastWin.showInactive(); // it also resizes a card on screen (the reply field, the menu): no second show
});
// the reply field opened: the card takes the keyboard, and is in use until it reports back, however long the user types
handle('toast:focus', () => { keepCard(); toastWin?.setFocusable(true); toastWin?.focus(); });
// the mouse on the card: the user is reading it, the next card does not take its place after the minute
handle('toast:hold', () => { if (toastGuard) { clearTimeout(toastGuard); toastGuard = setTimeout(nextToast, 10 * 60_000); } });
// the card is done: its window closes and 'closed' shows the next one in a new window. A hidden window kept ~30 MB,
// and its page drew nothing until shown again, so a card sent to it never came up.
// a closing card takes no more messages: the next one of its chat comes in a card of its own
handle('toast:hidden', () => { toastDraft = undefined; toastWin?.close(); });
handle('toast:click', (_e, action: number) => { // -1 = the card itself
  const run = toastHandlers[action + 1];
  if (run) toastGone = undefined; // answered: not "closed without an answer"
  toastDraft = undefined;
  toastWin?.close();
  run?.();
});
// A Claude session sends only what the user confirms in the app too: a card with the chat and the text, and the request
// waits for the click. No answer in 3 minutes, or the card closed: not sent.
export function confirmSend(title: string, text: string) {
  return new Promise<boolean>(ok => {
    let done = false;
    const answer = (yes: boolean) => { if (!done) { done = true; clearTimeout(timer); ok(yes); } };
    const card: Toast = { title: t('toast.claudeSends', { title }), message: text, tone: 'accent', stay: true, full: true, gone: () => answer(false),
      actions: [{ label: t('act.send'), run: () => answer(true) }, { label: t('act.dontSend'), run: () => answer(false) }] };
    // no answer in time: its card goes too, on screen or still waiting its turn (a click on it now would send nothing)
    const timer = setTimeout(() => {
      answer(false);
      const i = toastQueue.indexOf(card);
      if (i >= 0) toastQueue.splice(i, 1); else if (toastNow === card) toastWin?.close();
    }, 3 * 60_000);
    toast(card);
  });
}
// One card per chat or channel: its card on screen takes the next message, and stays on for it
export function intoCard(key: string, msgs: Msg[]) {
  if (toastDraft !== key || !toastWin) return false;
  if (toastGuard) { clearTimeout(toastGuard); toastGuard = setTimeout(nextToast, 60_000); } // not once the user is typing in it
  const wc = toastWin.webContents, send = () => wc.send('toast:msgs', { n: key, msgs, marks: marksOf(msgs) });
  if (wc.isLoading()) wc.once('did-finish-load', send); else send(); // a new card's page: after its own payload
  return true;
}
