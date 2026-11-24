// Calls and meetings: the ring card, the call in a window of its own, its compact view, sharing a screen
import { BaseWindow, BrowserWindow, desktopCapturer, nativeImage, screen } from 'electron';
import * as teams from './teams';
import { ui, handle, logged, log, errorText, serial, state, saveState } from './base';
import { t } from './i18n';
import { cardWindow, setRingCard } from './toast';
import { showTeams, teamsView, teamsWin, closeSettings, newTeamsPage, chatsAway, chatsBack, APP_ICON, preload, TEAMS_PAGE } from './main';

// Teams rings in its own page; the app mirrors that toast — who calls, and Teams' own buttons — in a window of its own
// at the bottom right, above the other cards (toast.ts placeRing), for as long as Teams rings. Teams shows a started meeting for a few seconds only: the app keeps it
// 5 minutes, and "Join" then goes through the meeting chat. Accepting or joining shows the Teams window with the call.
export let ring: (teams.Ring & { live: boolean }) | null = null; // live: Teams still shows its toast
let ringWin: BrowserWindow | null = null;
let ringTimer: NodeJS.Timeout | undefined;
let ringHidden = ''; // the caller whose toast the user hid with its cross: not shown again while that call rings
export function onRing(r: teams.Ring | null) {
  if (r) {
    if (r.title === ringHidden) return;
    if (!ring?.live) log(`ring: ${r.kind}, buttons: ${r.buttons.map(b => (b.reject ? 'reject' : 'other')).join('/')}, tids: ${r.outline.slice(0, 400)}`);
    clearTimeout(ringTimer);
    ring = { ...r, live: true };
    return showRing();
  }
  ringHidden = '';
  if (!ring?.live) return;
  log(`ring over: ${ring.kind}`);
  if (ring.kind === 'call') return closeRing(); // answered, declined, missed or taken on another device
  ring.live = false;
  showRing();
  ringTimer = setTimeout(closeRing, 5 * 60_000);
}
const ringView = () => ring && { kind: ring.kind, title: ring.title, message: ring.message, live: ring.live, buttons: ring.buttons };
function showRing() {
  if (ringWin) return ringWin.webContents.send('ring', ringView());
  const win = ringWin = cardWindow({ title: t('call.window') });
  win.loadFile(ui('call.html'));
  win.webContents.on('did-finish-load', () => win.webContents.send('ring', ringView())); // also after a reload (a language change)
  win.on('closed', () => { if (ringWin === win) { ringWin = null; setRingCard(null); } });
}
handle('ring:size', (_e, height: number) => { // bottom right with the other cards, above them (toast.ts placeRing)
  if (!ringWin) return;
  setRingCard(ringWin, height);
  ringWin.setAlwaysOnTop(true, 'screen-saver'); // above full-screen windows too, like Teams' own call toast
  ringWin.showInactive();
});
function closeRing() { clearTimeout(ringTimer); ring = null; ringWin?.close(); }
async function ringAction(i: number) {
  const r = ring;
  const b = r?.buttons[i];
  if (!r || !b) return closeRing();
  try {
    if (r.live) await teams.ringClick(i); // not through serial(): a call cannot wait for a summary that reads history
    else if (!b.reject) await serial(() => teams.joinMeeting(`${r.title} ${r.message}`));
  } catch (e) { showTeams(); throw e; } // the card stays and says why; Teams' own buttons are in its window

  // an accepted call, or a meeting's join screen, comes up in its own window (onCall): if it never does, show where it is
  if (!b.reject) setTimeout(() => { if (!callWin) showTeams(); }, 5000);
  closeRing();
}

// ---------- a call in a window of its own, as in the Teams app ----------
// The call lives in the Teams page, so the whole page moves to the call window, and the main window gets a fresh page
// for the chats, loaded only once needed (main.ts chatsAway). It moves already at a meeting's join screen, as Teams' own
// app opens its meeting window there. The window closes when the call ends, closing it leaves the call.
export let callWin: BaseWindow | null = null;
export let callPage: Electron.WebContents | null = null;
export function onCall(on: string) { // either page may report: the main one a call starting, the call window's one its end
  if (on && !callWin) popOutCall(on);
  // looked at again a few seconds on: between the join screen and the call itself neither may be in the page for a moment
  // (a page that does not answer in 5 s counts as ended)
  else if (!on && callWin) setTimeout(() => void Promise.race([callPage?.executeJavaScript(`!!document.querySelector(${JSON.stringify(teams.CALL_UI)})`),
    new Promise(r => setTimeout(() => r(false), 5000))]).then(still => { if (!still) callWin?.close(); }, () => callWin?.close()), 3000);
}
function popOutCall(what: string) {
  log(`call: to its own window (${what})`);
  callSince = Date.now();
  const view = teamsView;
  closeSettings();
  teamsWin.contentView.removeChildView(view);
  // where the user left it, if that is still on a screen; the first time in the middle, as Teams' own. Never smaller than
  // Teams' join screen needs without scrolling (the camera, the microphone and camera settings, Join): 820 × 560 was not (2026-11-24)
  const was = state.callBounds, wa = was && screen.getDisplayMatching(was).workArea;
  const seen = wa && was.x < wa.x + wa.width && was.x + was.width > wa.x && was.y < wa.y + wa.height && was.y + was.height > wa.y;
  const area = seen ? wa : screen.getPrimaryDisplay().workArea;
  const width = Math.min(Math.max(seen ? was.width : 1200, 1000), area.width), height = Math.min(Math.max(seen ? was.height : 820, 760), area.height);
  const at = seen ? { x: Math.max(area.x, Math.min(was.x, area.x + area.width - width)), y: Math.max(area.y, Math.min(was.y, area.y + area.height - height)) } : {};
  const win = callWin = new BaseWindow({ ...at, width, height, show: false, title: t('call.window'), icon: APP_ICON });
  win.removeMenu();
  win.contentView.addChildView(view);
  const fit = () => { const [width, height] = win.getContentSize(); view.setBounds({ x: 0, y: 0, width, height }); };
  fit();
  win.on('resize', fit);
  const page = callPage = view.webContents;
  // only the call: Teams' left rail and its title bar (search, the Office menu) stay in the chats window. The rail goes with
  // its whole layout area: hidden alone, it left its grid column as an empty band on the left
  void page.insertCSS('[data-tid="app-layout-area--nav"],[data-tid="title-bar"]{display:none!important}', { cssOrigin: 'user' });
  void page.executeJavaScript(`document.documentElement.removeAttribute('data-soroka-pick')`).catch(() => {}); // the call under "select a chat" (pickHint)
  // leaving: Teams' own hang-up first, so the others see it at once, then the page goes and its memory with it
  win.on('close', () => {
    state.callBounds = win.getNormalBounds(); saveState();
    void page.executeJavaScript(`document.querySelector(${JSON.stringify(teams.LEAVE)})?.click()`).catch(() => {}) // "Leave", seen in a live call
      .finally(() => setTimeout(() => page.close(), 1000));
  });
  win.on('closed', () => {
    if (callWin === win) { callWin = null; callPage = null; log('call: window closed'); }
    miniWin?.close();
    void chatsBack().catch(e => log('call: chats page:', errorText(e))); // what came during the call comes now
  });
  // as Teams' own: out of screenshots and of a shared screen, a red dot on the taskbar button, the compact view while minimized
  win.setContentProtection(true);
  win.setOverlayIcon(CALL_DOT, t('call.on'));
  win.on('minimize', showMini);
  win.on('restore', () => miniWin?.close());
  win.show();
  void teams.stopUi().catch(() => {}); // still this page's: newTeamsPage attaches the next one
  newTeamsPage();
  chatsAway();
}
const CALL_DOT = (() => { // Teams' red (#c4314b) in BGRA, its edge smoothed
  const s = 16, px = Buffer.alloc(s * s * 4);
  for (let y = 0; y < s; y++) for (let x = 0; x < s; x++) {
    const a = Math.max(0, Math.min(1, s / 2 - 0.5 - Math.hypot(x + 0.5 - s / 2, y + 0.5 - s / 2)));
    px.set([0x4b * a, 0x31 * a, 0xc4 * a, 255 * a], (y * s + x) * 4); // premultiplied, as Windows takes it
  }
  return nativeImage.createFromBitmap(px, { width: s, height: s });
})();

// ---------- the compact view of a call, as Teams' own ----------
// While the call window is minimized: a small window above the others, gone when the call window comes back. Who speaks and
// who is in, how long the call goes (since its window opened), and Teams' own microphone, camera, hand, people and leave
// buttons; folded, the buttons alone (state.miniMode, kept for the next call). A button of its own on the taskbar.
let miniWin: BrowserWindow | null = null;
let callSince = 0;
function showMini() {
  if (miniWin || !callPage) return;
  const page = callPage;
  const title = () => page.getTitle().replace(/^\(\d+\)\s*/, '').replace(/\s*\|\s*Microsoft Teams\s*$/i, '');
  // transparent like the app's cards: the page draws the card with its shadow and asks for its size (mini:size)
  const win = miniWin = new BrowserWindow({ width: 360, height: 160, show: false, frame: false, transparent: true, hasShadow: false, resizable: false,
    maximizable: false, minimizable: false, fullscreenable: false, alwaysOnTop: true, icon: APP_ICON, title: `${t('call.mini')} | ${title()}`, webPreferences: { preload } });
  win.setContentProtection(true);
  win.setOverlayIcon(CALL_DOT, t('call.on'));
  win.on('page-title-updated', e => e.preventDefault()); // the taskbar names the call, not the page
  win.loadFile(ui('mini.html'));
  let first = true;
  const look = () => void teams.callView(page).then(v => {
    if (first) { first = false; log(`call: compact view: speaker ${v.speaker ? 'found' : 'none'}, people ${v.people.length}, count ${v.count}; ${v.probe}`); }
    if (!win.isDestroyed()) win.webContents.send('mini', { ...v, probe: undefined, title: title(), since: callSince, folded: state.miniMode === 'bar' });
  }, () => {});
  win.webContents.on('did-finish-load', look);
  const timer = setInterval(look, 1000);
  win.on('moved', () => { const [x, y] = win.getPosition(); state.miniAt = { x, y }; saveState(); });
  win.on('closed', () => { clearInterval(timer); if (miniWin === win) miniWin = null; });
}
// The page's width: where the user left it, its right edge kept; the first time at the top right of the call's screen
handle('mini:size', (e, width: number, height: number) => {
  if (!miniWin || e.sender !== miniWin.webContents) return;
  const size = { width: Math.ceil(width), height: Math.ceil(height) };
  if (miniWin.isVisible()) { const b = miniWin.getBounds(); return miniWin.setBounds({ ...size, x: b.x + b.width - size.width, y: b.y }); }
  const at = state.miniAt, wa = screen.getDisplayNearestPoint(at ?? screen.getCursorScreenPoint()).workArea;
  const seen = at && at.x >= wa.x && at.x < wa.x + wa.width && at.y >= wa.y && at.y < wa.y + wa.height;
  const top = callWin ? screen.getDisplayMatching(callWin.getNormalBounds()).workArea : wa;
  miniWin.setBounds({ ...size, ...(seen ? at : { x: top.x + top.width - size.width - 16, y: top.y + 16 }) });
  miniWin.showInactive();
});

// ---------- sharing a screen or a window in a call ----------
// Teams asks with getDisplayMedia; Electron has no picker of its own on Windows, so a small window lists the screens and
// windows, as Teams' own app does. Only the picture goes: the computer's sound would bring the others' voices back to them.
let shareWin: BrowserWindow | null = null;
let shareSources: Electron.DesktopCapturerSource[] = [];
let shareDone: ((streams: Electron.Streams | null) => void) | null = null;
export async function pickShare(request: Electron.DisplayMediaRequestHandlerHandlerRequest, callback: (streams: Electron.Streams | null) => void) {
  if (shareWin || !request.videoRequested || !TEAMS_PAGE.test(`${request.securityOrigin}/`)) return callback(null); // one at a time, Teams only
  shareSources = await desktopCapturer.getSources({ types: ['screen', 'window'], thumbnailSize: { width: 320, height: 180 } });
  shareDone = callback;
  const at = (callWin ?? teamsWin).getBounds(), width = 760, height = 560; // over the call
  const win = shareWin = new BrowserWindow({ width, height, x: Math.round(at.x + (at.width - width) / 2), y: Math.round(at.y + (at.height - height) / 2),
    show: false, title: t('share.title'), icon: APP_ICON, minimizable: false, alwaysOnTop: true, webPreferences: { preload } });
  win.removeMenu();
  win.loadFile(ui('share.html'));
  win.once('ready-to-show', () => win.show());
  win.on('closed', () => { shareWin = null; shareSources = []; shareDone?.(null); shareDone = null; }); // closed without a choice: nothing shared
}
handle('share:sources', e => e.sender !== shareWin?.webContents ? [] // the thumbnails show the screen: only to the picker
  : shareSources.map(s => ({ id: s.id, name: s.name, screen: s.id.startsWith('screen:'), thumb: s.thumbnail.toDataURL() })));
handle('share:pick', (e, id: unknown) => {
  if (e.sender !== shareWin?.webContents) return;
  const source = shareSources.find(s => s.id === id);
  shareDone?.(source ? { video: source } : null);
  shareDone = null;
  shareWin.close();
});

handle('ring:click', logged('ring', (_e, i: number) => ringAction(Number(i))));
handle('ring:dismiss', () => { if (ring?.live) ringHidden = ring.title; closeRing(); });
handle('mini:act', logged('compact view', async (e, what: string) => {
  if (!miniWin || e.sender !== miniWin.webContents || !callWin || !callPage) return;
  if (what === 'fold' || what === 'unfold') { state.miniMode = what === 'fold' ? 'bar' : ''; return saveState(); } // the page folds itself
  if (what === 'leave') return callWin.close(); // as closing the call window: Teams' "Leave", then the page goes
  if (what === 'expand' || what === 'roster') { callWin.restore(); callWin.focus(); }
  if (what === 'mic' || what === 'video' || what === 'hand' || what === 'roster') await teams.callPress(callPage, what);
}));
