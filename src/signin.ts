// Teams asking to sign in again
import { BrowserWindow, session } from 'electron';
import { log, errorText, handle, ui } from './base';
import { refreshTray, PLAIN, teamsPage } from './main';
import { cardWindow, corner, placeRing } from './toast';

// Microsoft ended the sign-in (a password change, or IT asked) and Teams shows a banner whose "Sign in" does nothing here
// (teams.watchSignIn). The app says so even with its window hidden: a card above the clock that follows the sign-in
// through (ui/signin.html: asked → Microsoft's window open → signed in), "!" on the tray icon and "Sign in to Teams again"
// first in its menu while it lasts. "Remind me in an hour" or ✕ put the card away for an hour; the "!" and the item stay.
// signIn(): Microsoft's account page opens in Teams' session, hidden, and comes up when Microsoft asks for the sign-in.
// Back on the account page and staying there (the page itself goes to the sign-in first, when one is due) the user is
// signed in: the window hides at once, closes after the check and the chats reload, so Teams signs in quietly again. A call's page is left as it is:
// reloading it would drop the call.
export let signInAsked = false; // Teams' banner "sign in again" is up (onSignIn)
export let signInWin: BrowserWindow | null = null, signInCard: BrowserWindow | null = null;
let signInState: '' | 'ask' | 'wait' | 'done' = '', signInLater = 0, signInCardTimer: NodeJS.Timeout | undefined;
export function onSignIn(what: 'shown' | 'gone' | 'click') {
  if (what === 'click') return signIn();
  const asked = what === 'shown';
  if (asked === signInAsked) return;
  signInAsked = asked;
  log(`sign-in: Teams ${asked ? 'asks to sign in again' : 'is signed in'}`);
  refreshTray();
  if (asked) { if (Date.now() >= signInLater && !signInWin) showSignInCard('ask'); }
  else if (signInState === 'wait') showSignInCard('done'); // Teams got in by itself meanwhile
  else if (signInState === 'ask') signInCard?.close();
}
function showSignInCard(state: 'ask' | 'wait' | 'done') {
  signInState = state;
  clearTimeout(signInCardTimer);
  if (state === 'done') signInCardTimer = setTimeout(() => signInCard?.close(), 4000);
  if (signInCard) return signInCard.webContents.send('signin', state);
  const win = signInCard = cardWindow();
  win.loadFile(ui('signin.html'));
  win.webContents.on('did-finish-load', () => win.webContents.send('signin', signInState)); // also after a reload (a language change)
  win.on('closed', () => { if (signInCard === win) { signInCard = null; signInState = ''; placeRing(); } });
}
handle('signin:size', (e, height: number) => { // above the clock, under the message cards (toast:size leaves it room)
  if (!signInCard || e.sender !== signInCard.webContents) return;
  corner(signInCard, height);
  signInCard.setAlwaysOnTop(true, 'pop-up-menu');
  if (!signInCard.isVisible()) signInCard.showInactive();
  placeRing(); // a call card stays above it
});
handle('signin:act', (e, name: string) => {
  if (!signInCard || e.sender !== signInCard.webContents) return;
  if (name === 'go' || name === 'show') return signIn();
  if (name !== 'later') return;
  signInLater = Date.now() + 60 * 60_000;
  signInCard.close();
  setTimeout(() => { if (signInAsked && !signInCard && !signInWin && Date.now() >= signInLater) showSignInCard('ask'); }, 60 * 60_000);
});
// check: the user closed the window after Microsoft asked, maybe signed in by then: the same page, hidden, tells which.
export function signIn(check = false) {
  if (signInWin) { signInWin.show(); return signInWin.focus(); }
  log(check ? 'sign-in: window closed, checking' : 'sign-in: opening Microsoft');
  if (signInAsked || signInCard) showSignInCard('wait');
  const win = signInWin = new BrowserWindow({ ...PLAIN, width: 520, height: 680, show: false, webPreferences: { session: session.fromPartition('persist:teams') } });
  const ACCOUNT = /^https:\/\/myaccount\.microsoft\.com\//;
  let asked = false, settled = false, done: NodeJS.Timeout | undefined;
  win.webContents.on('did-navigate', (_e, url) => {
    clearTimeout(done);
    log('sign-in: at', new URL(url).host); // where a sign-in stops, next time one does
    if (!ACCOUNT.test(url)) { // Microsoft's sign-in, or the organization's own
      if (check) return win.close(); // still not signed in
      asked = true;
      if (!win.isVisible()) win.show(); // also when Microsoft asks once more after the account page
      return;
    }
    if (asked) win.hide(); // signed in: its check goes on unseen, no window hanging there for the user to close half-way
    const settle = () => {
      if (win.webContents.isLoadingMainFrame()) { done = setTimeout(settle, 2000); return; } // the page's own frames load on and on
      log(`sign-in: ${asked ? 'signed in again' : 'Microsoft did not ask'}`);
      settled = true;
      win.close();
      if (!asked && !signInAsked) return;
      teamsPage.reload(); // the banner up and Microsoft not asking: Teams' page gets in once it starts afresh
      if (signInCard) showSignInCard('done');
    };
    done = setTimeout(settle, 8000);
  });
  win.on('closed', () => {
    clearTimeout(done);
    if (signInWin === win) signInWin = null;
    if (settled || signInState !== 'wait') return;
    if (asked) signIn(true);
    else showSignInCard('ask'); // closed before signing in: back to asking
  });
  win.loadURL('https://myaccount.microsoft.com/').catch(e => { // ERR_ABORTED: the page went on to the sign-in
    if ((e as { code?: string }).code !== 'ERR_ABORTED') { log('sign-in:', errorText(e)); win.close(); }
  });
}
