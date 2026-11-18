// Updates: a newer release on GitHub comes as a card; "Install" downloads its installer and runs it
import { app, net, shell } from 'electron';
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { log, errorText } from './base';
import { t } from './i18n';
import { toast } from './toast';

// A minute after the start, then once a day; installed copies only. Nothing about the user goes along: a plain GET.
// ponytail: tells and installs on a click; electron-updater (downloads in the background, installs on quit) when wanted
const RELEASES = 'https://github.com/CrossWander/Soroka/releases';
const newer = (a: string, b: string) => {
  const x = a.split('.').map(Number), y = b.split('.').map(Number);
  for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] > y[i];
  return false;
};
// An installer runs only with its release's signature (scripts/sign-release.mjs in the release job, the key a secret of the
// GitHub environment "release", which only tag runs get): a file swapped in a release, or put there by anything but that
// job, is not run. The signature covers the file name, so one release's installer cannot pass for another version.
const UPDATE_KEY = crypto.createPublicKey('-----BEGIN PUBLIC KEY-----\nMCowBQYDK2VwAyEAwdwj95MyhSQA3ZEf3yPmIumHKpnTfPGIxQQShJK/ub0=\n-----END PUBLIC KEY-----\n');
const signedInstaller = (name: string, exe: Buffer, sig: string) =>
  crypto.verify(null, Buffer.from(`${name} ${crypto.createHash('sha256').update(exe).digest('hex')}`), UPDATE_KEY, Buffer.from(sig.trim(), 'base64'));
let updateShown = ''; // told once a run
async function checkUpdate() {
  if (!app.isPackaged) return; // a dev run is no installed version
  try {
    // the releases' feed, not GitHub's API: an office's computers share one address, whose 60 API calls an hour run out (403)
    const opts = { headers: { 'user-agent': `Soroka/${app.getVersion()}` }, credentials: 'omit' as const };
    const r = await net.fetch(`${RELEASES}.atom`, opts);
    if (!r.ok) return log(`update check: ${r.status}`);
    const version = /\/releases\/tag\/v(\d+\.\d+\.\d+)"/.exec(await r.text())?.[1]; // the newest release comes first
    if (!version || version === updateShown || !newer(version, app.getVersion())) return;
    // the installer of that very release, from this repository only
    const url = `${RELEASES}/download/v${version}/Soroka-Setup-${version}.exe`, there = async (u: string) => (await net.fetch(u, { ...opts, method: 'HEAD' })).ok;
    if (!(await there(url))) return log(`update: ${version} has no installer yet`);
    if (!(await there(`${url}.sig`))) return log(`update: ${version} is not signed yet`); // asked again tomorrow
    updateShown = version;
    log(`update: ${version} is out`);
    toast({ title: t('update.title', { v: version }), message: t('update.text', { cur: app.getVersion() }), tone: 'accent', long: true,
      click: () => void shell.openExternal(`${RELEASES}/tag/v${version}`), // missed, it is still in the tray's recent ones
      actions: [{ label: t('act.install'), run: () => void installUpdate(url, version) },
        { label: t('act.whatsNew'), run: () => void shell.openExternal(`${RELEASES}/tag/v${version}`) }] });
  } catch (e) { log('update check:', errorText(e)); }
}
// The installer runs as on a first install, with its progress on screen; the app quits for it, and it starts the new one.
async function installUpdate(url: string, version: string) {
  toast({ title: t('update.loading', { v: version }), message: t('update.loadingText') });
  try {
    const [r, s] = await Promise.all([url, `${url}.sig`].map(u => net.fetch(u, { credentials: 'omit' }))); // GitHub sends them from its file storage
    if (!r.ok || !s.ok) throw new Error(`HTTP ${r.ok ? s.status : r.status}`);
    const name = `Soroka-Setup-${version}.exe`, exe = Buffer.from(await r.arrayBuffer());
    if (!signedInstaller(name, exe, await s.text())) throw new Error(t('update.badSig'));
    const file = path.join(app.getPath('temp'), name);
    fs.writeFileSync(file, exe);
    log(`update: running the installer of ${version}`);
    spawn(file, [], { detached: true, stdio: 'ignore' }).unref();
    app.quit();
  } catch (e) {
    log('update:', errorText(e));
    toast({ title: t('update.failed'), message: errorText(e), tone: 'error', long: true,
      actions: [{ label: t('act.whatsNew'), run: () => void shell.openExternal(`${RELEASES}/tag/v${version}`) }] });
  }
}
setTimeout(() => void checkUpdate(), 60_000);
setInterval(() => void checkUpdate(), 86_400_000);
