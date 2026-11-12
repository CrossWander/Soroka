// The signature an installed Soroka checks before it runs an update (UPDATE_KEY in src/main.ts). The release workflow signs
// with the key kept as the secret UPDATE_SIGNING_KEY of the GitHub environment "release"; it never goes into the repository.
//   node scripts/sign-release.mjs keygen        makes the key pair once: the private key in %USERPROFILE%\.soroka-signing\update-key.pem,
//                                                the public key printed for UPDATE_KEY
//   node scripts/sign-release.mjs secret        puts the private key on the clipboard for the GitHub secret, kept out of the
//                                                clipboard history and the cloud clipboard, and clears the clipboard after the paste
//   node scripts/sign-release.mjs sign <file>   writes <file>.sig; the key comes from SOROKA_UPDATE_KEY (the secret: base64 of
//                                                its PKCS#8 DER, one line, so GitHub masks it whole in logs) or the local file
// The signature covers the file name, so one release's installer cannot be offered as another version.
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import readline from 'node:readline/promises';
import { spawnSync } from 'node:child_process';

const KEY = path.join(os.homedir(), '.soroka-signing', 'update-key.pem');
const [cmd, file] = process.argv.slice(2);
const privateKey = () => (process.env.SOROKA_UPDATE_KEY
  ? crypto.createPrivateKey({ key: Buffer.from(process.env.SOROKA_UPDATE_KEY, 'base64'), format: 'der', type: 'pkcs8' })
  : crypto.createPrivateKey(fs.readFileSync(KEY)));
// Windows PowerShell with the clipboard (STA); the script as -EncodedCommand, so no quoting can break it
const powershell = (script, input) => spawnSync('powershell.exe', ['-NoProfile', '-STA', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')],
  { input, encoding: 'utf8' });

if (cmd === 'keygen') {
  if (fs.existsSync(KEY)) { console.error(`${KEY} exists already: a new key would lock out every installed Soroka`); process.exit(1); }
  const { publicKey, privateKey: key } = crypto.generateKeyPairSync('ed25519');
  fs.mkdirSync(path.dirname(KEY), { recursive: true });
  fs.writeFileSync(KEY, key.export({ type: 'pkcs8', format: 'pem' }), { mode: 0o600 });
  console.log(`private key: ${KEY} (keep a copy offline)\nUPDATE_KEY for src/main.ts:\n${publicKey.export({ type: 'spki', format: 'pem' })}`);
} else if (cmd === 'secret') {
  const value = crypto.createPrivateKey(fs.readFileSync(KEY)).export({ type: 'pkcs8', format: 'der' }).toString('base64');
  // learn.microsoft.com, "Cloud Clipboard and Clipboard History Formats": any of these on the clipboard keeps it out of both
  const put = powershell(`Add-Type -AssemblyName System.Windows.Forms
$d = New-Object System.Windows.Forms.DataObject
$d.SetText([Console]::In.ReadToEnd().Trim())
foreach ($f in 'ExcludeClipboardContentFromMonitorProcessing', 'CanIncludeInClipboardHistory', 'CanUploadToCloudClipboard') {
  $d.SetData($f, (New-Object System.IO.MemoryStream(,[byte[]](0, 0, 0, 0))))
}
[System.Windows.Forms.Clipboard]::SetDataObject($d, $true)`, value);
  if (put.status !== 0) { console.error(put.stderr); process.exit(1); }
  console.log('The key is on the clipboard (not in Win+V history, not synced). Paste it as the secret UPDATE_SIGNING_KEY of the environment "release".');
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  await rl.question('Press Enter once it is saved, and the clipboard is cleared: ');
  rl.close();
  powershell('Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::Clear()');
  console.log('Clipboard cleared.');
} else if (cmd === 'sign' && file) {
  if (process.env.GITHUB_ACTIONS && !process.env.SOROKA_UPDATE_KEY) { console.error('no UPDATE_SIGNING_KEY in the environment "release": nothing is published'); process.exit(1); }
  const name = path.basename(file), hash = crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
  fs.writeFileSync(`${file}.sig`, `${crypto.sign(null, Buffer.from(`${name} ${hash}`), privateKey()).toString('base64')}\n`);
  console.log(`${name}: sha256 ${hash}, signed`);
} else {
  console.error('usage: node scripts/sign-release.mjs keygen | secret | sign <file>');
  process.exit(2);
}
