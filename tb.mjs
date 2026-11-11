// CLI for the Soroka app's local API (for Claude sessions):
//   node tb.mjs chats | unread | drafts | read "<chat>" [n] | summary "<chat>" [n] [plain] | send "<chat>" "<text>"
//             | draft "<chat>" "<text>" [note] | notify "<title>" ["<message>"] | gifs "<query>" [n] | gif "<chat>" "<query>" [src] [text]
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const appData = { win32: process.env.APPDATA, darwin: path.join(os.homedir(), 'Library', 'Application Support') }[process.platform]
  ?? path.join(os.homedir(), '.config');
// the data folder of Soroka, or of TeamsBridge (its name until 0.9.1) when the move has not happened yet
const tokenFile = ['Soroka', 'TeamsBridge'].map(dir => path.join(appData, dir, 'token.txt')).find(f => fs.existsSync(f));
if (!tokenFile) { console.error('Soroka has never run on this machine'); process.exit(3); }
const base = `http://127.0.0.1:${process.env.PORT ?? 8765}`;
const headers = { 'x-token': fs.readFileSync(tokenFile, 'utf8').trim(), 'content-type': 'application/json' };
const [cmd, a, b, c] = process.argv.slice(2);

const call = async (url, init) => {
  const res = await fetch(base + url, { headers, ...init }).catch(() => null);
  if (!res) { console.error('Soroka is not running'); process.exit(3); }
  const body = await res.json();
  if (!res.ok) { console.error(`${res.status}: ${body.error}`); process.exit(1); }
  return body;
};
const post = (url, body) => call(url, { method: 'POST', body: JSON.stringify(body) });

if (cmd === 'chats') {
  for (const x of await call("/chats")) console.log(`${x.unread ? "*" : " "}${x.mention ? "@" : " "} [${x.type}] ${x.title} · ${x.time} · ${x.preview}`);
} else if (cmd === 'unread') {
  for (const x of await call("/unread")) console.log(`[${x.type}${x.mention ? ", @" : ""}] ${x.title} · ${x.time} · ${x.preview}`);
} else if (cmd === 'drafts') {
  for (const [n, d] of Object.entries(await call('/drafts'))) console.log(`#${n} ${d.status} → ${d.title}: ${d.text}${d.error ? ` (${d.error})` : ''}`);
} else if (cmd === 'read' && a) {
  const r = await call(`/messages?chat=${encodeURIComponent(a)}&n=${Number(b ?? 20)}`);
  console.log(`# ${r.chat}`);
  for (const m of r.messages) console.log(`${new Date(Number(m.id)).toLocaleString('uk-UA')} | ${m.mine ? 'Я' : m.author}${m.replyTo ? ` (↩ ${m.replyTo}: «${m.quote}»)` : ''}: ${m.text}`);
} else if (cmd === 'summary' && a) {
  const r = await call(`/summary?chat=${encodeURIComponent(a)}&n=${Number(b ?? 100)}${c === 'plain' ? '&plain' : ''}`); // plain: by rules, no AI
  console.log(`# ${r.chat}
${r.summary}`);
} else if (cmd === 'send' && a && b) { // the app asks the user in a card first; waits up to 3 minutes for the click
  console.log('waiting for the user to confirm it in Soroka…');
  console.log(`sent to ${(await post('/send', { chat: a, text: b })).chat}`);
} else if (cmd === 'draft' && a && b) {
  console.log(`draft #${(await post('/draft', { chat: a, text: b, note: c })).draft} is in the popup, waiting for the user`);
} else if (cmd === 'notify' && a) {
  await post('/notify', { title: a, message: b }); console.log('shown');
} else if (cmd === 'gifs' && a) {
  (await post('/gif-search', { query: a, n: Number(b ?? 6) })).forEach((g, i) => console.log(`${i + 1}. ${g.label}\n   ${g.src}`));
} else if (cmd === 'gif' && a && b) { // confirmed in a card first, as send
  console.log('waiting for the user to confirm it in Soroka…');
  const r = await post('/send-gif', { chat: a, query: b, src: c || undefined, text: process.argv[6] }); // text goes in the same message
  console.log(`sent "${r.gif}" to ${r.chat}`);
} else {
  console.error('usage: node tb.mjs chats | unread | drafts | read "<chat>" [n] | summary "<chat>" [n] [plain] | send "<chat>" "<text>" | draft "<chat>" "<text>" [note] | notify "<title>" ["<message>"] | gifs "<query>" [n] | gif "<chat>" "<query>" [src] [text]');
  process.exit(2);
}
