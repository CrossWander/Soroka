// Local API for Claude sessions (tb.mjs); requests and answers signed with token.txt
import http from 'node:http';
import crypto from 'node:crypto';
import * as teams from './teams';
import { TOKEN, log, errorText, serial, settings, state, type Draft } from './base';
import { aiSkipped, skippedAuthor, authorOf, skipDraft, wanted, notToAi, forAi } from './filters';
import { t } from './i18n';
import { status, inChat, withLabel, addDraft, showDrafts } from './main';
import { windowAnswer } from './summary';
import { toast, confirmSend } from './toast';

// The token itself never travels: a program that took the port first (another user of this computer) would get it.
// tb.mjs signs each request (method, path, time, a one-time nonce) with HMAC-SHA256 under the token; Soroka signs its
// answer to that nonce, so tb.mjs shows nothing a stand-in on the port made up.
const sign = (...parts: string[]) => crypto.createHmac('sha256', TOKEN).update(parts.join('\n')).digest('hex');
const answering = new WeakMap<http.ServerResponse, string>(); // the request's nonce, which the answer is signed to
const json = (res: http.ServerResponse, code: number, body: unknown) => {
  const text = JSON.stringify(body), nonce = answering.get(res);
  res.writeHead(code, { 'content-type': 'application/json; charset=utf-8',
    ...(nonce ? { 'x-sig': sign(nonce, String(code), crypto.createHash('sha256').update(text).digest('hex')) } : {}) });
  res.end(text);
};
const body = (req: http.IncomingMessage) => new Promise<any>((ok, no) => {
  let b = '', big = false; req.setEncoding('utf8'); req.on('error', no);
  // a message, not a file: past 64 KB the rest is read and dropped, and the answer says why
  req.on('data', x => { if (big) return; b += x; if (b.length > 64 * 1024) { big = true; b = ''; no(Object.assign(new Error('body too large'), { status: 413 })); } });
  req.on('end', () => { try { ok(JSON.parse(b || '{}')); } catch { no(Object.assign(new Error('bad json'), { status: 400 })); } });
});
const need = (ok: unknown, what: string) => { if (!ok) throw Object.assign(new Error(`${what} required`), { status: 400 }); };
// The request's nonce when it is signed, fresh (30 s) and not seen before (a request cannot be played again), else ''.
// The request from this computer's programs only, checked before: no browser page (it sends an Origin) and no other host
// name pointed at 127.0.0.1 (a site's own DNS name, to reach the API through the user's browser).
const nonces = new Map<string, number>();
function signedBy(req: http.IncomingMessage) {
  const [nonce, ts, sig] = ['x-nonce', 'x-ts', 'x-sig'].map(h => String(req.headers[h] ?? '')), now = Date.now();
  for (const [n, at] of nonces) if (now - at > 60_000) nonces.delete(n);
  if (!/^[0-9a-f]{32}$/.test(nonce) || nonces.has(nonce) || !(Math.abs(now - Number(ts)) < 30_000)) return '';
  const want = sign(req.method ?? '', req.url ?? '', ts, nonce);
  if (sig.length !== want.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(want))) return ''; // constant time
  nonces.set(nonce, now);
  return nonce;
}

export function startApi() {
  const PORT = Number(process.env.PORT ?? 8765);
  http.createServer(async (req, res) => {
    try {
      if (req.headers.origin || ![`127.0.0.1:${PORT}`, `localhost:${PORT}`].includes(req.headers.host ?? '')) return json(res, 403, { error: 'local programs only' });
      const nonce = signedBy(req);
      if (!nonce) return json(res, 401, { error: 'not signed with token.txt' });
      answering.set(res, nonce);
      if (settings.aiOff) return json(res, 403, { error: t('toast.aiOff') }); // "Do not use AI": AI sessions get nothing from Teams either
      if (status !== 'ready') return json(res, 503, { error: `teams is not ready: ${status}` });
      const url = new URL(req.url ?? '/', 'http://x');
      const route = `${req.method} ${url.pathname}`;
      // a Claude session is AI too: the chats on the "Not to AI" list come without their messages
      // the chats on the "Not to AI" list, and the messages of the people on it, come without their text
      const hide = (c: teams.Chat) => (aiSkipped(c.title) || skippedAuthor(authorOf(c)) ? { ...c, preview: '' } : c);
      const hideDraft = ({ context, ...d }: Draft) => (aiSkipped(d.title) ? { ...d, last: '', preview: '', msgs: [] } : { ...d, last: skipDraft(d) ? '' : d.last,
        preview: skippedAuthor(authorOf({ type: 'group', preview: d.preview ?? '' })) ? '' : d.preview, msgs: d.msgs?.filter(m => !skippedAuthor(m.who ?? '')) });
      if (route === 'GET /chats') return json(res, 200, (await serial(teams.chatList)).map(hide));
      if (route === 'GET /unread') return json(res, 200, (await serial(teams.chatList)).filter(wanted).map(hide));
      if (route === 'GET /drafts') return json(res, 200, Object.fromEntries(Object.entries(state.drafts).map(([n, d]) => [n, hideDraft(d)])));
      if (route === 'POST /notify') { // a Claude session tells the user something without opening anything
        const { title, message } = await body(req);
        need(title?.trim(), 'title');
        toast({ title, message: message ?? '', long: true });
        return json(res, 200, { shown: true });
      }
      if (route === 'GET /messages') {
        const part = url.searchParams.get('chat');
        need(part, 'chat');
        const n = Math.min(Number(url.searchParams.get('n') ?? 20), 500);
        const c = await serial(() => teams.findChat(part!));
        notToAi(c.title);
        return json(res, 200, { chat: c.title, messages: forAi(await inChat(c.id, () => teams.history(c, n))) });
      }
      if (route === 'GET /summary') { // the same summary the app's window shows
        const part = url.searchParams.get('chat');
        need(part, 'chat');
        const c = await serial(() => teams.findChat(part!));
        notToAi(c.title); // the plain summary skips this check for the window, but here it goes to an AI session
        const r = await windowAnswer(c.id, c.title, `n${Number(url.searchParams.get('n') ?? 100)}`, { plain: url.searchParams.has('plain'), session: true }); // plain: no tokens
        return json(res, 200, { chat: c.title, summary: r.text, count: r.count });
      }
      if (route === 'POST /send') { // goes out once the user confirms it in the app's card (confirmSend)
        const { chat, text } = await body(req);
        need(typeof chat === 'string' && typeof text === 'string' && text.trim(), 'chat and text');
        const c = await serial(() => teams.findChat(chat));
        if (!(await confirmSend(c.title, withLabel(text)))) return json(res, 403, { error: 'not sent: the user did not confirm it' });
        return json(res, 200, await inChat(c.id, () => teams.send(c, withLabel(text))));
      }
      if (route === 'POST /draft') { // an answer card the user sends or rejects; nothing leaves without their click
        const { chat, text, note } = await body(req);
        need(typeof chat === 'string' && typeof text === 'string' && text.trim(), 'chat and text');
        const c = await serial(() => teams.findChat(chat));
        const n = addDraft({ chatId: c.id, title: c.title, note: typeof note === 'string' && note.trim() ? note : c.title, last: '', text, kind: 'request', preview: c.preview });
        toast({ title: t('toast.claudeSuggests'), message: `${c.title}: ${text}`, tone: 'accent', long: true, click: () => showDrafts(true) });
        return json(res, 200, { draft: n });
      }
      if (route === 'POST /gif-search') {
        const { query, n } = await body(req);
        need(query?.trim(), 'query');
        return json(res, 200, await serial(() => teams.gifSearch(query.trim(), Math.min(Number(n ?? 6), 20))));
      }
      if (route === 'POST /send-gif') {
        const { chat, query, src, text } = await body(req); // text: sent in the same message, before the GIF
        need(typeof chat === 'string' && typeof query === 'string' && query.trim(), 'chat and query');
        const c = await serial(() => teams.findChat(chat)), said = typeof text === 'string' && text.trim() ? withLabel(text) : undefined;
        if (!(await confirmSend(c.title, `${said ? `${said}\n` : ''}GIF: ${query.trim()}`))) return json(res, 403, { error: 'not sent: the user did not confirm it' });
        return json(res, 200, await inChat(c.id, () => teams.sendGif(c, query.trim(), typeof src === 'string' ? src : undefined, said)));
      }
      json(res, 404, { error: 'not found' });
    } catch (e) { json(res, (e as { status?: number }).status ?? 500, { error: errorText(e) }); }
  }).on('error', e => { // the port taken: Claude sessions reach that program instead, and tb.mjs refuses its answers
    log('api:', e.message);
    toast({ title: t('toast.apiBusy', { port: PORT }), message: t('toast.apiBusyText'), tone: 'error', long: true });
  }).listen(PORT, '127.0.0.1', () => log(`api on 127.0.0.1:${PORT}`));
}
