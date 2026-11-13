// "Hand to agent": a chat goes to the user's AI agent as a file next to its pictures, in a folder of its own. What is in the
// messages is data for the agent, never its instructions: the file says so, and the messages sit between two fence lines
// a colleague's text cannot know, so it cannot close the block and go on as Soroka.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

// What the messages carry for a look into code and logs: found by patterns, with no AI.
export type Marks = { ray: string[]; stack: string[]; jira: string[]; code: string[] };
const KINDS = ['ray', 'stack', 'jira', 'code'] as const;
const GUID = /\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi;
const PATTERNS: Record<keyof Marks, RegExp[]> = {
  ray: [GUID, /\bray[-_ ]?id\s*[:=]?\s*([\w-]{8,})/gi], // the value after "ray-id" when it is no GUID
  stack: [/\b\w*Exception\b|\bat [\w.<>`]+(?=\()/g], // .NET: "…Exception", "at Ns.Type.Method("
  jira: [/\b[A-Z][A-Z0-9]{1,9}-\d{1,6}\b/g],
  code: [/\b[\w.-]+\.(?:cs|ts|js|proto|sql|xaml|json)\b/g],
};
export function findMarks(texts: string[]): Marks {
  const out: Marks = { ray: [], stack: [], jira: [], code: [] };
  for (const text of texts) {
    const guids = text.match(GUID) ?? [];
    for (const kind of KINDS)
      for (const re of PATTERNS[kind])
        for (const m of text.matchAll(re)) {
          const value = m[1] ?? m[0];
          if (kind === 'jira' && guids.some(g => g.includes(value))) continue; // a piece of a GUID, not an issue key
          if (!out[kind].includes(value)) out[kind].push(value);
        }
  }
  return out;
}
export const markNames = (m: Marks): (keyof Marks)[] => KINDS.filter(k => m[k].length);

// "By question": the talk the newest message belongs to — back from it while the messages follow each other closely.
// ids are epoch ms (teams.Message).
export function byQuestion<T extends { id: string }>(msgs: T[], gapMs: number): T[] {
  let i = msgs.length - 1;
  while (i > 0 && Number(msgs[i].id) - Number(msgs[i - 1].id) <= gapMs) i--;
  return msgs.slice(Math.max(i, 0));
}

// The questions asked about the chat in Soroka, and the AI's answers
export type Talk = { q: string; a: string };
// images: the picture files next to chat.md, '' for one that could not be taken from Teams
export type ChatLine = { author: string; mine: boolean; time: string; text: string; images: string[] };
export function renderChat(o: { header: string; title: string; period: string; lines: ChatLine[]; marks: Marks; talk: Talk[]; fence: string; noImage: string }): string {
  const found = markNames(o.marks).map(k => `- ${k}: ${o.marks[k].join(', ')}`);
  const block = [
    ...o.lines.map(l => `[${l.time}] ${l.author}${l.mine ? ' (me)' : ''}: ${l.text}${l.images.map(f => ` ${f ? `[${f}]` : o.noImage}`).join('')}`),
    ...(found.length ? ['', 'Found in the messages:', ...found] : []),
    ...(o.talk.length ? ['', 'Asked about them in Soroka:', ...o.talk.flatMap(x => [`Q: ${x.q}`, `A: ${x.a}`])] : []),
  ];
  // the fence goes only on its own lines: a message that has it inside cannot pass for one
  return [o.header, '', `# ${o.title}`, o.period, '', `<<< ${o.fence}`, ...block.map(l => l.split(o.fence).join('')), `>>> ${o.fence}`, ''].join('\n');
}

// The short request the agent opens with: links and Windows Terminal take no ";" (wt splits commands on it), quotes or line breaks.
const flat = (s: string) => s.replace(/[\r\n]+/g, ' ').replace(/;/g, ',').replace(/"/g, "'").trim();
export const agentPrompt = (template: string, file: string, question?: string) =>
  flat(template.replace('{file}', () => file).replace('{question}', () => flat(question ?? '').slice(0, 300)));

// Away from the app's own data (its Teams session and keys) and from the project (git would see it).
export const HANDOFF_ROOT = path.join(os.tmpdir(), 'Soroka-handoff');
const stamp = (d = new Date()) => {
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
};
export function writeHandoff(title: string, files: { name: string; data: string | Buffer }[], root = HANDOFF_ROOT): string {
  const slug = title.replace(/[<>:"/\\|?*\u0000-\u001f]+/g, ' ').trim().replace(/\s+/g, '-').slice(0, 40) || 'chat';
  fs.mkdirSync(root, { recursive: true });
  let dir = path.join(root, `${stamp()}-${slug}`);
  for (let k = 2; fs.existsSync(dir); k++) dir = path.join(root, `${stamp()}-${slug}-${k}`); // two in one second
  fs.mkdirSync(dir);
  for (const f of files) fs.writeFileSync(path.join(dir, f.name), f.data);
  return dir;
}
// The handoffs older than the kept time go; the agents' own session history is theirs, not touched here.
export function cleanHandoffs(keepMs: number, root = HANDOFF_ROOT, now = Date.now()): number {
  let gone = 0;
  for (const name of fs.existsSync(root) ? fs.readdirSync(root) : []) {
    const dir = path.join(root, name);
    try { if (fs.statSync(dir).isDirectory() && fs.statSync(dir).mtimeMs < now - keepMs) { fs.rmSync(dir, { recursive: true, force: true }); gone++; } }
    catch { /* in use or gone meanwhile: the next round takes it */ }
  }
  return gone;
}
