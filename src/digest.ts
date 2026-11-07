// A summary without AI: what in the messages concerns the user, found by rules — at once, and nothing leaves the computer.
// Not a retelling (what was said and decided needs a model): what waits for the user, issues, links, files, dates, who wrote.
import type { Message } from './teams';
import { t } from './i18n';

export type Part = { id: string; title: string; msgs: Message[]; direct?: boolean; preview?: string }; // preview: an unread chat left unopened
type Opts = { me: string; when: (id: string) => string; several: boolean }; // me: the user's name as Teams shows it

// JS \b knows ASCII letters only: Cyrillic words start where no letter or digit stands before them
const START = '(?<![\\p{L}\\p{N}])';
const JIRA = /\b[A-Z][A-Z0-9]{1,9}-\d{1,6}\b/g;
const LINK = /https?:\/\/[^\s<>"')]+/g;
const FILE = /[\p{L}\p{N}_()-][\p{L}\p{N}_.()-]{0,60}\.(?:xlsx?|docx?|pdf|pptx?|csv|zip|rar|7z|png|jpe?g|txt|log|json|xml|sql)(?![\p{L}\p{N}])/giu;
// ponytail: dates as dd.mm(.yy) — not "1/2" or a version like "2.1.318"; a day only after "до/к/на/by…" ("до пятницы",
// not "завтра света не будет"); deadline words in uk/ru/en by their stems. Other languages and "next sprint" are missed.
// File names keep their last word when they have spaces.
const DAY = 'сьогодні|сегодня|завтра|післязавтра|послезавтра|today|tomorrow|tonight|понеділ|понедел|вівтор|вторник|серед|сред[уаы]|четвер|п.ятниц|пятниц|'
  + 'субот|неділ|воскрес|monday|tuesday|wednesday|thursday|friday|кінця|конца|end of';
const DUE = new RegExp(`${START}(?:(?<!\\.)(?:0?[1-9]|[12]\\d|3[01])\\.(?:0?[1-9]|1[0-2])(?:\\.(?:20)?\\d{2})?(?!\\.?\\d)|`
  + `(?:до|к|ко|на|by|until|till|before)\\s+(?:${DAY})|дедлайн|deadline|eod|терміново|срочно|asap|urgent)`, 'iu');
// ponytail: the name by stems ("Павл" finds "Павле", "Павлом"); patronymics are left out, nicknames ("Паша") are missed
const nameRule = (me: string) => {
  const stems = me.replace(/\(.*?\)/g, '').split(/\s+/).filter(w => w.length >= 4 && !/(ович|евич|івна|ївна|овна|евна)$/i.test(w))
    .map(w => w.slice(0, Math.max(4, w.length - 1)).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  return stems.length ? new RegExp(`${START}(?:${stems.join('|')})`, 'iu') : /$^/;
};
const cut = (text: string, at = 0, size = 160) => {
  if (text.length <= size) return text;
  const from = Math.max(0, Math.min(at - 40, text.length - size));
  return `${from ? '…' : ''}${text.slice(from, from + size).trim()}…`;
};
const item = (m: Message, o: Opts, at = 0) => `• ${m.mine ? t('dg.me') : m.author}${m.replyTo ? ` ↩ ${m.replyTo}` : ''} ${o.when(m.id)}: ${cut(m.text, at)}`;

// What waits for the user: in a personal chat everything after their last message; in a group, since then, a reply to
// them, their name, or a question that is not a reply to someone else
function waiting(p: Part, o: Opts, named: RegExp) {
  const last = p.msgs.map(m => m.mine).lastIndexOf(true);
  const after = p.msgs.slice(last + 1).filter(m => !m.mine && m.text);
  return (p.direct ? after : after.filter(m => named.test(m.replyTo ?? '') || named.test(m.text) || (!m.replyTo && m.text.replace(LINK, '').includes('?')))).slice(-6);
}
function links(msgs: Message[]) {
  const all = msgs.map(m => m.text).join('\n');
  const keys = new Map<string, number>();
  for (const k of all.match(JIRA) ?? []) keys.set(k, (keys.get(k) ?? 0) + 1);
  return [
    [...keys].sort((a, b) => b[1] - a[1]).slice(0, 8).map(([k, n]) => (n > 1 ? `${k} ×${n}` : k)).join(' · '),
    ...[...new Set(all.match(LINK) ?? [])].slice(-5).map(l => cut(l, 0, 90)),
    ...[...new Set((all.match(FILE) ?? []).map(f => f.trim()))].slice(-5),
  ].filter(Boolean).map(x => `• ${x}`);
}
// a date with a year and a time is a pasted quote's header ("Name 06.11.2026 9:10"), not a deadline
const STAMP = /\d{1,2}\.\d{1,2}\.\d{4},? \d{1,2}:\d{2}/g;
const dues = (msgs: Message[], o: Opts) => msgs.map(m => ({ m, at: m.text.replace(LINK, s => ' '.repeat(s.length)).replace(STAMP, s => ' '.repeat(s.length)).search(DUE) }))
  .filter(x => x.at >= 0).slice(-6).map(x => item(x.m, o, x.at));
function who(msgs: Message[]) {
  const n = new Map<string, number>();
  for (const m of msgs) { const a = m.mine ? t('dg.me') : m.author || '?'; n.set(a, (n.get(a) ?? 0) + 1); }
  return [...n].sort((a, b) => b[1] - a[1]).slice(0, 8).map(([a, k]) => `${a} ${k}`).join(' · ');
}
const section = (name: string, lines: string[]) => (lines.length ? [name, ...lines] : []);

export function text(parts: Part[], o: Opts) {
  const named = nameRule(o.me);
  const out = o.several
    // several chats: per chat what waits for the user, dates and issue keys, a few lines each
    ? parts.filter(p => p.msgs.length || p.preview).flatMap(p => {
      if (!p.msgs.length) return [p.title, `• ${cut(p.preview ?? '')}`, ''];
      const keys = [...new Set(p.msgs.map(m => m.text).join('\n').match(JIRA) ?? [])].slice(0, 6).join(' · ');
      const lines = [...waiting(p, o, named).slice(-3).map(m => item(m, o)), ...dues(p.msgs, o).slice(-2)]; // a message can be both
      return [`${p.title} · ${t('dg.count', { n: p.msgs.length })}`, ...new Set(lines),
        ...(keys ? [`• ${keys}`] : []), ''];
    })
    : parts.flatMap(p => {
      const wait = waiting(p, o, named).map(m => item(m, o));
      return [
        ...section(t('sum.onMe'), wait.length ? wait : [t('dg.nothing')]),
        ...section(t('dg.links'), links(p.msgs)),
        ...section(t('dg.due'), dues(p.msgs, o)),
        ...section(t('dg.who'), [`• ${who(p.msgs)}`]),
      ];
    });
  return [...out, '', t('dg.note')].join('\n').replace(/\n{3,}/g, '\n\n').trim();
}
