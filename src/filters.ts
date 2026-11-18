// Which chats Soroka tells the user about (personal, @mentions, keywords, "Never from", "Only from"), what it keeps
// from the AI ("Not to AI") and the quiet hours.
import { settings, type Draft } from './base';
import { t } from './i18n';
import type * as teams from './teams';

const list = (s: string) => s.split(/[,;\n]/).map(x => x.trim().toLowerCase()).filter(Boolean);
const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
// A name from a list as a whole word in the text: "Анна" is not in "Ганна", "pos" is in "POS team".
const hasName = (text: string, names: string) => list(names).some(n => new RegExp(`(^|[^\\p{L}\\p{N}])${escapeRe(n)}($|[^\\p{L}\\p{N}])`, 'iu').test(text));
// "Not to AI": the chats and people whose messages never go to the AI — no drafts, summaries, style or Claude sessions.
// A chat by its title; a person also by the messages they write in other chats, which leave the transcript (forAi).
// Matched by any part, unlike the notification filters: better a chat too many kept from the AI than one too few.
export const aiSkipped = (title: string) => list(settings.aiSkip).some(name => title.toLowerCase().includes(name));
// An author as Teams shows them: a group's preview often has the first name only, so one word of a listed name counts.
export const skippedAuthor = (who: string) => !!who.trim() && (aiSkipped(who) || list(settings.aiSkip).some(name => name.split(/\s+/).includes(who.trim().toLowerCase())));
export const notToAi = (title: string) => { if (aiSkipped(title)) throw Object.assign(new Error(t('err.aiSkip', { title })), { status: 403 }); };
// their quote in someone else's reply goes too
export const forAi = (msgs: teams.Message[]) => msgs.filter(m => m.mine || !skippedAuthor(m.author))
  .map(m => (m.replyTo && skippedAuthor(m.replyTo) ? { ...m, replyTo: undefined, quote: undefined } : m));
export const hhmm = () => new Date().toTimeString().slice(0, 5);
// Quiet hours may cross midnight (22:00–08:00); "Do not disturb" in Teams counts too when the setting says so.
let teamsDnd = false; // read from the page on every look at the chat list
export function quiet() {
  if (settings.quietDnd && teamsDnd) return true;
  const { quietFrom: a, quietTo: b } = settings;
  if (!a || !b || a === b) return false;
  const now = hhmm();
  return a < b ? now >= a && now < b : now >= a || now < b;
}
// ponytail: "Do not disturb" (and presenting) by the words of a ru/uk/en Teams; another UI language never reads as busy
export const setDnd = (status: string) => { teamsDnd = /не беспокоить|не турбувати|do not disturb|презент|presenting/i.test(status); };

export const keywordHit = (c: teams.Chat) => list(settings.keywords).some(k => `${c.title} ${c.preview}`.toLowerCase().includes(k));
// a message for the user (personal, an @mention, a keyword hit); "Every message in group chats" brings a group's others too
export const called = (c: teams.Chat) => c.type === 'direct' || c.mention || keywordHit(c);
export const groupWide = (c: teams.Chat) => settings.groups && c.type === 'group';
export const wanted = (c: teams.Chat) => c.unread && (called(c) || groupWide(c));
// "Never from" (bots, recordings, noisy chats) wins over everything; "Only from": personal messages and @mentions from the
// listed people only; keyword hits always count. A name counts in the chat's title and as the author of its last message
// ("Author: text" in a group's preview), as a whole word: a name in a message's text, or a part of another name, does not.
export const authorOf = (c: Pick<teams.Chat, 'type' | 'preview'>) => (c.type === 'direct' ? '' : /^([^:\n]{1,80}): /.exec(c.preview)?.[1] ?? '');
export const named = (c: teams.Chat, names: string) => hasName(c.title, names) || hasName(authorOf(c), names);
export const listed = (c: teams.Chat) => !named(c, settings.muteFrom) && (!list(settings.onlyFrom).length || keywordHit(c) || named(c, settings.onlyFrom));

// a card the AI does not get: its chat, or the author of the message it answers, is on the "Not to AI" list
export const skipDraft = (d: Pick<Draft, 'title' | 'last'>) => aiSkipped(d.title) || skippedAuthor(authorOf({ type: 'group', preview: d.last }));
