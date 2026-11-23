// The mark on ▤ and on the "Replies" tab. Red, with the count: a card that moves only on the user's click (an AI session's
// draft, a card back from a reminder, a send that failed). A purple dot: a result the user asked for, ready (or failed)
// after the panel was last on screen. Plain incoming messages get no mark: Teams shows them as unread.
import type { Draft } from './base';

export type Badge = { red: number; dot: boolean };
// seenAt: when "Replies" was last on screen; Infinity while it is
export function badge(cards: Pick<Draft, 'from' | 'back' | 'failed' | 'ready' | 'busy'>[], seenAt: number): Badge {
  const red = cards.filter(d => d.from === 'ai' || d.back || d.failed).length;
  return { red, dot: !red && cards.some(d => !d.busy && (d.ready ?? 0) > seenAt) }; // red wins over the dot
}
