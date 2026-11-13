// The user's own AI agent (Claude Code, Codex, …), opened in a project folder with a request about a chat: the card's
// "Look into it in the agent". now: the agent starts on the request at once, in its cautious mode (it asks before it acts);
// else the request waits in the agent's field for the user to send.
import { app, shell } from 'electron';
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { onPath } from './ai';

export type AgentId = 'claude-vscode' | 'claude-terminal' | 'claude-desktop' | 'codex' | 'gemini' | 'copilot' | 'cursor';
export type Agent = { id: AgentId; name: string; now: boolean };

const extension = (prefix: string) => {
  try { return fs.readdirSync(path.join(os.homedir(), '.vscode', 'extensions')).some(d => d.startsWith(prefix)); } catch { return false; }
};

// The agents of this computer: a program on PATH, a VS Code extension, the handler of a link.
export function detectAgents(): Agent[] {
  const code = !!onPath(['code']);
  const all: [AgentId, string, boolean, boolean][] = [
    ['claude-vscode', 'Claude Code · VS Code', false, code && extension('anthropic.claude-code-')],
    ['claude-terminal', 'Claude Code · terminal', true, !!onPath(['claude'])],
    ['claude-desktop', 'Claude Desktop', false, !!app.getApplicationNameForProtocol('claude://')],
    ['codex', 'Codex CLI', true, !!onPath(['codex'])],
    ['gemini', 'Gemini CLI', true, !!onPath(['gemini'])],
    ['copilot', 'GitHub Copilot · VS Code', false, code && extension('github.copilot-chat-')],
    ['cursor', 'Cursor', false, !!onPath(['cursor'])],
  ];
  return all.filter(a => a[3]).map(([id, name, now]) => ({ id, name, now }));
}

// The terminal agents and their cautious mode: they start on the request at once. Claude Code plans: it only reads until the
// user approves the plan — the user's allow rules would let it act unasked in "default".
const TERMINAL: Partial<Record<AgentId, [string, string[]]>> = {
  'claude-terminal': ['claude', ['--permission-mode', 'plan']],
  codex: ['codex', ['--sandbox', 'read-only', '-a', 'on-request']],
  gemini: ['gemini', ['--approval-mode', 'default', '-i']],
};

// cmd.exe starts them: the `.cmd` launchers need it, and `start` gives a terminal agent a window of its own (Windows
// Terminal's where that is the default terminal). No argument holds a `"`: Windows paths cannot, agentPrompt swaps it out.
// ponytail: cmd expands a %NAME% of an existing variable in the request; a question naming one is rare
const q = (s: string) => `"${s}"`;
const cmd = (line: string) => new Promise<void>((ok, no) => {
  spawn(process.env.ComSpec ?? 'cmd.exe', ['/d', '/s', '/c', q(line)], { windowsVerbatimArguments: true, windowsHide: true, stdio: 'ignore' })
    .on('error', no).on('exit', () => ok());
});
const link = (base: string, params: Record<string, string>) => `${base}?${Object.entries(params).map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join('&')}`;

export async function launch(id: AgentId, dir: string, prompt: string, images: string[]) {
  const term = TERMINAL[id];
  if (term) {
    const exe = onPath([term[0]]);
    if (!exe) throw new Error(`${term[0]}: not found`);
    // Codex takes the pictures itself; the others read them from the folder chat.md names
    const args = [...term[1], ...(id === 'codex' ? [...images.flatMap(i => ['-i', i]), '-C', dir] : []), prompt];
    return cmd(`start "Soroka" /D ${q(dir)} ${[exe, ...args].map(q).join(' ')}`);
  }
  if (id === 'claude-desktop') return shell.openExternal(link('claude://code/new', { folder: dir, q: prompt }));
  if (id === 'copilot') return shell.openExternal(link('vscode://agents/new', { prompt, workspace: dir }));
  // The folder's window first; the request then goes to the editor's window in front.
  // ponytail: a fixed 2 s; wait for the window if the request lands in another one
  const editor = id === 'cursor' ? 'cursor' : 'code';
  await cmd(`${q(onPath([editor]) ?? editor)} ${q(dir)}`);
  await new Promise(r => setTimeout(r, 2000));
  await shell.openExternal(id === 'cursor' ? link('cursor://anysphere.cursor-deeplink/prompt', { text: prompt }) : link('vscode://anthropic.claude-code/open', { prompt }));
}
