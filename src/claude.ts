// Claude Code found on this computer: no API keys, every call counts against the plan of whoever signed in to Claude Code.
import { execFile } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { t } from './i18n';

export type ClaudeInfo = { path?: string; version?: string; loggedIn?: boolean; expiresAt?: string; authMethod?: string; plan?: string; error?: string };

function candidates() {
  const home = os.homedir();
  const names = process.platform === 'win32' ? ['claude.exe', 'claude.cmd'] : ['claude'];
  const dirs = [
    path.join(home, '.local', 'bin'), // native installer
    ...(process.env.PATH ?? '').split(path.delimiter),
    ...(process.platform === 'win32' ? [path.join(process.env.APPDATA ?? '', 'npm')] : ['/opt/homebrew/bin', '/usr/local/bin', path.join(home, '.claude', 'local')]),
    ...desktop(),
  ];
  return dirs.filter(Boolean).flatMap(d => names.map(n => path.join(d, n)));
}
// Claude Desktop brings its own Claude Code, signed in with the same account: <app data>/Claude/claude-code/<version>/, newest first
function desktop() {
  const dir = process.platform === 'win32' ? path.join(process.env.APPDATA ?? '', 'Claude', 'claude-code') : path.join(os.homedir(), 'Library', 'Application Support', 'Claude', 'claude-code');
  try { return fs.readdirSync(dir).sort((a, b) => b.localeCompare(a, undefined, { numeric: true })).map(v => path.join(dir, v)); } catch { return []; }
}
export const find = () =>candidates().find(p => { try { return fs.statSync(p).isFile(); } catch { return false; } });

// Never the API: with one of these set, Claude Code would bill an API key or a cloud account instead of the signed-in plan.
const API_VARS = ['ANTHROPIC_API_KEY', 'ANTHROPIC_AUTH_TOKEN', 'CLAUDE_CODE_USE_BEDROCK', 'CLAUDE_CODE_USE_VERTEX'];
const env = Object.fromEntries(Object.entries(process.env).filter(([k]) => !API_VARS.includes(k.toUpperCase())));

// npm installs a .cmd wrapper, which Node can start only through cmd.exe: then every argument needs quotes.
function exec(exe: string, args: string[], input: string | undefined, timeout: number, signal?: AbortSignal) {
  const shell = exe.endsWith('.cmd');
  const argv = shell ? args.map(a => `"${a}"`) : args;
  return new Promise<string>((ok, no) => {
    const p = execFile(shell ? `"${exe}"` : exe, argv, { timeout, maxBuffer: 1 << 21, windowsHide: true, shell, env, signal },
      (err, out, errOut) => (err ? no(Object.assign(err, { output: `${out}\n${errOut}`, stdout: out })) : ok(out.trim())));
    p.stdin?.end(input ?? '');
  });
}

export async function info(): Promise<ClaudeInfo> {
  const exe = find();
  if (!exe) return {};
  try {
    const version = (await exec(exe, ['--version'], undefined, 20_000)).split(' ')[0];
    const auth = JSON.parse(await exec(exe, ['auth', 'status'], undefined, 20_000));
    return { path: exe, version, loggedIn: !!auth.loggedIn, expiresAt: auth.loginExpiresAt, authMethod: auth.authMethod, plan: auth.subscriptionType };
  } catch (e) { return { path: exe, error: (e as Error).message.split('\n')[0] }; }
}

// Tokens each answer took (input, cache and output), as Claude Code reports them: the app shows the day's total.
export const usage = { onSpent: (_tokens: number) => {} };
type Result = { result?: string; is_error?: boolean; usage?: Record<string, number> };

// One answer, no tools, no MCP, no project settings: only the text the app passes in.
export async function ask(system: string, prompt: string, model: string, dataDir: string, signal?: AbortSignal) {
  const exe = find();
  if (!exe) throw Object.assign(new Error(t('err.claudeMissing')), { claude: 'missing' });
  const systemFile = path.join(dataDir, `prompt-${Date.now()}-${Math.random().toString(36).slice(2)}.txt`); // calls may overlap
  fs.writeFileSync(systemFile, system);
  try {
    const out = await exec(exe, ['-p', '--model', model, '--tools', '', '--system-prompt-file', systemFile, '--no-session-persistence',
      '--strict-mcp-config', '--setting-sources', '', '--output-format', 'json'], prompt, 180_000, signal)
      .catch(e => { const json = (e as { stdout?: string }).stdout?.trim(); if (json && /^[[{]/.test(json) && !signal?.aborted) return json; throw e; }); // an error answer exits with 1
    const r: Result = [JSON.parse(out)].flat().at(-1) ?? {}; // a list of events in current versions, one object in older ones
    const u = r.usage ?? {};
    usage.onSpent((u.input_tokens ?? 0) + (u.output_tokens ?? 0) + (u.cache_read_input_tokens ?? 0) + (u.cache_creation_input_tokens ?? 0));
    if (r.is_error) throw Object.assign(new Error(`Claude Code: ${r.result}`), { output: r.result });
    return String(r.result ?? '').trim();
  } catch (e) {
    if (signal?.aborted) throw new Error(t('err.stopped'));
    throw failure(e);
  } finally { fs.rmSync(systemFile, { force: true }); }
}
// A failed run in a few words: sign-in, or the last line Claude Code printed
function failure(e: unknown) {
  const out = (e as { output?: string }).output ?? '';
  if (/log ?in|logged out|authenticat|\/login|expired/i.test(out)) return Object.assign(new Error(t('err.claudeLogin')), { claude: 'login' });
  return new Error(`Claude Code: ${out.trim().split('\n').pop() || (e as Error).message}`);
}

// The model an alias ("opus") stands for today and whether the plan includes it: the CLI names the model in its
// first JSON event and reports "requires usage credits" and the like in the last one.
export async function probe(alias: string): Promise<{ id: string; error?: string }> {
  const exe = find();
  if (!exe) throw Object.assign(new Error(t('err.claudeMissing')), { claude: 'missing' });
  const out = await exec(exe, ['-p', '--model', alias, '--tools', '', '--system-prompt', 'Reply with one word.', '--no-session-persistence',
    '--strict-mcp-config', '--setting-sources', '', '--output-format', 'json'], 'OK', 60_000)
    .catch(e => { const json = (e as { stdout?: string }).stdout?.trim(); if (json && /^[[{]/.test(json)) return json; throw failure(e); }); // a model the plan lacks exits with 1
  const events = [JSON.parse(out)].flat();
  const result = events.at(-1);
  return { id: events.find(e => e.type === 'system')?.model || alias, error: result?.is_error ? String(result.result) : undefined };
}
