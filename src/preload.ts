// Bridge for the app's own pages (answers, settings, summary, toast, the strip above Teams). The Teams page gets no preload.
import { contextBridge, ipcRenderer } from 'electron';

const CHANNELS = ['i18n:get', 'settings:get', 'settings:set', 'settings:close', 'status:get', 'ai:login', 'ai:install', 'teams:show', 'teams:check', 'report:copy', 'notify:test', 'style:learn',
  'drafts:get', 'draft:send', 'draft:suggest', 'draft:open', 'draft:reject', 'draft:snooze', 'panel:tab', 'panel:close','react:open', 'react:get', 'react:pick', 'ring:click', 'ring:dismiss', 'ring:size',
  'chats:list', 'summary:run', 'summary:stop', 'summary:last', 'summary:handoff', 'draft:handoff', 'models:get', 'models:check', 'open:jira', 'jira:issue', 'jira:login', 'jira:token', 'jira:tokenSet', 'share:sources', 'share:pick', 'toast:size', 'toast:hidden', 'toast:click', 'toast:focus', 'draft:keep', 'strip:get', 'strip:act', 'menu:open', 'menu:size', 'menu:act', 'mini:size', 'mini:act', 'signin:size', 'signin:act', 'draft:whole', 'draft:brief', 'draft:ask', 'draft:continue', 'agents:list', 'dir:pick'];
const EVENTS = ['drafts', 'toast', 'toast:variants', 'toast:msgs', 'summary:preset', 'strip', 'ring', 'mini', 'signin'];
contextBridge.exposeInMainWorld('bridge', {
  invoke: (channel: string, ...args: unknown[]) => {
    if (!CHANNELS.includes(channel)) throw new Error(`unknown channel ${channel}`);
    return ipcRenderer.invoke(channel, ...args);
  },
  on: (event: string, fn: (payload: unknown) => void) => {
    if (!EVENTS.includes(event)) throw new Error(`unknown event ${event}`);
    ipcRenderer.on(event, (_e, payload) => fn(payload));
  },
});
