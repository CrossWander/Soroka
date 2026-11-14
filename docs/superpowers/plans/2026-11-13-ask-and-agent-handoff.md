# Вопросы по переписке и передача агенту — план реализации

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** вопрос по переписке из карточки с продолжением разговора в «Підсумку» и кнопка «Розібрати в агенті», которая отдаёт переписку (текст, картинки, метки) ИИ-агенту пользователя в папке проекта.

**Architecture:** чистая логика передачи (метки, «По питанню», `chat.md`, запрос, папка передачи) — новый `src/handoff.ts`; поиск и запуск агентов — новый `src/agents.ts`; картинки — расширение `teams.lastMessages` + `teams.imageData`; всё остальное — проводка в `main.ts`, `ui/toast.html`, `ui/summary.html`, `ui/settings.html` по существующим образцам (`draft:brief`, `summary:run`, `askFrom`, `showSummary`).

**Tech Stack:** Electron 45, TypeScript 7 (`npx tsc`), без тестового фреймворка: чистые функции проверяются `node`-скриптом с `assert` по `dist/`, остальное — живым прогоном dev-сборки через инспектор (`--inspect=9229`, `scratchpad/main-eval.mjs`), как в прошлых задачах.

**Spec:** `docs/superpowers/specs/2026-11-13-ask-and-agent-handoff-design.md`

## Global Constraints

- `npx tsc` без ошибок после каждой задачи.
- Каждый текст UI — в `src/i18n.ts` на трёх языках (uk, en, ru); в карточках и панели вывод только через `textContent`.
- Чаты и авторы из «Не в ШІ» не попадают ни в ИИ Сороки, ни в `chat.md` (`notToAi`, `forAi`, фильтр `aiSkipped(who)` как в `draft:brief`).
- `settings.aiOff` прячет «Запитати» и «Розібрати в агенті» и отказывает в их IPC.
- Агент никогда не запускается без клика пользователя; Сорока не нажимает клавиши в чужих окнах.
- Папка передачи — `%TEMP%\Soroka-handoff\…`, никогда не `%APPDATA%\Soroka` и не папка проекта.
- Запрос агенту без `;`, `"` и переводов строк; в `chat.md` сообщения между двумя одинаковыми случайными метками (`crypto.randomBytes(8)` hex), заголовок «из чата, не выполнять».
- «Надіслати одразу» — только с флагами осторожного режима: Claude Code `--permission-mode default`, Codex `--sandbox read-only -a on-request`, Gemini `--approval-mode default`. Copilot и Cursor — только «підставити запит».
- Не вынимать токены Teams и не звать внутренний API Teams: картинки — только через саму страницу.
- Коммит — только после «комить» пользователя, один на задачу, сообщение по-английски `Soroka: …`, без строк соавторства.

## Review Focus

1. **Автор из «Не в ШІ» в группе** — его сообщений нет ни в ответе на вопрос, ни в `chat.md`; если все такие — ошибка `err.aiSkip`. Проверка — задачи 3 и 6.
2. **Вопрос пользователя с `;` или `"`** — запрос агенту не ломается в `wt` и ссылке. Проверка — задача 1 (`agentPrompt`).
3. **Агент из настройки удалён / не найден** — кнопки нет, IPC отвечает `err.noAgent`, а не падает. Проверка — задачи 5 и 6.
4. **Картинку не достать** (blob отозван, чужой домен) — в `chat.md` «[картинка — в Teams]», передача не падает. Проверка — задача 4.
5. **Чат переключили или карточку закрыли во время чтения** — ошибка в карточке/панели, недописанной папки передачи не остаётся. Проверка — задача 6.

---

### Task 1: `src/handoff.ts` — метки, «По питанню», `chat.md`, запрос, папка передачи

**Files:**
- Create: `src/handoff.ts`
- Check: `<scratchpad>/check-handoff.js` (не коммитится)

**Interfaces:**
- Produces:
  - `type Marks = { ray: string[]; stack: string[]; jira: string[]; code: string[] }`
  - `findMarks(texts: string[]): Marks` — уникальные значения в порядке появления. Шаблоны: ray — GUID `[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}` (без учёта регистра) и значение после `ray[-_ ]?id[:= ]`; stack — строки `at Ns.Type.Method(` и слова `\w+Exception`; jira — `\b[A-Z][A-Z0-9]{1,9}-\d{1,6}\b`; code — `\b[\w.-]+\.(cs|ts|js|proto|sql|xaml|json)\b`.
  - `markNames(m: Marks): (keyof Marks)[]` — непустые виды в порядке `ray, stack, jira, code`.
  - `byQuestion<T extends { id: string }>(msgs: T[], gapMs: number): T[]` — `id` = epoch ms; хвост от последнего сообщения назад, пока разрыв между соседними ≤ `gapMs`.
  - `type Talk = { q: string; a: string }`
  - `type ChatLine = { author: string; mine: boolean; time: string; text: string; images: string[] }` — `images`: имена файлов рядом (`img-1.png`) или `''` для недоставшейся.
  - `renderChat(o: { header: string; title: string; period: string; lines: ChatLine[]; marks: Marks; talk: Talk[]; fence: string; noImage: string }): string`
  - `agentPrompt(template: string, file: string, question?: string): string` — подставляет `{file}`, `{question}` (≤ 300 символов); убирает `;` → `,`, `"` → `'`, переводы строк → пробел.
  - `HANDOFF_ROOT = path.join(os.tmpdir(), 'Soroka-handoff')`
  - `writeHandoff(title: string, files: { name: string; data: string | Buffer }[], root = HANDOFF_ROOT): string` — папка `<yyyyMMdd-HHmmss>-<title без недопустимых символов, ≤ 40>`, возвращает её путь.
  - `cleanHandoffs(keepMs: number, root = HANDOFF_ROOT, now = Date.now()): number` — удаляет подпапки старше `keepMs` по `mtime`, возвращает число.

- [ ] **Step 1: Проверочный скрипт** `check-handoff.js` (требует `../…/dist/handoff.js` по абсолютному пути), `assert`:
  - `findMarks(['ray-id: 7f3c9a1e-1b2c-4d5e-8f90-a1b2c3d4e5f6', 'NullReferenceException at Shop.Orders.Recount(', 'см. ABC-1234 и OrderSheet.cs'])` → `ray` = `['7f3c9a1e-1b2c-4d5e-8f90-a1b2c3d4e5f6']`, `stack` содержит `NullReferenceException`, `jira` = `['ABC-1234']`, `code` = `['OrderSheet.cs']`;
  - `markNames` того же → `['ray','stack','jira','code']`; на `['привіт']` → `[]`;
  - `byQuestion` на id `[0, 1h, 5h, 5h+10m, 6h]` при `gapMs = 2h` → последние три;
  - `agentPrompt('Прочитай {file}. {question}', 'C:\\x\\chat.md', 'а; "б"\nв')` не содержит `;`, `"`, `\n`;
  - `renderChat` с `fence: 'abc'` содержит ровно две строки с `abc`, заголовок до первой метки, `noImage` для `images: ['']`;
  - `writeHandoff` во временный корень создаёт папку с файлами; `cleanHandoffs(0, root, Date.now() + 1000)` → `1`, папки нет.
- [ ] **Step 2:** `npx tsc` → ошибка импорта (модуля нет) / скрипт падает.
- [ ] **Step 3:** Реализовать `src/handoff.ts` по Interfaces.
- [ ] **Step 4:** `npx tsc` → 0 ошибок; `node <scratchpad>\check-handoff.js` → `ok`.
- [ ] **Step 5:** Коммит после «комить»: `Soroka: handoff helpers — marks, "by question", chat.md, the agent's prompt, the handoff folder`.

### Task 2: «По питанню» и разговор в «Підсумку»

**Files:**
- Modify: `src/main.ts` — `Range` (+`gap?: number`), `range()` (`'q'` → `{ n: 100, gap: 7_200_000 }`), `summaryRequest` (применяет `byQuestion`, когда `gap`), `Ask` (+`history?: Talk[]`), `askFrom` (≤ 10 пар, каждая строка ≤ 2000), `asked(question, prompt, history)` (блок «Earlier in this conversation» перед `Question:`), `showSummary(chatId?, result?, talk?: { period: string; history: Talk[] })` (в preset).
- Modify: `ui/summary.html` — опции `value="q"`, `value="n20"`, `value="n50"` в `#period` (`n100` уже есть); лента вопросов/ответов вместо одиночного ответа; `talk` сбрасывается при смене чата или периода; preset с `period`/`history` выставляет период и засевает ленту.
- Modify: `src/i18n.ts` — `summary.q`, `summary.n20`, `summary.n50`, `summary.talkHead`.

**Interfaces:**
- Consumes: `byQuestion`, `Talk` (Task 1).
- Produces: `range('q')`; `askFrom` принимает `history`; `showSummary(chatId, undefined, { period, history })` — используется в Task 3.

- [ ] **Step 1: Живой сценарий** (`<scratchpad>/talk-test.js`, dev-сборка): в self-чате панель «Підсумок», период «по питанню»; вопрос 1, затем вопрос 2 «а хто це має зробити?». Подмена `ai.ask` записывает промпты: во втором есть текст вопроса 1 и ответа 1; смена периода → третий промпт без них.
- [ ] **Step 2:** Прогон до правок → во втором промпте нет первого вопроса.
- [ ] **Step 3:** Реализовать.
- [ ] **Step 4:** `npx tsc`; прогон → ок; на экране лента из двух пар.
- [ ] **Step 5:** Коммит после «комить»: `Soroka: the summary tab talks — follow-up questions see the earlier ones, and "by question" reads back to the last long pause`.

### Task 3: «Запитати» в карточке и «Продовжити в Підсумку»

**Files:**
- Modify: `src/main.ts` — IPC `draft:ask(n, question)` → `{ text, msgs }` (как `draft:brief`: `notToAi`, `wholeMsgs`, фильтр авторов, `askAi(asked(...), 'summaryModel')`); IPC `draft:continue(n, talk: Talk)` → `showSummary(d.chatId, undefined, { period: 'q', history: [talk] })`; в `announce` payload `reply.ask = aiReady() && !aiSkipped(d.title)`.
- Modify: `ui/toast.html` — ссылка «Запитати» в `#subline` рядом с «✨ Коротко»; поле `#ask-field` + ➤ в карточке; ответ — в блоке `#sum` (заголовок `card.answerHead`) с кнопкой «Продовжити в Підсумку →»; запрос через существующий `request()`.
- Modify: `src/preload.ts` — `draft:ask`, `draft:continue`.
- Modify: `src/i18n.ts` — `card.ask`, `card.askPh`, `card.answerHead`, `card.continue`, `card.asking`.

**Interfaces:**
- Consumes: `asked` с `history` и `showSummary(..., talk)` (Task 2).

- [ ] **Step 1: Живой сценарий** (`<scratchpad>/card-ask-test.js`, как `fix-test.js`: выдуманные чаты через подмену `teams.chatList`, `ai.ask` подменён): группа из 4 сообщений, автор «олена пробна» в «Не в ШІ»; «Запитати» → в промпте нет её текстов; ответ в карточке; «Продовжити» → панель «Підсумок» с периодом `q` и одной парой в ленте; все авторы в «Не в ШІ» → ошибка `err.aiSkip` в карточке.
- [ ] **Step 2:** Прогон до правок → нет `#ask-field`.
- [ ] **Step 3:** Реализовать.
- [ ] **Step 4:** `npx tsc`; прогон → ок; тестовые черновики отклонены, «Не в ШІ» восстановлен.
- [ ] **Step 5:** Коммит после «комить»: `Soroka: a card answers a question about its messages and carries it on in the summary tab`.

### Task 4: Картинки из сообщений

**Files:**
- Modify: `src/teams.ts` — `Message.images?: string[]` (src картинок из `content-<mid>`, без эмодзи: `src` с `/emoticons/` и `img[itemtype*="Emoji"]` пропускаются); `export const imageData = (srcs: string[]): Promise<(string | null)[]>` — в странице: `blob:`/same-origin → `fetch(src)` → data URL; иначе уже загруженный `<img>` с этим `src` → canvas → `toDataURL`; не вышло → `null`. Комментарий: где селектор/путь проверен живьём.

**Interfaces:**
- Produces: `Message.images`, `teams.imageData` — используются в Task 6.

- [ ] **Step 1: Живой сценарий** (`<scratchpad>/images-test.js`): пользователь кладёт в self-чат одно сообщение со скриншотом (с его «да»); `teams.history(self, 5)` → у сообщения `images.length === 1`; `imageData(images)` → строка `data:image/…`, длина > 1000; сообщение с эмодзи → `images` пусто.
- [ ] **Step 2:** Прогон до правок → `images` undefined.
- [ ] **Step 3:** Реализовать.
- [ ] **Step 4:** `npx tsc`; прогон → ок; лог — только длины и типы, без текстов. Тестовое сообщение удалить.
- [ ] **Step 5:** Коммит после «комить»: `Soroka: messages carry their pictures, read through the Teams page itself`.

### Task 5: `src/agents.ts` и настройки агента

**Files:**
- Create: `src/agents.ts`
- Modify: `src/ai.ts` — `export` у `onPath`.
- Modify: `src/main.ts` — `Settings` + `agent: string; projectDir: string; projectAsk: boolean; handoffKeep: string`; умолчания `agent: ''`, `projectDir: ''`, `projectAsk: false`, `handoffKeep: '1'`; `OPTIONS.handoffKeep = ['1', '24', '168']`; `CHECKS`: `agent` — `''` или id из `detectAgents()`, `projectDir` — `isText(500)`, `projectAsk` — boolean; `state.lastDir: string`; IPC `agents:list` → `Agent[]`, `dir:pick` → путь из `dialog.showOpenDialog({ properties: ['openDirectory'], defaultPath })` или `undefined`; таймер `cleanHandoffs(Number(settings.handoffKeep) * 3_600_000)` при старте и раз в час.
- Modify: `ui/settings.html` — в разделе «ШІ» строки «Агент» (select из `agents:list`, пустой список → заметка `set.agentNone`), «Папка проектів» (поле + «Обрати…» через `dir:pick` + переключатель «питати щоразу»), «Зберігати файли передачі» (1 год / 1 день / тиждень) с заметкой `set.handoffNote`.
- Modify: `src/preload.ts`, `src/i18n.ts` — ключи `set.agent`, `set.agentNote`, `set.agentNone`, `set.projectDir`, `set.projectPick`, `set.projectAsk`, `set.handoffKeep`, `set.keep1`, `set.keep24`, `set.keep168`, `set.handoffNote`, имена агентов.

**Interfaces:**
- Consumes: `ai.onPath`.
- Produces:
  - `type AgentId = 'claude-vscode' | 'claude-terminal' | 'claude-desktop' | 'codex' | 'gemini' | 'copilot' | 'cursor'`
  - `type Agent = { id: AgentId; name: string; now: boolean }` — `now`: отправляет сразу.
  - `detectAgents(): Agent[]` — синхронно: `code` на PATH + `~/.vscode/extensions/anthropic.claude-code-*` (claude-vscode) / `github.copilot-chat-*` (copilot); `claude` на PATH + `wt.exe` (claude-terminal); `HKCU\Software\Classes\claude` (claude-desktop, через `reg query`, один раз на вызов); `codex`, `gemini` на PATH + `wt.exe`; `cursor` на PATH (cursor).
  - `launch(id: AgentId, dir: string, prompt: string, images: string[]): Promise<void>` — по таблице спеки; «одразу» — через `wt -w new -d <dir> --title Soroka <exe> <флаги> <prompt>` с флагами из Global Constraints, Codex + `-C <dir>` и `-i <images через запятую>`; claude-vscode — `code <dir>`, дождаться выхода, 2 с, затем `shell.openExternal('vscode://anthropic.claude-code/open?prompt=…')` (*ponytail: фиксированные 2 с; ждать окна, если ссылка уходит в другое окно*); claude-desktop — `claude://code/new?folder=…&q=…`; copilot — `vscode://agents/new?prompt=…&workspace=…`; cursor — `cursor <dir>`, 2 с, `cursor://anysphere.cursor-deeplink/prompt?text=…`. Значения — `encodeURIComponent`.

- [ ] **Step 1: Живой сценарий** (`<scratchpad>/agents-test.js`): `detectAgents()` на этой машине содержит `claude-vscode` и `claude-terminal`; `launch('claude-terminal', <scratch-папка>, 'Скажи ok', [])` открывает окно терминала (проверить глазами пользователя или по процессу `WindowsTerminal`), `launch('claude-vscode', …)` открывает VS Code с подставленным запросом — пользователь подтверждает. Настройки: выбор агента и папки сохраняется в `settings.json` (`settings:get`).
- [ ] **Step 2:** Прогон до правок → модуля нет.
- [ ] **Step 3:** Реализовать.
- [ ] **Step 4:** `npx tsc`; прогон → ок; открытые тестовые окна пользователь закрывает сам.
- [ ] **Step 5:** Коммит после «комить»: `Soroka: finds the user's AI agents and opens one in a project folder with a request`.

### Task 6: «Розібрати в агенті» — от карточки и панели до запуска

**Files:**
- Modify: `src/main.ts`:
  - `async function toAgent(src: { chatId: string; title: string; period: string; question?: string; history?: Talk[] } | { title: string; msgs: Msg[] }): Promise<string>` — возвращает имя агента. Порядок: `aiOff` → ошибка; агент (`settings.agent` или первый из `detectAgents()`; нет → `err.noAgent`); папка (`projectAsk` или папки нет на диске → `dir:pick`-диалог от `state.lastDir`; отмена → `err.noDir`); `notToAi(title)`; для чата — внутри одного `inChat`: `teams.history` по `range(period)` (+ `byQuestion`), затем `teams.imageData` всех картинок; `forAi`; `findMarks`; `renderChat` с `fence = randomBytes(8)`; `writeHandoff` (картинки `img-<k>.<ext>`); `agentPrompt(t('agent.prompt' | 'agent.promptAsk'), file, question)`; `launch`; `state.lastDir = dir`. Ошибка до `launch` — папка передачи удаляется.
  - IPC `draft:agent(n, period)` (период по умолчанию `'q'`), `summary:agent(chatId, title, period, ask)` (вопрос — последний из `ask.history`/`ask.question`).
  - Карточка канала: в `checkActivity` действие `↗ ${t('card.agent')}` → `toAgent({ title: a.where, msgs: card.msgs })`, только когда агент найден.
  - `announce`: `reply.agent = !settings.aiOff && detectAgents().length > 0`, `reply.marks = markNames(findMarks(msgs.map(m => m.text)))`.
  - `Summary` + `marks?: string[]` для одного чата (из прочитанных сообщений).
- Modify: `ui/toast.html` — кнопка `#agent` «↗ Агент» в `#row` перед ⋯ (класс `hot` и `title` со списком меток, когда `marks.length`), рядом ▾ с меню периодов (`q`, `d0`, `d7`, `n50`); запрос через `request()`; после — подсказка `card.agentDone`.
- Modify: `ui/summary.html` — кнопка «↗ Розібрати в агенті» рядом с «Відкрити в чаті ШІ» (передаёт текущие чат, период, ленту), выделение по `r.marks`.
- Modify: `src/preload.ts`, `src/i18n.ts` — `card.agent`, `card.agentTip`, `card.agentDone`, `mark.ray`, `mark.stack`, `mark.jira`, `mark.code`, `summary.agent`, `agent.header`, `agent.prompt`, `agent.promptAsk`, `agent.noImage`, `err.noAgent`, `err.noDir`, `busy.handoff`.

**Interfaces:**
- Consumes: всё из Tasks 1–5.

- [ ] **Step 1: Живой сценарий** (`<scratchpad>/agent-flow-test.js`, `agents.launch` подменён — записывает аргументы, ничего не открывает; `ai.ask` не нужен):
  - карточка self-чата с сообщением `ray-id 7f3c9a1e-1b2c-4d5e-8f90-a1b2c3d4e5f6` → кнопка `#agent` с классом `hot`;
  - клик → `launch` вызван с папкой из настроек и запросом без `;`/`"`; в папке передачи `chat.md` с заголовком, двумя метками и ray-id в списке; картинка из Task 4 — файлом рядом;
  - автор из «Не в ШІ» — его текста в `chat.md` нет;
  - `settings.agent` указывает на отсутствующего агента и других нет → `err.noAgent`, папки передачи нет;
  - панель «Підсумок» → «Розібрати в агенті» с лентой из Task 2 → вопросы и ответы в `chat.md`;
  - `cleanHandoffs` с `keepMs = 0` удаляет созданные папки.
- [ ] **Step 2:** Прогон до правок → нет `#agent`.
- [ ] **Step 3:** Реализовать.
- [ ] **Step 4:** `npx tsc`; прогон → ок; затем один настоящий прогон без подмены: «Розібрати в агенті» из карточки self-чата открывает VS Code с подставленным запросом (подтверждает пользователь); тестовые сообщения и черновики удалить.
- [ ] **Step 5:** Коммит после «комить»: `Soroka: "Hand to agent" — the chat, its pictures and found ray-ids go to the user's AI agent in the project folder`.

### Task 7: Сборка и установка

- [ ] **Step 1:** Корректно закрыть dev (`flushStore` + `app.quit()`), `npm run dist`, в `release/` только новый установщик, установить `/S`, запустить Сороку — процессы есть.
- [ ] **Step 2:** Обновить заметку в памяти (`project_soroka_ai_providers.md`): что сделано, коммиты, что не проверено живьём.
