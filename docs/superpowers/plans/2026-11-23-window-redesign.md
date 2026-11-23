# Окно Сороки: одна полоса, кнопки у поля, честный счётчик — план работ

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Убрать «Сорока ▾» и вторую серую полосу: действия — кнопками над полем ввода Teams, ▤ и ⚙ — в шапке Teams, а красный счётчик загорается только когда без клика пользователя ничего не сдвинется.

**Architecture:** Один внедрённый в страницу Teams наблюдатель (`teams.watchUi`) держит в чужом DOM ряд кнопок над полем ввода и группу ▤ ⚙ в шапке; клики приходят в приложение через DevTools-привязку `tbAct`, найденная/потерянная шапка — через `tbBar`. Приложение отдаёт странице состояние одним `teams.showUi(PageUi)` при каждом изменении (как `pushStrip`). Шапка найдена — страница Teams, панель и настройки начинаются с y = 0 и `titleBarOverlay` берёт высоту и цвет шапки; нет — прежняя полоса `ui/strip.html` 32 px. Счётчик — чистая функция `badge()` по карточкам и одной отметке «панель показана».

**Tech Stack:** Electron 45, TypeScript (`npm run build` = `tsc`), без тестового фреймворка: чистая логика — scratch-скрипт с `node:assert`, остальное — живой прогон на профиле пользователя.

**Spec:** `docs/superpowers/specs/2026-11-23-window-redesign-design.md` (наброски: https://claude.ai/artifact/6k4NeioUNcrMdSboBPC6Bm, доски A, «Верх 3», «Лічильник»).

## Global Constraints

- Репо `C:\Users\CrossWander\Desktop\GitHub\Soroka`, ветка `master`. Коммит — только после того, как пользователь посмотрел дифф и сказал; без `Co-Authored-By`; пушит и ставит теги пользователь.
- В рабочем дереве лежит незакоммиченная работа 0.33.3 (`src/icon.ts`, `build/`, `README*.md`, `docs/icon.png`, `package*.json`): в коммиты этого плана не попадает — `git add` только файлов задачи.
- Разметка в странице Teams — только `createElement` (Trusted Types); страница → приложение — только `bind()` в `src/teams.ts`, вызовы принимаются лишь со страницы `TEAMS_PAGE`.
- Стили в страницу Teams — только `insertCSS(..., { cssOrigin: 'user' })` (CSP страницы режет inline-`<style>`); правило `*{app-region:no-drag!important}` из `dom-ready` остаётся.
- Каждая новая строка интерфейса — на uk / en / ru в `src/i18n.ts`.
- Селекторы шапки Teams и контейнера поля ввода находятся вживую; рядом комментарий «seen live 2026-11-xx» (как у `RING`/`TOAST`).
- Кнопки у поля только запускают `ACTIONS.*`; ничего не отправляют.
- Живой прогон: установленную Сороку пользователь закрывает через трей (Сороку не убивать); dev — `scratchpad\electron-dev\electron.exe <repo> --remote-debugging-port=9335` (копия с выключенным предохранителем шифрования cookie), в логе `sign-in cookies kept: 22`; выход — `node scratchpad\browser-close.mjs 9335`; в конце — запуск установленной Сороки.
- Не читать и не коммитить `%USERPROFILE%\.soroka-signing\update-key.pem` и `%APPDATA%\Soroka`.
- Комментарии в коде — по-английски, в стиле файла.

## Review Focus

1. **Шапки Teams нет** (страница входа Microsoft, загрузка, ошибка, Teams переделал шапку) — окно всё равно таскается, разворачивается и закрывается: показана полоса 32 px. Проверка — Task 3, шаг 8.
2. **Масштаб экрана 125/150 %** — кнопки окна Windows не закрывают «···» и аватар Teams: отступ шапки берётся из измеренной ширины overlay, не из 138. Проверка — Task 3, шаг 8 (лог `controls:` + клик по аватару).
3. **Teams перерисовал поле** (смена чата, ответ в ветке, правка своего сообщения) — ровно один ряд кнопок, не внутри редактируемого текста. Проверка — Task 2, шаг 6.
4. **Клик по кнопке у поля** не уводит каретку из поля, не отправляет форму Teams и не теряет набранный текст (`type="button"`, `mousedown` → `preventDefault`). Проверка — Task 2, шаг 6.
5. **Красное «залипает»** после того, как карточку отправили, отклонили или ответили в самом Teams; точка не возвращается после открытия панели. Проверка — Task 1, шаги 1 и 7.

---

### Task 1: Счётчик

**Files:**
- Create: `src/badge.ts`
- Modify: `src/base.ts:89-91` (тип `Draft`), `src/api.ts:97`, `src/main.ts` (`stripState` 153, `openPanel`/`closePanel` 546-573, `suggest` 617-640, `snooze` ~726, `remind` 743-746, `sendDraft` 980, `handle('drafts:get')` 1004), `src/preload.ts` (CHANNELS), `ui/strip.html` (`show()`), `ui/panel.js` (`count`)
- Test: `scratchpad\badge-check.js` (не коммитится)

**Interfaces:**
- Produces: в `Draft` — `from?: 'ai'` (черновик сессии ИИ), `back?: boolean` (вернулась по напоминанию), `failed?: boolean` (отправка не удалась), `ready?: number` (когда готов результат, о котором просил пользователь, успех или ошибка).
- Produces: `src/badge.ts` — `export type Badge = { red: number; dot: boolean }`; `export function badge(cards: Pick<Draft, 'from' | 'back' | 'failed' | 'ready' | 'busy'>[], seenAt: number): Badge` — `red` = число карточек с `from === 'ai' || back || failed`; `dot` = `red === 0` и есть карточка `!busy && ready > seenAt`. Импорт `Draft` — только `import type` (файл должен грузиться в голом node).
- Produces: в `main.ts` — `badgeNow(): Badge` = `badge(pending(), panelTab === 'drafts' ? Infinity : panelSeenAt)`; `stripState().badge` вместо `.pending`; IPC `drafts:badge` → `Badge`.

- [ ] **Step 1: Scratch-проверка**

`scratchpad\badge-check.js`: `const { badge } = require('C:/Users/CrossWander/Desktop/GitHub/Soroka/dist/badge.js')` и `assert.deepStrictEqual`:
- `badge([{}], 0)` → `{ red: 0, dot: false }` (обычная входящая)
- `badge([{ from: 'ai' }], 0)`, `badge([{ back: true }], 0)`, `badge([{ failed: true }], 0)` → `{ red: 1, dot: false }`
- `badge([{ ready: 100 }], 50)` → `{ red: 0, dot: true }`; `badge([{ ready: 100 }], 150)` → `dot: false`; `badge([{ ready: 100, busy: true }], 50)` → `dot: false`
- `badge([{ ready: 100 }], Infinity)` → `dot: false` (панель «Відповіді» открыта)
- `badge([{ from: 'ai' }, { ready: 100 }], 50)` → `{ red: 1, dot: false }` (красное побеждает)

- [ ] **Step 2:** `node scratchpad\badge-check.js` → падает: `Cannot find module …/dist/badge.js`.

- [ ] **Step 3:** Поля в `Draft` и `src/badge.ts` по Interfaces; `npm run build`.

- [ ] **Step 4:** `node scratchpad\badge-check.js` → без вывода, код 0.

- [ ] **Step 5: Признаки на карточках** (`main.ts`, `api.ts`)
- `api.ts:97` (`POST /draft`): `from: 'ai'` в `addDraft`.
- `remind()`: в ветке конца snooze и в ветке «досі без відповіді» — `d.back = true` рядом с `d.reminded = true` (реакция и ИИ-сортировка `reminded` ставят как раньше, `back` — нет).
- `snooze()`: `delete d.back`.
- `sendDraft` catch: `d.failed = true`; `suggest()` при старте рядом с `delete d.error` — `delete d.failed` (ошибка отправки с карточки ушла — ушло и красное).
- `suggest()` в `finally`: `if (!auto) d.ready = Date.now()` — все вызовы без `auto` идут от пользователя (`suggestThenShow` из ✨/Покращити/Нагадати, `draft:suggest` с карточки и панели); режим «drafts» зовёт с `auto: true`.

- [ ] **Step 6: Отметка «панель показана»** — `let panelSeenAt = Date.now()` (в памяти); ставится `Date.now()`, когда «Відповіді» уходят с экрана: `closePanel()` при `panelTab === 'drafts'` и `openPanel(tab)` при переходе с `'drafts'` на другую вкладку. `badgeNow()`, `handle('drafts:badge', badgeNow)`, `'drafts:badge'` в CHANNELS `preload.ts`. Трей (`refreshTray`) не трогать.

- [ ] **Step 7: Отметка в UI**
- `ui/strip.html` `show(s)`: `s.badge.red` → красный `.badge` как сейчас (`9+` после 9); иначе `s.badge.dot` → кружок 8 px цвета `var(--dot)` в том же углу ▤; иначе ничего.
- `ui/panel.js` `count`: `bridge.invoke('drafts:badge')` вместо длины `drafts:get`; то же правило, точка — `var(--accent)`.
- `npm run build`; живой прогон: при старте на ▤ нет красного от входящих; `node tb.mjs draft "<свой чат>" "тест"` → красное «1»; «Відхилити» в панели → пропало; ✨ на карточке в панели, панель закрыть до готовности → фиолетовая точка, открыть панель → точка погасла.

- [ ] **Step 8: Коммит** — после просмотра диффа пользователем:
```bash
git add src/badge.ts src/base.ts src/api.ts src/main.ts src/preload.ts ui/strip.html ui/panel.js
git commit -m "Soroka: the red count on ▤ only for what waits on the user's click — an AI session's draft, a card back from a reminder, a failed send; a purple dot for a result they asked for and have not seen"
```

---

### Task 2: Кнопки над полем ввода

**Files:**
- Modify: `src/teams.ts` (новые `UiAct`, `Chips`, `PageUi`, `UI_CSS`, `watchUi`, `showUi` рядом с `watchRings`; `selfCheck` 56-74), `src/main.ts` (`ACTIONS` 160-168, `watch()` 314, `dom-ready` 307-310, `pushStrip` 155), `src/i18n.ts`

**Interfaces:**
- Consumes: `Badge` (Task 1) — пока не нужен.
- Produces (`teams.ts`):
  - `export type UiAct = 'suggest' | 'improve' | 'summary' | 'drafts' | 'settings'`
  - `export type Chips = { suggest: string; improve: string; summary: string; tips: Record<'suggest' | 'improve' | 'summary', string>; empty: string; ai: boolean; running: UiAct[] }`
  - `export type PageUi = { chips: Chips }` (Task 3 добавит `bar`)
  - `export async function watchUi(icon: string, onAct: (act: UiAct) => void): Promise<void>` — привязка `tbAct`, принимает только значения `UiAct`
  - `export const showUi = (ui: PageUi): Promise<void>` — `run()`: кладёт в `window.__tbUi` и вызывает `window.__tbRender?.()`
  - `export const UI_CSS: string`
- Produces (`main.ts`): `const running = new Set<teams.UiAct>()`; `pageUi(): teams.PageUi`; `pushPage()` — зовёт `teams.showUi` только если JSON изменился (сброс при каждом `watch()`), ошибки `run()` глотает.

- [ ] **Step 1: Строки** (`src/i18n.ts`): `strip.summary` → `['📝 Підсумок чату', '📝 Chat summary', '📝 Выжимка чата']`; новые `strip.improveEmpty` → `['Спершу напишіть текст', 'Type your text first', 'Сначала напишите текст']`, `diag.chips` → `['кнопки над полем вводу', 'buttons above the compose box', 'кнопки над полем ввода']`. Подписи и подсказки кнопок — существующие `strip.suggest`, `strip.improve`, `strip.summary`, `strip.*Tip` (в них уже горячие клавиши).

- [ ] **Step 2: Найти вживую** (dev-прогон, CDP-eval в странице Teams через порт 9335): контейнер поля ввода `COMPOSE` — предок `BOX`, внутри которого поле и панель значков Teams, ряд встаёт его предыдущим соседом; признак «чат, а не канал» — `[id^="chat-header-"]` (проверить, что в канале его нет); что `var(--colorBrandForeground1)` и соседние токены Fluent разрешаются у места вставки. Записать в комментарий.

- [ ] **Step 3: `watchUi` — часть «поле»** (`teams.ts`). Страничная функция ставит `MutationObserver` на `document.body` с задержкой 300 мс (как `watchChats`) и `window.__tbRender`:
  - Есть `[id^="chat-header-"]` и `COMPOSE` → ровно один `div[data-tb-chips]` как его `previousElementSibling`; любой другой `[data-tb-chips]` удаляется. Нет — рядов нет.
  - Ряд: `img` значка 16 px (`icon` из аргумента), затем кнопки `type="button"` с `data-tb-act`; `mousedown` → `preventDefault()` (каретка остаётся в поле); `click` → `window.tbAct(act)`, если кнопка не `tb-run`.
  - `ai: false` (ИИ выключен или не готов) → только «Підсумок чату» (решение пользователя).
  - «Покращити» серая с подсказкой `empty`, пока текст поля без цитаты пуст (как `boxText`); пересчёт на `input` в поле и в наблюдателе.
  - Кнопка из `running` — класс `tb-run`: крутилка вместо эмодзи, клик игнорируется.
- `UI_CSS`: таблетки 28 px высотой, скругление 14 px, 12 px шрифт, зазор 6 px, ряд с отступом 4 px снизу; цвета — токены Fluent с запасными значениями: текст `var(--colorBrandForeground1, #5b5fc7)`, рамка `var(--colorBrandStroke2, #c5cbfa)`, фон прозрачный, hover `var(--colorBrandBackground2, #ebf3fc)`; серые — `opacity: .45`; крутилка — `@keyframes`, с `prefers-reduced-motion` медленнее (как в `strip.html`).

- [ ] **Step 4: Приложение** (`main.ts`)
  - `ACTIONS.suggest/improve/summary` → через `once(name, fn)`: имя уже в `running` — ничего; иначе добавить, `pushPage()`, `fn()` с `.catch(e => log(...))`, в `finally` убрать и `pushPage()`. Горячие клавиши идут тем же путём.
  - `pageUi().chips`: подписи и подсказки из `t()`; `ai = !settings.aiOff && aiReady()`.
  - `watch()`: `teams.watchUi(APP_ICON_URL, act => ACTIONS[act]())`, затем сброс JSON и `pushPage()`; `pushStrip()` в конце зовёт `pushPage()`.
  - `dom-ready` (для `main()`): `page.insertCSS(teams.UI_CSS, { cssOrigin: 'user' })`.
  - `selfCheck`: `chips: chatOpen && has('[id^="chat-header-"]') ? has('[data-tb-chips]') : null`.

- [ ] **Step 5:** `npm run build` → без ошибок.

- [ ] **Step 6: Живой прогон** (светлая и тёмная тема Teams):
  - ряд есть в личном и групповом чате, нет в канале и в «Активності»; после 5 переключений чатов и правки своего сообщения `document.querySelectorAll('[data-tb-chips]').length === 1`;
  - набрать текст → клик по каждой кнопке: текст в поле цел, каретка в поле, ничего не отправлено; клик = горячая клавиша; повторный клик во время крутилки ничего не запускает;
  - пустое поле → «Покращити» серая с подсказкой; `aiOff` (или ИИ не готов) → только «Підсумок чату»;
  - лог `teams check:` содержит `chips ok`.

- [ ] **Step 7: Коммит** — после просмотра:
```bash
git add src/teams.ts src/main.ts src/i18n.ts
git commit -m "Soroka: «Запропонувати відповідь», «Покращити текст» and «Підсумок чату» as buttons above Teams' compose box in a chat, with a spinner while they run"
```

---

### Task 3: Одна полоса

**Files:**
- Create: `ui/frame.js`
- Delete: `ui/menu.html`
- Modify: `src/teams.ts` (`watchUi`, `PageUi`, `UI_CSS`, `selfCheck`), `src/main.ts` (`refreshTray` 129-141, 145-215 полоса и меню, `createTeamsWindow` 251-264, `layout` 275-280, `newTeamsPage` 287-310, `showSettings`/`closeSettings` 585-609), `src/preload.ts`, `src/i18n.ts`, `ui/strip.html`, `ui/panel.js`, `ui/drafts.html`, `ui/summary.html`, `ui/settings.html`

**Interfaces:**
- Consumes: `Badge`, `badgeNow()` (Task 1); `watchUi`, `showUi`, `PageUi`, `UI_CSS`, `pushPage()` (Task 2).
- Produces (`teams.ts`):
  - `export type Bar = { height: number; bg: string; fg: string }` — цвета `#rrggbb`
  - `export type BarUi = { badge: Badge; busy: string; panel: boolean; settings: boolean; controls: number; tips: { drafts: string; settings: string } }`
  - `PageUi = { chips: Chips; bar: BarUi }`
  - `watchUi(icon: string, onAct: (act: UiAct) => void, onBar: (bar: Bar | null) => void)` — вторая привязка `tbBar` (JSON `Bar` или `''`)
- Produces (`main.ts`): `let bar: teams.Bar | null`; `let controls = 138` (ширина кнопок окна, DIP); `top() = bar ? 0 : STRIP`; `frame() = { top: bar?.height ?? 0, right: bar ? controls : 0 }`; IPC `strip:controls(width)`, `frame:get`; событие `frame` в страницы панели и настроек.
- Produces (`ui/frame.js`): ставит на `:root` `--frame-top`/`--frame-right` в px и класс `framed` при `top > 0`; по `frame:get` при загрузке и по событию `frame`.

- [ ] **Step 1: Найти вживую** шапку Teams `HEADER` — общий предок поиска, `[data-tid="more-options-header"]` и аватара (`#idna-me-control-avatar-trigger`) во всю ширину страницы; обёртку «···», перед которой встаёт группа; где у шапки фон (первый непрозрачный `backgroundColor` вверх по предкам). Комментарий с датой.

- [ ] **Step 2: `watchUi` — часть «шапка»** (`teams.ts`, тот же наблюдатель):
  - `HEADER` найден → `data-tb-bar` на нём; `img[data-tb-logo]` 18 px первым ребёнком (не кнопка); `span[data-tb-group]` перед обёрткой «···»: ▤ (`aria-pressed` = `panel`; красное число / точка по `badge`; `busy` → маленькая крутилка и `title = busy`, иначе `tips.drafts`) и ⚙ (`aria-pressed` = `settings`) → `tbAct('drafts' | 'settings')`; `document.documentElement.style.setProperty('--tb-controls', controls + 'px')`.
  - Сообщение в `tbBar`: `{ height, bg, fg }` — найденная шапка, стабильная 300 мс; `''` — шапки нет 1000 мс подряд; только при изменении JSON (смена темы Teams меняет `bg`/`fg` → новое сообщение).
  - `UI_CSS` += (суть — задаётся точно):
```css
[data-tb-bar], [data-tb-bar] * { app-region: drag !important; }
[data-tb-bar] :is(button, a, input, textarea, select, [role="button"], [role="search"], [role="combobox"], [role="menuitem"], [role="tab"], [contenteditable="true"]),
[data-tb-bar] :is(button, a, input, textarea, select, [role="button"], [role="search"], [role="combobox"], [role="menuitem"], [role="tab"], [contenteditable="true"]) * { app-region: no-drag !important; }
[data-tb-bar] { padding-right: calc(var(--tb-controls, 0px) + 8px) !important; box-sizing: border-box !important; }
```
  плюс вид группы: фон `var(--colorBrandBackground2, #e2e3f7)`, кнопки 32×28, скругление 4 px (доска «Верх 3»).
  - `selfCheck`: `bar: has('[data-tb-group]')`.

- [ ] **Step 3: Режимы окна** (`main.ts`)
  - `onBar(b)`: то же JSON — выход; иначе `bar = b`, лог `teams bar: 48px` / `teams bar: none`, `layout()`, `pushStrip()`, `frame` в страницы.
  - `layout()`: всё от `top()` вместо `STRIP` (`teamsView`, панель, `settingsView`) и в конце `teamsWin.setTitleBarOverlay(overlay())`; `showSettings`/`closeSettings` зовут `layout()` вместо своих `setBounds`.
  - `overlay()`: `height = bar?.height ?? STRIP`; цвета — `bar.bg`/`bar.fg`, если `bar && !panelTab && !settingsView`, иначе нынешние цвета полосы по `nativeTheme` (их же рисует полоса у панели и настроек).
  - `pageUi().bar`: `badge: badgeNow()`, `busy: activity.at(-1) ?? ''`, `panel`, `settings`, `controls: panelTab ? 0 : controls` (у панели кнопки окна над ней), `tips` из `strip.draftsTip`/`strip.settingsTip`.
  - `newTeamsPage`: `did-start-navigation` главного фрейма, не same-document, для `main()` → `onBar(null)` (загрузка, F5, вход — полоса 32 px до новой шапки); `watchUi` получает `onBar`.
  - `strip:controls`: `controls = Math.round(width) || 138`, при изменении лог `controls: <n>`, `pushPage()`, `frame` в страницы. `frame:get` → `frame()`. В `preload.ts`: CHANNELS + `strip:controls`, `frame:get`; EVENTS + `frame`.

- [ ] **Step 4: Запасная полоса** (`ui/strip.html`): `#menu` → `<span id="name" data-i18n="strip.menu">` (значение «Сорока»), обработчик убрать; при загрузке и на `navigator.windowControlsOverlay` `geometrychange` → `bridge.invoke('strip:controls', innerWidth - (r.x + r.width))`, где `r = navigator.windowControlsOverlay.getTitlebarAreaRect()`; нет API — не звать (останется 138).

- [ ] **Step 5: Убрать «Сорока ▾»**: `ui/menu.html`; в `main.ts` — `menuView`, `menuX`, `menuClosedAt`, `keyOf`, `menuItems`, `closeMenu` (и вызов в `hide`), обработчики `menu:open/size/act`, `stripState().menu`; `chatInView` и обработчик `page-title-updated`, если больше никто не читает; в `preload.ts` — `menu:*`; в `i18n.ts` — `strip.chat`, `strip.refresh`, `strip.refreshTip`. Трей: после «Налаштування…» пункт `t('tray.reload')` → `ACTIONS.refresh()`; `tray.reload` = `['Перезавантажити Teams', 'Reload Teams', 'Перезагрузить Teams']`; `diag.bar` = `['кнопки Сороки в шапці Teams', "Soroka's buttons in Teams' header", 'кнопки Сороки в шапке Teams']`.

- [ ] **Step 6: Верх панели и настроек**: `ui/frame.js` по Interfaces, подключить в `drafts.html`, `summary.html`, `settings.html` после `i18n.js`.
  - `ui/panel.js`: `.ptabs` — `min-height: var(--frame-top, 0px)`, `app-region: drag`, `padding-right: calc(var(--frame-right, 0px) + 6px)`; `.ptab, .px { app-region: no-drag }`; `.framed .ptabs` — фон и цвет полосы (`#ebebeb`/`#242424`, тёмная `#1f1f1f`/`#ffffff`).
  - `ui/settings.html`: при `.framed` — полоса перетаскивания высотой `--frame-top` во всю ширину тех же цветов, содержимое ниже неё; ✕ раздела — под полосой.

- [ ] **Step 7:** `npm run build` → без ошибок; `git grep -n "menu:open\|menuItems\|strip.refresh"` → пусто.

- [ ] **Step 8: Живой прогон** (светлая и тёмная тема Teams; масштаб экрана пользователя):
  - своей полосы нет, у чата +32 px; окно таскается за пустое место шапки, двойной щелчок разворачивает, правый — системное меню; ✕ прячет в трей; поиск, «···», аватар Teams кликаются и не под кнопками окна (лог `controls:` ≈ 138 — в DIP при любом масштабе);
  - ▤ и ⚙ открывают и закрывают панель и настройки; у них окно таскается за верх; цвет кнопок окна — шапки Teams, при открытой панели/настройках — полосы Сороки;
  - F5 и выход на страницу входа (CDP: `location.href = 'https://login.microsoftonline.com/'`) → полоса 32 px; вернулась шапка → полоса ушла; переходов ровно по одному, без мигания;
  - крутилка на ▤ во время «Запропонувати», подсказка — текст статуса;
  - трей → «Перезавантажити Teams» перезагружает; лог `teams check:` — `bar ok`.

- [ ] **Step 9: Коммит** — после просмотра:
```bash
git add src/teams.ts src/main.ts src/preload.ts src/i18n.ts ui/strip.html ui/panel.js ui/frame.js ui/drafts.html ui/summary.html ui/settings.html
git rm ui/menu.html
git commit -m "Soroka: one title bar — ▤ and ⚙ in Teams' own header, which moves the window; Soroka's strip only while Teams has no header; «Сорока ▾» gone, «Перезавантажити Teams» in the tray"
```

---

### Task 4: Карточка звонка над уведомлениями

**Files:**
- Modify: `src/toast.ts` (`corner`, `toast:size`, `toastWin` `closed`), `src/calls.ts` (`ring:size` 40-45, `closeRing`), `src/signin.ts` (размер и закрытие карточки входа)
- Test: `scratchpad\stack-check.js` (не коммитится)

**Interfaces:**
- Produces (`toast.ts`): `export function stackAbove(heights: number[]): number[]` — чистая: снизу вверх по высотам карточек отдаёт, на сколько px от низа рабочей области стоит каждая (`[a, b, c]` → `[0, a, a + b]`); `export function setRingCard(win: BrowserWindow | null, height?: number)` — запоминает карточку звонка и ставит её; `export function restack()` — ставит карточку входа (низ), уведомление над ней, звонок над ними (видимые окна; невидимые — высота 0).
- Consumes: `signInCard` из `signin.ts` (уже импортируется в `toast.ts`).

- [ ] **Step 1:** `scratchpad\stack-check.js`: `stackAbove([])` → `[]`; `stackAbove([120])` → `[0]`; `stackAbove([120, 200, 90])` → `[0, 120, 320]`; `stackAbove([0, 200, 90])` → `[0, 0, 200]` (карточки входа нет).
- [ ] **Step 2:** `node scratchpad\stack-check.js` → падает (`stackAbove is not a function`).
- [ ] **Step 3:** `stackAbove`, `restack`, `setRingCard` в `toast.ts`; `toast:size` и `closed` уведомления, размер и `closed` карточки входа — через `restack()`; `ring:size` → `setRingCard(ringWin, height)` (вместо `corner(..., 'top')`, `setAlwaysOnTop(true, 'screen-saver')` остаётся), `closeRing`/`closed` → `setRingCard(null)`. Вариант `'top'` у `corner` убрать, если больше не нужен.
- [ ] **Step 4:** `npm run build`; `node scratchpad\stack-check.js` → код 0.
- [ ] **Step 5: Живой прогон:** dev-сборка с `--inspect=9339` (инспектор главного процесса); через него `process.mainModule.require('./calls').onRing({ kind: 'call', title: 'Тест', message: 'Відеодзвінок', buttons: [{ label: 'Прийняти', reject: false }, { label: 'Відхилити', reject: true }], outline: '' })` — кнопки карточки не нажимать; затем `node tb.mjs notify "Тест" "уведомление"`. Ожидание: карточка звонка внизу справа; пришло уведомление → звонок над ним; уведомление закрылось → звонок опустился; `onRing(null)` → карточка закрылась.
- [ ] **Step 6: Коммит** — после просмотра:
```bash
git add src/toast.ts src/calls.ts src/signin.ts
git commit -m "Soroka: the call card comes up at the bottom right with the notification cards, above them, and moves as they come and go"
```

---

### Task 5: Итоговая проверка

- [ ] **Step 1:** Пройти весь раздел «Проверка» спеки подряд на dev-прогоне, светлая и тёмная тема Teams; расхождения — в задачу, к которой они относятся, и повторить её живой шаг.
- [ ] **Step 2:** Выйти из dev (`browser-close.mjs 9335`), запустить установленную Сороку; сказать пользователю, что версию и сборку инсталлятора делает он (или по его слову).
