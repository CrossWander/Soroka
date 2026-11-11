// The app's strings in the chosen language: fills [data-i18n], [data-i18n-title], [data-i18n-placeholder] and [data-i18n-aria],
// and gives scripts t(key, vars), plural(n) and the locale for dates. Pages await `i18n` before using them.
window.i18n = bridge.invoke('i18n:get').then(({ lang, locale, strings }) => {
  document.documentElement.lang = lang;
  window.locale = locale;
  window.t = (key, vars = {}) => (strings[key] ?? key).replace(/\{(\w+)\}/g, (_, k) => vars[k] ?? '');
  const rules = new Intl.PluralRules(lang);
  window.plural = n => `${n} ${strings[`msg.${rules.select(n)}`] ?? strings['msg.other']}`;
  for (const e of document.querySelectorAll('[data-i18n]')) e.textContent = t(e.dataset.i18n);
  for (const e of document.querySelectorAll('[data-i18n-title]')) e.title = t(e.dataset.i18nTitle);
  for (const e of document.querySelectorAll('[data-i18n-placeholder]')) e.placeholder = t(e.dataset.i18nPlaceholder);
  for (const e of document.querySelectorAll('[data-i18n-aria]')) e.setAttribute('aria-label', t(e.dataset.i18nAria));
});
