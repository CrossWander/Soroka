// The panel's head, the same on both of its pages (body data-tab="drafts" | "summary"): the two tabs, "Replies" with the
// count waiting, and ✕. A click on the other tab asks the app to show that page; the colours are the page's own.
(async () => {
  await i18n;
  const style = document.createElement('style');
  style.textContent = `
    .ptabs { flex: none; display: flex; align-items: flex-end; gap: 2px; padding: 6px 6px 0 10px; border-bottom: 1px solid var(--line); }
    .ptab { font: 600 13px "Segoe UI", system-ui, sans-serif; color: var(--muted); background: none; border: 0; border-bottom: 2px solid transparent; border-radius: 0;
      padding: 7px 8px; display: flex; gap: 6px; align-items: center; cursor: pointer; }
    .ptab:hover, .ptab[aria-selected="true"] { color: var(--fg); }
    .ptab[aria-selected="true"] { border-bottom-color: var(--accent); }
    .pbadge { background: #c4314b; color: #fff; border-radius: 9px; font-size: 10px; line-height: 14px; padding: 0 5px; }
    .px { margin-left: auto; align-self: center; background: none; border: 0; color: var(--muted); font-size: 14px; padding: 4px 8px; border-radius: 4px; cursor: pointer; }
    .px:hover { color: var(--fg); }`;
  document.head.append(style);
  const here = document.body.dataset.tab;
  const tab = (name, label) => {
    const b = Object.assign(document.createElement('button'), { className: 'ptab', textContent: label });
    b.setAttribute('role', 'tab'); b.setAttribute('aria-selected', String(name === here));
    b.onclick = () => { if (name !== here) bridge.invoke('panel:tab', name); };
    return b;
  };
  const replies = tab('drafts', t('drafts.head')), badge = Object.assign(document.createElement('span'), { className: 'pbadge', hidden: true });
  replies.append(badge);
  const close = Object.assign(document.createElement('button'), { className: 'px', textContent: '✕', title: t('drafts.hide') });
  close.setAttribute('aria-label', t('drafts.hide'));
  close.onclick = () => bridge.invoke('panel:close');
  const head = Object.assign(document.createElement('header'), { className: 'ptabs' });
  head.setAttribute('role', 'tablist');
  head.append(replies, tab('summary', t('summary.tab')), close);
  document.body.prepend(head);
  const count = async () => { const n = (await bridge.invoke('drafts:get')).length; badge.textContent = n > 9 ? '9+' : n; badge.hidden = !n; };
  bridge.on('drafts', count);
  count();
})();
