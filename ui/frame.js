// The top of a page that starts at the window's top (the panel, the settings) while Teams' header is the window's title bar
// (src/main.ts frame()): --frame-top is the band that moves the window, --frame-right the width Windows' — ☐ ✕ take over its
// right end; html.framed while there is a band. With Soroka's own strip the page sits under it: no band.
(() => {
  const set = f => {
    const root = document.documentElement;
    root.style.setProperty('--frame-top', `${f.top}px`);
    root.style.setProperty('--frame-right', `${f.right}px`);
    root.classList.toggle('framed', f.top > 0);
  };
  bridge.on('frame', set);
  bridge.invoke('frame:get').then(set);
})();
