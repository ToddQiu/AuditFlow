// The native bridge is absent in ordinary browser sessions.
window.addEventListener('pywebviewready', () => {
  const header = document.querySelector('.masthead');
  if (!header || header.classList.contains('desktop-shell')) return;
  const bar = document.createElement('div');
  bar.className = 'desktop-brand pywebview-drag-region';
  bar.innerHTML = `<div class="window-controls" role="group" aria-label="窗口控制">
    <button class="close" aria-label="关闭窗口" title="关闭窗口"><span class="signal"><svg viewBox="0 0 10 10" aria-hidden="true"><path d="M2 2l6 6M8 2L2 8"/></svg></span></button>
    <button class="minimize" aria-label="最小化" title="最小化"><span class="signal"><svg viewBox="0 0 10 10" aria-hidden="true"><path d="M1.5 5h7"/></svg></span></button>
    <button class="maximize" aria-label="最大化或还原" title="最大化或还原"><span class="signal"><svg viewBox="0 0 10 10" aria-hidden="true"><path d="M2 6V2h4M8 4v4H4"/></svg></span></button>
  </div>`;
  const api = window.pywebview.api;
  const run = async (action) => {
    try { await action(); }
    catch (error) { alert(`窗口操作失败：${error.message || error}`); }
  };
  bar.querySelector('.close').onclick = () => run(api.close);
  bar.querySelector('.minimize').onclick = () => run(api.minimize);
  bar.querySelector('.maximize').onclick = () => run(api.toggle_maximize);
  header.ondblclick = (event) => {
    if (event.target === header || event.target === bar) run(api.toggle_maximize);
  };
  bar.append(header.querySelector('.wordmark'));
  header.prepend(bar);
  header.classList.add('desktop-shell', 'pywebview-drag-region');
  const grip = document.createElement('button');
  grip.className = 'window-resize';
  grip.setAttribute('aria-label', '调整窗口大小（方向键或拖动）');
  grip.title = '拖动调整窗口大小';
  let drag;
  let pendingSize;
  let resizing = false;
  const resize = async (width, height) => {
    pendingSize = [width, height];
    if (resizing) return;
    resizing = true;
    try {
      while (pendingSize) {
        const size = pendingSize;
        pendingSize = null;
        await api.resize(...size);
      }
    } finally { resizing = false; }
  };
  grip.onpointerdown = (e) => {
    if (e.button !== 0) return;
    drag = { x: e.screenX, y: e.screenY, width: innerWidth, height: innerHeight };
    grip.setPointerCapture(e.pointerId);
    e.preventDefault();
  };
  grip.onpointermove = (e) => {
    if (drag) run(() => resize(drag.width + e.screenX - drag.x, drag.height + e.screenY - drag.y));
  };
  grip.onlostpointercapture = () => { drag = null; };
  grip.onkeydown = (e) => {
    const delta = {ArrowLeft: [-20, 0], ArrowRight: [20, 0], ArrowUp: [0, -20], ArrowDown: [0, 20]}[e.key];
    if (delta) { e.preventDefault(); run(() => resize(innerWidth + delta[0], innerHeight + delta[1])); }
  };
  document.body.append(grip);
});
