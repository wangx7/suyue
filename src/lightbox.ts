// 图片查看器：适应窗口打开；捏合 / Ctrl + 滚轮缩放（指针下的内容不动），拖动或双指滚动平移，
// 双击在适应窗口和原始大小之间切换；Esc、关闭按钮或点空白处关闭。
// 事件在遮罩上截住，不会传到 window 上的文档缩放和快捷键。

import { icon } from './icons';
import { el } from './ui';

const MIN = 0.05;
const MAX = 8;
const GAP = 32;

export function openLightbox(source: HTMLImageElement) {
  const nw = source.naturalWidth || source.width;
  const nh = source.naturalHeight || source.height;
  if (!nw || !nh) return;

  const overlay = el('div', 'lightbox');
  const img = el('img');
  img.src = source.currentSrc;
  img.alt = source.alt;
  img.draggable = false;
  img.style.width = `${nw}px`;
  img.style.height = `${nh}px`;
  const close = el('button', 'lightbox-close');
  close.title = '关闭（Esc）';
  close.append(icon('close'));
  overlay.append(img, close);
  document.body.append(overlay);

  let scale = 1;
  let x = 0;
  let y = 0;
  const apply = () => (img.style.transform = `translate(${x}px, ${y}px) scale(${scale})`);
  // 适应窗口；文档里本来就放大显示的小图，保持文档里的大小
  const fitScale = () =>
    Math.min((innerWidth - GAP * 2) / nw, (innerHeight - GAP * 2) / nh, Math.max(1, source.getBoundingClientRect().width / nw));
  const center = (s: number) => {
    scale = s;
    x = (innerWidth - nw * s) / 2;
    y = (innerHeight - nh * s) / 2;
    apply();
  };
  const zoomAt = (next: number, cx: number, cy: number) => {
    next = Math.min(MAX, Math.max(MIN, next));
    x = cx - ((cx - x) * next) / scale;
    y = cy - ((cy - y) * next) / scale;
    scale = next;
    apply();
  };
  center(fitScale());

  const stop = (e: Event) => {
    e.preventDefault();
    e.stopPropagation();
  };
  overlay.addEventListener(
    'wheel',
    (e) => {
      stop(e);
      const dy = e.deltaMode === 1 ? e.deltaY * 16 : e.deltaY;
      const dx = e.deltaMode === 1 ? e.deltaX * 16 : e.deltaX;
      if (e.ctrlKey) zoomAt(scale * Math.exp(-Math.max(-25, Math.min(25, dy)) * 0.01), e.clientX, e.clientY);
      else {
        x -= dx;
        y -= dy;
        apply();
      }
    },
    { passive: false },
  );
  // Safari 的捏合手势
  let last = 1;
  overlay.addEventListener('gesturestart', (e) => {
    stop(e);
    last = 1;
  });
  overlay.addEventListener('gesturechange', (e) => {
    stop(e);
    const g = e as unknown as { scale: number; clientX: number; clientY: number };
    zoomAt((scale * g.scale) / last, g.clientX, g.clientY);
    last = g.scale;
  });

  // 拖动平移；拖过的那次点击不算“点空白处关闭”
  let drag: { id: number; sx: number; sy: number; x: number; y: number } | null = null;
  let moved = false;
  img.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    drag = { id: e.pointerId, sx: e.clientX, sy: e.clientY, x, y };
    moved = false;
    img.setPointerCapture(e.pointerId);
    img.classList.add('dragging');
  });
  img.addEventListener('pointermove', (e) => {
    if (drag?.id !== e.pointerId) return;
    moved ||= Math.hypot(e.clientX - drag.sx, e.clientY - drag.sy) > 3;
    x = drag.x + e.clientX - drag.sx;
    y = drag.y + e.clientY - drag.sy;
    apply();
  });
  const endDrag = () => {
    drag = null;
    img.classList.remove('dragging');
  };
  img.addEventListener('pointerup', endDrag);
  img.addEventListener('pointercancel', endDrag);
  img.addEventListener('dblclick', (e) => {
    const fit = fitScale();
    if (Math.abs(scale - fit) < 0.01) zoomAt(1, e.clientX, e.clientY);
    else center(fit);
  });
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay && !moved) dismiss();
    moved = false;
  });
  close.addEventListener('click', dismiss);

  const onKey = (e: KeyboardEvent) => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      dismiss();
    }
  };
  const onResize = () => center(fitScale());
  window.addEventListener('keydown', onKey, true);
  window.addEventListener('resize', onResize);
  close.focus();

  function dismiss() {
    window.removeEventListener('keydown', onKey, true);
    window.removeEventListener('resize', onResize);
    overlay.remove();
  }
}
