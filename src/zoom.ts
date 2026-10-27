// 触控板双指缩放。Chrome / Edge / Firefox 把捏合手势报告为带 ctrlKey 的 wheel 事件，
// Safari 用 gesture* 事件。按住 Ctrl 滚动鼠标滚轮也走同一条路径。

export const ZOOM_MIN = 0.25;
export const ZOOM_MAX = 4;

export interface Zoomable {
  get(): number;
  /** anchor 为指针相对于缩放容器左上角的坐标，实现方应让该点下的内容保持不动 */
  set(z: number, anchor: { x: number; y: number }): void;
}

export const clampZoom = (z: number) => Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, z));

interface GestureEvent extends UIEvent {
  scale: number;
  clientX: number;
  clientY: number;
}

export function attachZoom(target: HTMLElement, current: () => Zoomable | undefined, onChange: (z: number) => void) {
  let factor = 1;
  let anchor = { x: 0, y: 0 };
  let raf = 0;

  const flush = () => {
    raf = 0;
    const z = current();
    if (!z) return void (factor = 1);
    const next = clampZoom(z.get() * factor);
    factor = 1;
    if (Math.abs(next - z.get()) < 0.001) return;
    z.set(next, anchor);
    onChange(next);
  };
  const queue = (f: number, clientX: number, clientY: number) => {
    factor *= f;
    const r = target.getBoundingClientRect();
    anchor = { x: clientX - r.left, y: clientY - r.top };
    raf ||= requestAnimationFrame(flush);
  };

  // 挂在 window 上：指针在侧栏上捏合时也要阻止浏览器整页缩放
  window.addEventListener(
    'wheel',
    (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      const dy = e.deltaMode === 1 ? e.deltaY * 16 : e.deltaY;
      // 触控板每次增量很小；鼠标滚轮一格约 100，截断后一格约缩放 1.28 倍
      queue(Math.exp(-Math.max(-25, Math.min(25, dy)) * 0.01), e.clientX, e.clientY);
    },
    { passive: false },
  );

  let last = 1;
  window.addEventListener('gesturestart', (e) => {
    e.preventDefault();
    last = 1;
  });
  window.addEventListener('gesturechange', (e) => {
    e.preventDefault();
    const g = e as GestureEvent;
    queue(g.scale / last, g.clientX, g.clientY);
    last = g.scale;
  });
}

/** 文档类视图的通用实现：给内容元素设置 CSS zoom，并按锚点修正滚动位置 */
export function cssZoom(scroller: HTMLElement, content: HTMLElement, initial = 1, after?: () => void): Zoomable {
  let z = initial;
  content.style.zoom = String(z);
  return {
    get: () => z,
    set(next, anchor) {
      const ratio = next / z;
      const px = scroller.scrollLeft + anchor.x;
      const py = scroller.scrollTop + anchor.y;
      z = next;
      content.style.zoom = String(z);
      scroller.scrollLeft = px * ratio - anchor.x;
      scroller.scrollTop = py * ratio - anchor.y;
      after?.();
    },
  };
}
