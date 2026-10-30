import { el } from '../ui';
import { cssZoom } from '../zoom';
import { trackHeadings, type Heading } from './outline';
import { drawSlide } from './pptx-draw';
import { loadDeck } from './pptx-parse';
import type { Renderer } from './types';

// 演示文稿：幻灯片按原尺寸绘制，整体缩放到适合窗口宽度，纵向排列；备注显示在每页下方。
// 文字一次全部画出（浏览器 / VS Code 的查找能搜到），图片等滚到附近再解压。

const MAX_WIDTH = 1100;
const SIDE_PAD = 32;

export const renderPptx: Renderer = async (data, { host, shell, saved }) => {
  const deck = await loadDeck(data);
  if (!deck.slides.length) throw new Error('这个演示文稿没有幻灯片');

  const scroller = el('div', 'doc-scroll paper');
  const body = el('div', 'pptx-deck');
  body.style.setProperty('--sw', String(deck.width));
  body.style.setProperty('--sh', String(deck.height));
  scroller.append(body);
  host.append(scroller);

  const fit = () => body.style.setProperty('--fit', String(Math.min(MAX_WIDTH, Math.max(160, scroller.clientWidth - SIDE_PAD * 2)) / deck.width));
  fit();

  const headings: Heading[] = [];
  const frames: HTMLElement[] = [];
  for (const [i, slide] of deck.slides.entries()) {
    const drawn = await drawSlide(deck, slide, i);
    const item = el('section', 'pptx-item');
    const frame = el('div', 'pptx-frame');
    frame.append(drawn.el);
    item.append(frame);
    if (slide.hidden) item.classList.add('pptx-hidden');
    const notes = drawn.notes.filter(Boolean);
    if (notes.length) {
      const box = el('div', 'pptx-notes');
      for (const line of notes) box.append(el('p', '', line));
      item.append(box);
    }
    body.append(item);
    frames.push(frame);
    headings.push({
      el: item,
      item: {
        id: `s${i}`,
        title: drawn.title || `第 ${i + 1} 页`,
        level: 1,
        meta: [drawn.title && `第 ${i + 1} 页`, slide.hidden && '已隐藏'].filter(Boolean).join(' · ') || undefined,
      },
    });
  }

  // 图片延迟加载：幻灯片进入视口前后一屏时才解压
  const load = async (frame: HTMLElement) => {
    for (const img of frame.querySelectorAll<HTMLImageElement>('img[data-src]')) {
      const url = await deck.pkg.url(img.dataset.src!);
      if (url) img.src = url;
      else img.remove();
    }
    for (const n of [...frame.querySelectorAll<HTMLElement>('[data-bg]')]) {
      const url = await deck.pkg.url(n.dataset.bg!);
      if (url) n.style.backgroundImage = `url("${url}")`;
    }
  };
  const io = new IntersectionObserver(
    (entries) => {
      for (const e of entries) {
        if (!e.isIntersecting) continue;
        io.unobserve(e.target);
        void load(e.target as HTMLElement);
      }
    },
    { root: scroller, rootMargin: '100% 0px' },
  );
  frames.forEach((f) => io.observe(f));

  const tracker = trackHeadings(scroller, shell);
  tracker.set(headings, '');
  const zoom = cssZoom(scroller, body, saved.zoom ?? 1, tracker.refresh);

  // 窗口宽度变化时重新适配，保持阅读位置
  let lastWidth = scroller.clientWidth;
  const ro = new ResizeObserver(() => {
    if (scroller.clientWidth === lastWidth) return;
    lastWidth = scroller.clientWidth;
    const ratio = scroller.scrollTop / Math.max(1, scroller.scrollHeight - scroller.clientHeight);
    fit();
    // 位置没变就不写：写 scrollTop 会打断正在进行的平滑滚动（侧栏展开动画期间会连续触发）
    const top = ratio * Math.max(0, scroller.scrollHeight - scroller.clientHeight);
    if (Math.abs(top - scroller.scrollTop) >= 1) scroller.scrollTop = top;
    tracker.refresh();
  });
  ro.observe(scroller);

  scroller.scrollTop = (saved.scroll ?? 0) * Math.max(0, scroller.scrollHeight - scroller.clientHeight);

  return {
    scroller,
    zoom,
    state: () => ({
      zoom: zoom.get(),
      scroll: scroller.scrollTop / Math.max(1, scroller.scrollHeight - scroller.clientHeight),
    }),
    destroy() {
      tracker.destroy();
      io.disconnect();
      ro.disconnect();
      deck.pkg.destroy();
    },
  };
};
