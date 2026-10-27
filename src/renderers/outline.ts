import type { OutlineItem, Shell } from '../shell';

export interface Heading {
  el: HTMLElement;
  item: OutlineItem;
}

/** 把文档中的标题同步到侧栏目录：点击跳转，滚动时高亮当前章节 */
export function trackHeadings(scroller: HTMLElement, shell: Shell) {
  let list: Heading[] = [];
  let tops: number[] = [];
  let raf = 0;
  let lockedTo = ''; // 点击目录后平滑滚动期间，不让途经的标题抢走高亮

  const measure = () => {
    const base = scroller.getBoundingClientRect().top - scroller.scrollTop;
    tops = list.map((h) => h.el.getBoundingClientRect().top - base);
  };

  const update = () => {
    raf = 0;
    if (!list.length || lockedTo) return;
    if (tops.length !== list.length) measure();
    let i = 0;
    if (scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 4) {
      i = list.length - 1; // 滚到底时最后几个标题到不了顶部，直接高亮最后一个
    } else {
      const y = scroller.scrollTop + 80;
      let lo = 0;
      let hi = tops.length - 1;
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (tops[mid] <= y) lo = mid;
        else hi = mid - 1;
      }
      i = lo;
    }
    shell.setActive(list[i].item.id);
  };
  const schedule = () => {
    raf ||= requestAnimationFrame(update);
  };
  const invalidate = () => {
    tops = [];
    schedule();
  };
  const unlock = () => {
    lockedTo = '';
  };

  scroller.addEventListener('scroll', schedule, { passive: true });
  scroller.addEventListener('scrollend', unlock);
  // 图片加载、缩放、窗口变化都会让标题位置改变
  const ro = new ResizeObserver(invalidate);
  ro.observe(scroller);
  if (scroller.firstElementChild) ro.observe(scroller.firstElementChild);

  return {
    set(headings: Heading[], emptyText: string) {
      list = headings;
      tops = [];
      shell.setOutline(
        headings.map((h) => h.item),
        (id) => {
          const i = list.findIndex((h) => h.item.id === id);
          if (i < 0) return;
          measure();
          lockedTo = id;
          scroller.scrollTo({ top: Math.max(0, tops[i] - 24), behavior: 'smooth' });
          window.setTimeout(unlock, 1200); // 不支持 scrollend 或无需滚动时的兜底
        },
        emptyText,
      );
      schedule();
    },
    refresh: invalidate,
    destroy() {
      cancelAnimationFrame(raf);
      ro.disconnect();
      scroller.removeEventListener('scroll', schedule);
      scroller.removeEventListener('scrollend', unlock);
    },
  };
}
