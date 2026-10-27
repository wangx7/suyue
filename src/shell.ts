// 界面骨架：左侧目录 + 正文 + 右上角几个按钮，没有顶栏。
// 向下阅读时按钮淡出，向上滚动或鼠标移到顶部时出现。
import { icon, type IconName } from './icons';
import type { RecentItem } from './store';
import { cycleTheme, getThemeMode, onThemeChange, THEME_LABEL } from './theme';
import { el } from './ui';

export interface OutlineItem {
  id: string;
  title: string;
  level: number;
  meta?: string;
}

export interface Action {
  icon: IconName;
  title: string;
  run: () => void;
}

interface ShellOptions {
  onHome(): void;
  onZoomReset(): void;
}

// 存储键沿用旧名，改名后用户的设置和最近打开不丢
const SIDE_KEY = 'anyue.side';
const root = document.documentElement;

function fab(name: IconName, title: string, run: () => void): HTMLButtonElement {
  const b = el('button', 'fab-btn');
  b.title = title;
  b.setAttribute('aria-label', title);
  b.append(icon(name));
  b.addEventListener('click', run);
  return b;
}

function relTime(t: number): string {
  const m = Math.round((Date.now() - t) / 60000);
  if (m < 1) return '刚刚';
  if (m < 60) return `${m} 分钟前`;
  if (m < 60 * 24) return `${Math.round(m / 60)} 小时前`;
  return `${Math.round(m / 60 / 24)} 天前`;
}

export function recentList(items: RecentItem[], opts: { onOpen(item: RecentItem): void; onRemove(item: RecentItem): void }): HTMLUListElement {
  const ul = el('ul', 'recent-list');
  for (const item of items) {
    const li = el('li', 'recent-item');
    const open = el('button', 'recent-open');
    open.title = item.url ?? item.name;
    open.append(el('span', 'recent-name', item.name), el('span', 'recent-time', relTime(item.time)));
    open.addEventListener('click', () => opts.onOpen(item));
    const del = el('button', 'recent-del');
    del.title = '从列表移除';
    del.setAttribute('aria-label', '从列表移除');
    del.append(icon('close'));
    del.addEventListener('click', () => opts.onRemove(item));
    li.append(open, del);
    ul.append(li);
  }
  return ul;
}

export function createShell(opts: ShellOptions) {
  // ---------- 侧栏：文档名 + 目录 ----------
  const side = el('aside', 'side');
  const title = el('div', 'side-title');
  const nav = el('nav', 'side-nav');
  side.append(title, nav);
  const mask = el('div', 'side-mask');
  const stage = el('main', 'stage');

  // ---------- 悬浮按钮 ----------
  const left = el('div', 'fab fab-left');
  left.append(fab('sidebar', '目录 (⌘/Ctrl + \\)', () => toggleSide()));
  const actionsBox = el('div', 'fab-actions');
  const zoomPill = el('button', 'fab-btn zoom-pill');
  zoomPill.title = '恢复 100%';
  zoomPill.addEventListener('click', opts.onZoomReset);
  const themeBtn = fab('auto', '', () => {
    cycleTheme();
    hud(`主题：${THEME_LABEL[getThemeMode()]}`);
  });
  const right = el('div', 'fab fab-right');
  const homeBtn = fab('home', '首页', opts.onHome);
  homeBtn.classList.add('fab-home');
  right.append(actionsBox, zoomPill, themeBtn, homeBtn);

  const hudEl = el('div', 'hud');
  document.body.append(side, mask, stage, left, right, hudEl);

  const renderTheme = () => {
    const mode = getThemeMode();
    themeBtn.replaceChildren(icon(mode === 'system' ? 'auto' : mode === 'dark' ? 'moon' : 'sun'));
    themeBtn.title = `主题：${THEME_LABEL[mode]}`;
  };
  renderTheme();
  onThemeChange(renderTheme);

  // ---------- 侧栏开合：宽屏记住选择；窄屏（≤960px）作为浮层临时展开 ----------
  const narrow = matchMedia('(max-width: 960px)');
  let sideOpen = true;
  try {
    sideOpen = localStorage.getItem(SIDE_KEY) !== '0';
  } catch {
    /* 默认展开 */
  }
  root.classList.toggle('side-collapsed', !sideOpen);
  function toggleSide(force?: boolean) {
    if (narrow.matches) {
      root.classList.toggle('side-expanded', force);
      return;
    }
    sideOpen = force ?? !sideOpen;
    root.classList.toggle('side-collapsed', !sideOpen);
    try {
      localStorage.setItem(SIDE_KEY, sideOpen ? '1' : '0');
    } catch {
      /* 忽略 */
    }
  }
  mask.addEventListener('click', () => toggleSide(false));
  narrow.addEventListener('change', () => root.classList.remove('side-expanded'));

  // ---------- 提示 ----------
  let hudTimer = 0;
  function hud(text: string, ms = 900) {
    hudEl.textContent = text;
    hudEl.classList.add('show');
    clearTimeout(hudTimer);
    hudTimer = window.setTimeout(() => hudEl.classList.remove('show'), ms);
  }

  // ---------- 阅读时按钮淡出 ----------
  let watched: HTMLElement | null = null;
  let lastTop = 0;
  const setCalm = (on: boolean) => root.classList.toggle('calm', on);
  const onScroll = () => {
    const top = watched?.scrollTop ?? 0;
    const delta = top - lastTop;
    lastTop = top;
    stage.classList.toggle('scrolled', top > 8);
    if (top < 80 || delta < -6) setCalm(false);
    else if (delta > 6) setCalm(true);
  };
  function watchScroll(scroller: HTMLElement | null) {
    watched?.removeEventListener('scroll', onScroll);
    watched = scroller;
    lastTop = scroller?.scrollTop ?? 0;
    watched?.addEventListener('scroll', onScroll, { passive: true });
    setCalm(false);
    onScroll();
  }
  window.addEventListener('mousemove', (e) => e.clientY < 72 && setCalm(false), { passive: true });

  // ---------- 目录 ----------
  let links = new Map<string, HTMLAnchorElement>();
  let activeId = '';
  function setOutline(items: OutlineItem[] | null, onPick?: (id: string) => void, emptyText = '') {
    links = new Map();
    activeId = '';
    nav.replaceChildren();
    if (!items?.length) {
      if (emptyText) nav.append(el('div', 'side-empty', emptyText));
      return;
    }
    const ul = el('ul', 'toc');
    for (const item of items) {
      const li = el('li', `toc-item toc-l${Math.min(item.level, 6)}`);
      const a = el('a', 'toc-link');
      a.href = '#';
      a.style.paddingLeft = `calc(${item.level - 1} * 0.9em + 24px)`;
      a.append(el('span', 'toc-title', item.title));
      if (item.meta) a.append(el('span', 'toc-meta', item.meta));
      a.title = item.title;
      a.addEventListener('click', (e) => {
        e.preventDefault();
        onPick?.(item.id);
        setActive(item.id);
        if (narrow.matches) toggleSide(false);
      });
      links.set(item.id, a);
      li.append(a);
      ul.append(li);
    }
    nav.append(ul);
  }
  function setActive(id: string) {
    if (id === activeId) return;
    links.get(activeId)?.classList.remove('active');
    activeId = id;
    const a = links.get(id);
    if (!a) return;
    a.classList.add('active');
    // 让当前项保持在侧栏可视范围内
    const nr = nav.getBoundingClientRect();
    const ar = a.getBoundingClientRect();
    if (ar.top < nr.top + 8 || ar.bottom > nr.bottom - 8) nav.scrollTop += ar.top - nr.top - nr.height / 3;
  }

  return {
    stage,
    /** 当前文档；null 表示首页（隐藏侧栏和文档相关按钮） */
    setDoc(name: string | null, info = '') {
      title.textContent = name ?? '';
      title.title = name ? `${name}\n${info}` : '';
      root.classList.toggle('home', !name);
    },
    setOutline,
    setActive,
    setActions(actions: Action[]) {
      actionsBox.replaceChildren(...actions.map((a) => fab(a.icon, a.title, a.run)));
    },
    setZoom(z: number | null) {
      zoomPill.textContent = z == null ? '' : `${Math.round(z * 100)}%`;
      zoomPill.hidden = z == null || Math.abs(z - 1) < 0.005;
    },
    hud,
    watchScroll,
    toggleSide,
  };
}

export type Shell = ReturnType<typeof createShell>;

/** 只在 live() 为真时才转发调用，防止慢的旧文件渲染完后改掉新文件的侧栏 */
export function scopedShell(shell: Shell, live: () => boolean): Shell {
  return new Proxy(shell, {
    get(target, prop) {
      const v = target[prop as keyof Shell];
      return typeof v === 'function' ? (...args: unknown[]) => (live() ? (v as (...a: unknown[]) => unknown)(...args) : undefined) : v;
    },
  });
}
