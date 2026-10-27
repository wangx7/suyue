import { parseAsync, renderDocument } from 'docx-preview';
import { darkColor, parseColor } from '../color';
import type { OutlineItem } from '../shell';
import { isDark } from '../theme';
import { openLightbox } from '../lightbox';
import { el } from '../ui';
import { cssZoom } from '../zoom';
import { fixFonts } from './docx-fonts';
import { fitEdgeTabs, fitLayout, markEdgeTabs } from './docx-layout';
import { trackHeadings } from './outline';
import type { Renderer } from './types';

type View = 'read' | 'page';

const BASE = {
  className: 'docx',
  experimental: true,
  renderFootnotes: true,
  renderEndnotes: true,
  renderComments: false,
  renderChanges: false,
};
const VIEW_OPTIONS: Record<View, object> = {
  // 阅读视图：去掉纸张和分页，正文居中流式排版。保留纸张宽度（被 CSS 覆盖），供 fitTables 换算表格比例
  read: { inWrapper: false, ignoreWidth: false, ignoreHeight: true, breakPages: false, renderHeaders: false, renderFooters: false },
  page: { inWrapper: true, breakPages: true, renderHeaders: true, renderFooters: true },
};

/** docx-preview 解析出的样式（只用到这几个字段） */
interface DocStyle {
  id?: string;
  name?: string;
  basedOn?: string;
  target?: string;
  cssName?: string;
  paragraphProps?: { outlineLevel?: number };
}

/** 标题级别：优先看大纲级别（沿 basedOn 继承），再看样式名 “heading 1” / “标题 1” */
function headingLevels(styles: DocStyle[]): Map<string, number> {
  const byId = new Map(styles.map((s) => [s.id, s]));
  const levelOf = (s: DocStyle | undefined, depth = 0): number | undefined => {
    if (!s || depth > 10) return undefined;
    const ol = s.paragraphProps?.outlineLevel;
    if (ol != null && ol < 9) return ol + 1;
    const m = /^(?:heading|标题)\s*(\d)$/i.exec(s.name ?? '');
    if (m) return Number(m[1]);
    return levelOf(s.basedOn ? byId.get(s.basedOn) : undefined, depth + 1);
  };
  const result = new Map<string, number>();
  for (const s of styles) {
    const level = s.target === 'p' && s.cssName ? levelOf(s) : undefined;
    if (level) result.set(s.cssName!, level);
  }
  return result;
}

const SIDES = ['top', 'right', 'bottom', 'left'];

/**
 * 智能深色：把文字、底纹、边框的实际颜色逐个重映射（见 color.ts），图片不动。
 * 先统一读取再统一写入，避免读写交替导致反复样式计算。
 */
function applyDark(root: HTMLElement) {
  const nodes = [...root.querySelectorAll<HTMLElement>('section.docx, section.docx *')].filter(
    (n) => !(n instanceof HTMLImageElement) && !n.closest('svg'),
  );
  const plans = nodes.map((n) => {
    const cs = getComputedStyle(n);
    const borders = SIDES.map((s) => (cs.getPropertyValue(`border-${s}-style`) !== 'none' ? cs.getPropertyValue(`border-${s}-color`) : ''));
    return { n, color: cs.color, bg: cs.backgroundColor, borders };
  });
  const dark = (c: string, role: 'text' | 'fill' = 'fill') => {
    const rgba = parseColor(c);
    return rgba && rgba[3] > 0 ? darkColor(rgba, role) : '';
  };
  for (const { n, color, bg, borders } of plans) {
    const c = dark(color, 'text');
    if (c) n.style.color = c;
    const b = dark(bg);
    if (b) n.style.backgroundColor = b;
    borders.forEach((bc, i) => {
      const d = bc && dark(bc);
      if (d) n.style.setProperty(`border-${SIDES[i]}-color`, d);
    });
  }
}

export const renderDocx: Renderer = async (data, { host, shell, saved }) => {
  const doc = await parseAsync(data, BASE);
  const styles: DocStyle[] = doc.stylesPart?.styles ?? [];

  const scroller = el('div', 'doc-scroll');
  const body = el('div');
  scroller.append(body);
  host.append(scroller);

  let view: View = saved.view ?? 'read';
  const tracker = trackHeadings(scroller, shell);
  const zoom = cssZoom(scroller, body, saved.zoom ?? 1, tracker.refresh);
  body.addEventListener('click', (e) => {
    const img = (e.target as Element).closest?.<HTMLImageElement>('section.docx img');
    if (img?.currentSrc) openLightbox(img);
  });

  async function render(restoreRatio?: number) {
    const ratio = restoreRatio ?? scroller.scrollTop / Math.max(1, scroller.scrollHeight - scroller.clientHeight);
    const read = view === 'read';
    markEdgeTabs(doc, read);
    const nodes = await renderDocument(doc, { ...BASE, ...VIEW_OPTIONS[view] });
    body.className = read ? 'docx-read' : 'docx-page';
    // docx-preview 默认整篇自动断字，Word 只在开启“自动断字”时才断
    body.classList.toggle('docx-hyphenate', !!doc.settingsPart?.settings?.autoHyphenation);
    scroller.classList.toggle('paper', !read);
    body.replaceChildren(...nodes);
    fixFonts(body);
    fitLayout(body, doc, read);
    if (isDark()) applyDark(body);
    // 在深色重映射之后拆分，前导符直接沿用已映射的文字颜色
    if (read) fitEdgeTabs(body);

    // 目录：渲染后样式才有 cssName，按文档顺序收集标题段落
    const levels = headingLevels(styles);
    const headings: { el: HTMLElement; item: OutlineItem }[] = [];
    if (levels.size) {
      for (const p of body.querySelectorAll<HTMLElement>('p')) {
        let level = 0;
        for (const c of p.classList) level ||= levels.get(c) ?? 0;
        const text = p.textContent?.replace(/\s+/g, ' ').trim();
        if (level && text) headings.push({ el: p, item: { id: `h${headings.length}`, title: text, level } });
      }
    }
    tracker.set(headings, '此文档没有标题目录');

    scroller.scrollTop = ratio * Math.max(0, scroller.scrollHeight - scroller.clientHeight);
    renderActions();
  }

  function renderActions() {
    shell.setActions([
      view === 'read'
        ? { icon: 'page', title: '切换到页面视图', run: () => switchView('page') }
        : { icon: 'read', title: '切换到阅读视图', run: () => switchView('read') },
    ]);
  }
  function switchView(next: View) {
    view = next;
    shell.hud(next === 'read' ? '阅读视图' : '页面视图');
    void render();
  }

  await render(saved.scroll ?? 0);

  return {
    scroller,
    zoom,
    onTheme: () => void render(),
    state: () => ({
      view,
      zoom: zoom.get(),
      scroll: scroller.scrollTop / Math.max(1, scroller.scrollHeight - scroller.clientHeight),
    }),
    destroy: tracker.destroy,
  };
};
