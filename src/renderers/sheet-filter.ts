import { el } from '../ui';

// 列筛选菜单：搜索、按值勾选，和 Excel 的自动筛选一样

export interface FilterMenuOptions {
  title: string;
  /** 列里出现的值（已排序）和出现次数 */
  values: [string, number][];
  /** 当前允许显示的值；null 表示未筛选 */
  allowed: Set<string> | null;
  /** 返回 null 表示清除这一列的筛选 */
  onApply(allowed: Set<string> | null): void;
  onClose(): void;
}

const MAX_SHOWN = 300;
const BLANK = '（空白）';

export function showFilterMenu(box: HTMLElement, x: number, y: number, opts: FilterMenuOptions): () => void {
  const menu = el('div', 'filter-menu');
  const title = el('div', 'filter-title', opts.title);
  const search = el('input', 'filter-search');
  search.placeholder = '搜索';
  const ops = el('div', 'filter-ops');
  const selectAll = el('button', 'filter-link', '全选');
  const clearAll = el('button', 'filter-link', '清除');
  const count = el('span', 'filter-count');
  ops.append(selectAll, clearAll, count);
  const list = el('div', 'filter-list');
  const foot = el('div', 'filter-foot');
  const cancel = el('button', 'filter-btn', '取消');
  const ok = el('button', 'filter-btn filter-ok', '确定');
  foot.append(cancel, ok);
  menu.append(title, search, ops, list, foot);

  const checked = new Set(opts.allowed ?? opts.values.map(([v]) => v));
  const matching = () => {
    const q = search.value.trim().toLowerCase();
    return q ? opts.values.filter(([v]) => (v || BLANK).toLowerCase().includes(q)) : opts.values;
  };

  const refresh = () => {
    const matched = matching();
    const shown = matched.slice(0, MAX_SHOWN);
    list.replaceChildren();
    for (const [value, n] of shown) {
      const row = el('label', 'filter-item');
      const cb = el('input');
      cb.type = 'checkbox';
      cb.checked = checked.has(value);
      cb.addEventListener('change', () => {
        if (cb.checked) checked.add(value);
        else checked.delete(value);
        updateCount();
      });
      const text = el('span', value ? 'filter-text' : 'filter-text filter-blank', value || BLANK);
      text.title = value;
      row.append(cb, text, el('span', 'filter-n', String(n)));
      list.append(row);
    }
    if (matched.length > shown.length) list.append(el('div', 'filter-more', `仅显示前 ${MAX_SHOWN} 项，输入关键字缩小范围`));
    if (!matched.length) list.append(el('div', 'filter-more', '没有匹配项'));
    selectAll.textContent = search.value.trim() ? '全选结果' : '全选';
    updateCount();
  };
  const updateCount = () => (count.textContent = `${opts.values.filter(([v]) => checked.has(v)).length} / ${opts.values.length}`);

  // 搜索时“全选 / 清除”只作用于匹配到的项
  selectAll.addEventListener('click', () => {
    for (const [v] of matching()) checked.add(v);
    refresh();
  });
  clearAll.addEventListener('click', () => {
    for (const [v] of matching()) checked.delete(v);
    refresh();
  });
  search.addEventListener('input', refresh);

  const close = () => {
    document.removeEventListener('mousedown', outside, true);
    menu.remove();
    opts.onClose();
  };
  const outside = (e: MouseEvent) => {
    if (!menu.contains(e.target as Node)) close();
  };
  cancel.addEventListener('click', close);
  ok.addEventListener('click', () => {
    const all = opts.values.every(([v]) => checked.has(v));
    opts.onApply(all ? null : new Set(checked));
    close();
  });
  menu.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') close();
    else if (e.key === 'Enter' && e.target === search) ok.click();
  });

  box.append(menu);
  const left = Math.max(4, Math.min(x, box.clientWidth - menu.offsetWidth - 8));
  menu.style.left = `${left}px`;
  menu.style.top = `${y}px`;
  document.addEventListener('mousedown', outside, true);
  refresh();
  search.focus();
  return close;
}
