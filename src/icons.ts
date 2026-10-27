// 线条图标（24×24 网格，描边 1.6，颜色跟随 currentColor）
const PATHS = {
  sidebar: '<rect x="3.5" y="4.5" width="17" height="15" rx="2.5"/><path d="M9.5 4.5v15"/>',
  sun: '<circle cx="12" cy="12" r="3.75"/><path d="M12 3v1.75M12 19.25V21M5.64 5.64l1.24 1.24M17.12 17.12l1.24 1.24M3 12h1.75M19.25 12H21M5.64 18.36l1.24-1.24M17.12 6.88l1.24-1.24"/>',
  moon: '<path d="M19.5 14.5A7.75 7.75 0 0 1 9.5 4.5a7.75 7.75 0 1 0 10 10z"/>',
  auto: '<circle cx="12" cy="12" r="8"/><path d="M12 4a8 8 0 0 0 0 16z" fill="currentColor"/>',
  home: '<path d="M4.5 10.25 12 4.5l7.5 5.75V19a1 1 0 0 1-1 1h-4v-5.5h-5V20h-4a1 1 0 0 1-1-1z"/>',
  page: '<rect x="5.5" y="3.5" width="13" height="17" rx="2"/><path d="M9 8.5h6M9 12h6M9 15.5h3.5"/>',
  read: '<path d="M5 6.5h14M5 10.5h14M5 14.5h14M5 18.5h9"/>',
  search: '<circle cx="11" cy="11" r="6.25"/><path d="m19.5 19.5-4-4"/>',
  close: '<path d="M7 7l10 10M17 7 7 17"/>',
  copy: '<rect x="8.5" y="8.5" width="11" height="11" rx="2"/><path d="M15.5 8.5V6a1.5 1.5 0 0 0-1.5-1.5H6A1.5 1.5 0 0 0 4.5 6v8A1.5 1.5 0 0 0 6 15.5h2.5"/>',
  check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
};

export type IconName = keyof typeof PATHS;

export function icon(name: IconName): SVGSVGElement {
  const wrap = document.createElement('div');
  wrap.innerHTML = `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${PATHS[name]}</svg>`;
  return wrap.firstElementChild as SVGSVGElement;
}
