// 主题：跟随系统 / 深色 / 浅色。深色时文档内容走“智能反色”，浅色时显示原样。
export type ThemeMode = 'system' | 'dark' | 'light';

const KEY = 'anyue.theme';
const ORDER: ThemeMode[] = ['system', 'dark', 'light'];
export const THEME_LABEL: Record<ThemeMode, string> = { system: '跟随系统', dark: '深色', light: '浅色' };

const media = matchMedia('(prefers-color-scheme: dark)');
const listeners = new Set<(dark: boolean) => void>();
let mode: ThemeMode = 'system';

export function isDark(): boolean {
  return mode === 'dark' || (mode === 'system' && media.matches);
}

export function getThemeMode(): ThemeMode {
  return mode;
}

export function onThemeChange(fn: (dark: boolean) => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

function apply() {
  const dark = isDark();
  const root = document.documentElement;
  if (root.dataset.theme === (dark ? 'dark' : 'light')) return;
  root.dataset.theme = dark ? 'dark' : 'light';
  listeners.forEach((fn) => fn(dark));
}

export function initTheme() {
  try {
    const saved = localStorage.getItem(KEY) as ThemeMode | null;
    if (saved && ORDER.includes(saved)) mode = saved;
  } catch {
    /* 存储不可用时使用默认值 */
  }
  apply();
  media.addEventListener('change', () => mode === 'system' && apply());
}

export function cycleTheme(): ThemeMode {
  mode = ORDER[(ORDER.indexOf(mode) + 1) % ORDER.length];
  try {
    localStorage.setItem(KEY, mode);
  } catch {
    /* 忽略 */
  }
  apply();
  return mode;
}
