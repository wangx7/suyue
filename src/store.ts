// 阅读位置（localStorage）和最近打开（IndexedDB，可以存文件句柄）

export interface ReadState {
  zoom?: number;
  /** 滚动位置占总高度的比例（文档类） */
  scroll?: number;
  view?: 'read' | 'page';
  sheet?: number;
  cell?: [number, number];
  sx?: number;
  sy?: number;
}

const POS_KEY = 'anyue.positions';
const POS_LIMIT = 200;

type PosMap = Record<string, ReadState & { t: number }>;

function readPositions(): PosMap {
  try {
    return JSON.parse(localStorage.getItem(POS_KEY) ?? '{}') as PosMap;
  } catch {
    return {};
  }
}

export function fileKey(file: File): string {
  return `f:${file.name}:${file.size}:${file.lastModified}`;
}

export function urlKey(url: string): string {
  return `u:${url}`;
}

export function loadState(key: string): ReadState {
  return readPositions()[key] ?? {};
}

export function saveState(key: string, state: ReadState) {
  const all = readPositions();
  all[key] = { ...state, t: Date.now() };
  const keys = Object.keys(all);
  if (keys.length > POS_LIMIT) {
    keys.sort((a, b) => all[a].t - all[b].t).slice(0, keys.length - POS_LIMIT).forEach((k) => delete all[k]);
  }
  try {
    localStorage.setItem(POS_KEY, JSON.stringify(all));
  } catch {
    /* 存储满或不可用时放弃 */
  }
}

// ---------- 最近打开 ----------
export interface RecentItem {
  id: string;
  name: string;
  kind: 'docx' | 'doc' | 'sheet' | 'md';
  size?: number;
  time: number;
  url?: string;
  handle?: FileSystemFileHandle;
}

const RECENT_LIMIT = 12;
let dbPromise: Promise<IDBDatabase> | null = null;

function db(): Promise<IDBDatabase> {
  dbPromise ??= new Promise((resolve, reject) => {
    const req = indexedDB.open('anyue', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('recent', { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

function tx<T>(mode: IDBTransactionMode, run: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return db().then(
    (d) =>
      new Promise<T>((resolve, reject) => {
        const req = run(d.transaction('recent', mode).objectStore('recent'));
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      }),
  );
}

export async function listRecent(): Promise<RecentItem[]> {
  try {
    const all = await tx<RecentItem[]>('readonly', (s) => s.getAll());
    return all.sort((a, b) => b.time - a.time);
  } catch {
    return [];
  }
}

export async function addRecent(item: RecentItem) {
  try {
    await tx('readwrite', (s) => s.put(item));
    const all = await listRecent();
    for (const old of all.slice(RECENT_LIMIT)) await removeRecent(old.id);
  } catch {
    /* IndexedDB 不可用（如隐私模式）时不记录 */
  }
}

export async function removeRecent(id: string) {
  try {
    await tx('readwrite', (s) => s.delete(id));
  } catch {
    /* 忽略 */
  }
}
