// 扩展后台：点图标打开查看器；右键文档链接“用素阅打开”；本地文档拖进标签页时改用素阅打开
const EXTS = [
  'docx', 'docm', 'doc', 'xlsx', 'xlsm', 'xlsb', 'xls', 'csv', 'tsv', 'ods', 'md', 'markdown', 'mdown', 'mkd', 'mdx',
  'pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'epub',
];
const LINK_PATTERNS = EXTS.flatMap((e) => [`*://*/*.${e}`, `*://*/*.${e}?*`, `file:///*.${e}`]);
const MENU_ID = 'open-in-anyue';

function viewerUrl(url?: string): string {
  const page = chrome.runtime.getURL('index.html');
  return url ? `${page}?url=${encodeURIComponent(url)}` : page;
}

function openViewer(url?: string) {
  chrome.tabs.create({ url: viewerUrl(url) });
}

// 同一个文件可能同时触发下载和标签页跳转（比如 CSV），几秒内只打开一次
const claimed = new Map<string, number>();
function claim(url: string): boolean {
  const now = Date.now();
  if (now - (claimed.get(url) ?? 0) < 3000) return false;
  claimed.set(url, now);
  return true;
}

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: MENU_ID,
    title: '用素阅打开',
    contexts: ['link'],
    targetUrlPatterns: LINK_PATTERNS,
  });
});

chrome.contextMenus.onClicked.addListener((info) => {
  if (info.menuItemId === MENU_ID && info.linkUrl) openViewer(info.linkUrl);
});

chrome.action.onClicked.addListener(() => openViewer());

function isDocUrl(url: string): boolean {
  const m = /\.([a-z0-9]+)$/i.exec(url.split(/[?#]/)[0]);
  return !!m && EXTS.includes(m[1].toLowerCase());
}

function localPath(url: string): string {
  try {
    return decodeURIComponent(new URL(url).pathname);
  } catch {
    return '';
  }
}

// 把本地文件拖到任意标签页，Chrome 会打开 file:// 地址，显示不了的格式就当成下载。
// 内置页面（新标签页等）不能注入脚本，只能在下载这一步拦下来改用素阅打开
chrome.downloads.onCreated.addListener(async (item) => {
  const url = item.finalUrl || item.url;
  if (!url.startsWith('file:') || !isDocUrl(url)) return;
  if (claim(url)) openViewer(url);
  try {
    await chrome.downloads.cancel(item.id);
    // 本地文件复制得很快，可能来不及取消：删掉复制到下载目录的那份。
    // Chrome 下载时会给重名文件改名，不会落到原文件上，这里再核对一次路径
    const [now] = await chrome.downloads.search({ id: item.id });
    if (now?.state === 'complete' && now.exists && now.filename !== localPath(url)) await chrome.downloads.removeFile(item.id);
    await chrome.downloads.erase({ id: item.id });
  } catch {
    /* 下载已被别处处理，忽略 */
  }
});

// Markdown 这类文本格式 Chrome 能直接显示原文，不走下载：把显示原文的标签页换成素阅。
// 有 file:///* 主机权限才拿得到这些标签页的地址
chrome.tabs.onUpdated.addListener((tabId, info) => {
  const url = info.url;
  if (!url?.startsWith('file:') || !isDocUrl(url) || !claim(url)) return;
  chrome.tabs.update(tabId, { url: viewerUrl(url) });
});
