// 扩展后台：点图标打开查看器；右键文档链接“用素阅打开”
const EXTS = ['docx', 'doc', 'xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'ods', 'md', 'markdown'];
const LINK_PATTERNS = EXTS.flatMap((e) => [`*://*/*.${e}`, `*://*/*.${e}?*`, `file:///*.${e}`]);
const MENU_ID = 'open-in-anyue';

function openViewer(url?: string) {
  const page = chrome.runtime.getURL('index.html');
  chrome.tabs.create({ url: url ? `${page}?url=${encodeURIComponent(url)}` : page });
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
