// VS Code 插件：把 Word、表格、Markdown、CSV 用素阅的界面打开。
// 扩展这边只负责读文件、存阅读位置和打开链接，渲染全在 webview 里，和网页版共用代码。
import * as vscode from 'vscode';

// Word、表格默认用素阅打开；Markdown、CSV 是文本，默认仍是文本编辑器，素阅作为可选的打开方式。
// 表格画在 Canvas 上，浏览器的查找搜不到，用自带搜索，所以不开 VS Code 的查找框
const VIEW_TYPES: Record<string, { find: boolean }> = {
  'suyue.word': { find: true },
  'suyue.sheet': { find: false },
  'suyue.markdown': { find: true },
  'suyue.csv': { find: false },
};
const TEXT_VIEW: Record<string, string> = {
  md: 'suyue.markdown',
  markdown: 'suyue.markdown',
  mdown: 'suyue.markdown',
  mkd: 'suyue.markdown',
  mdx: 'suyue.markdown',
  csv: 'suyue.csv',
  tsv: 'suyue.csv',
};
// webview 的 localStorage 不保证跨会话保留，主题和侧栏开合交给扩展存
const PREF_KEYS = ['anyue.theme', 'anyue.side'];
const POS_KEY = 'positions';
const POS_LIMIT = 200;

type Positions = Record<string, { state: unknown; t: number }>;

type Message =
  | { type: 'ready' }
  | { type: 'state'; state: unknown }
  | { type: 'pref'; key: string; value: string }
  | { type: 'open'; path: string };

function basename(uri: vscode.Uri): string {
  return decodeURIComponent(uri.path.split('/').pop() ?? '');
}

function extOf(uri: vscode.Uri): string {
  return /\.([a-z0-9]+)$/i.exec(uri.path)?.[1].toLowerCase() ?? '';
}

function escapeAttr(s: string): string {
  return s.replace(/[&"<>]/g, (c) => `&#${c.charCodeAt(0)};`);
}

class Provider implements vscode.CustomReadonlyEditorProvider {
  constructor(private readonly ctx: vscode.ExtensionContext) {}

  openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
    return { uri, dispose() {} };
  }

  resolveCustomEditor(doc: vscode.CustomDocument, panel: vscode.WebviewPanel) {
    const { webview } = panel;
    const dir = vscode.Uri.joinPath(doc.uri, '..');
    const media = vscode.Uri.joinPath(this.ctx.extensionUri, 'dist', 'webview');
    // 文档所在目录和工作区都放开，Markdown 里相对路径的图片才能显示
    const folders = vscode.workspace.workspaceFolders?.map((f) => f.uri) ?? [];
    webview.options = { enableScripts: true, localResourceRoots: [media, dir, ...folders] };
    webview.html = this.html(webview, media);

    const name = basename(doc.uri);
    const key = doc.uri.toString();
    const send = async () => {
      try {
        // readFile 给的可能是 Node 的 Buffer，postMessage 不会按二进制传；复制成普通的 Uint8Array
        const data = new Uint8Array(await vscode.workspace.fs.readFile(doc.uri));
        const saved = this.ctx.globalState.get<Positions>(POS_KEY, {})[key]?.state ?? {};
        await webview.postMessage({ type: 'load', name, data, saved, base: `${webview.asWebviewUri(dir)}/` });
      } catch (err) {
        vscode.window.showErrorMessage(`素阅无法读取「${name}」：${err instanceof Error ? err.message : err}`);
      }
    };

    // 文件在磁盘上改动（比如旁边的文本编辑器保存了 Markdown）就重新加载
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(dir, name));
    const subs = [
      watcher,
      watcher.onDidChange(send),
      watcher.onDidCreate(send),
      webview.onDidReceiveMessage((msg: Message) => {
        if (msg.type === 'ready') send();
        else if (msg.type === 'state') this.savePosition(key, msg.state);
        else if (msg.type === 'pref' && PREF_KEYS.includes(msg.key)) this.ctx.globalState.update(`pref:${msg.key}`, msg.value);
        else if (msg.type === 'open') openLink(msg.path, dir);
      }),
    ];
    panel.onDidDispose(() => subs.forEach((d) => d.dispose()));
  }

  private savePosition(key: string, state: unknown) {
    const all = { ...this.ctx.globalState.get<Positions>(POS_KEY, {}), [key]: { state, t: Date.now() } };
    const keys = Object.keys(all);
    if (keys.length > POS_LIMIT) {
      keys.sort((a, b) => all[a].t - all[b].t).slice(0, keys.length - POS_LIMIT).forEach((k) => delete all[k]);
    }
    this.ctx.globalState.update(POS_KEY, all);
  }

  private html(webview: vscode.Webview, media: vscode.Uri): string {
    const src = (file: string) => webview.asWebviewUri(vscode.Uri.joinPath(media, file)).toString();
    const prefs: Record<string, string> = {};
    for (const k of PREF_KEYS) {
      const v = this.ctx.globalState.get<string>(`pref:${k}`);
      if (v != null) prefs[k] = v;
    }
    // docx-preview 会插入 <style>，所以样式要允许 inline；表格解析的 Worker 用 blob 地址建
    const csp = [
      "default-src 'none'",
      `img-src ${webview.cspSource} blob: data: https:`,
      `media-src ${webview.cspSource} blob: data: https:`,
      `style-src ${webview.cspSource} 'unsafe-inline'`,
      `font-src ${webview.cspSource} blob: data:`,
      `script-src ${webview.cspSource}`,
      'worker-src blob:',
      `connect-src ${webview.cspSource}`,
    ].join('; ');
    return `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta http-equiv="Content-Security-Policy" content="${csp}" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="suyue-prefs" content="${escapeAttr(JSON.stringify(prefs))}" />
    <link rel="stylesheet" href="${src('webview.css')}" />
    <script type="module" src="${src('webview.js')}"></script>
  </head>
  <body></body>
</html>`;
  }
}

/** 文档里的链接：网址交给浏览器，相对路径按文档目录解析后在 VS Code 里打开 */
function openLink(path: string, dir: vscode.Uri) {
  if (/^[a-z][a-z0-9+.-]*:/i.test(path)) {
    if (/^(https?|mailto):/i.test(path)) vscode.env.openExternal(vscode.Uri.parse(path));
    return;
  }
  let rel = path.split(/[?#]/)[0];
  try {
    rel = decodeURIComponent(rel);
  } catch {
    /* 保持原样 */
  }
  if (!rel) return;
  const target = vscode.Uri.joinPath(dir, rel);
  const view = TEXT_VIEW[extOf(target)];
  if (view) vscode.commands.executeCommand('vscode.openWith', target, view);
  else vscode.commands.executeCommand('vscode.open', target);
}

export function activate(ctx: vscode.ExtensionContext) {
  const provider = new Provider(ctx);
  for (const [viewType, { find }] of Object.entries(VIEW_TYPES)) {
    ctx.subscriptions.push(
      vscode.window.registerCustomEditorProvider(viewType, provider, {
        webviewOptions: { enableFindWidget: find, retainContextWhenHidden: true },
        supportsMultipleEditorsPerDocument: true,
      }),
    );
  }
  ctx.subscriptions.push(
    vscode.commands.registerCommand('suyue.open', (uri?: vscode.Uri) => {
      const target = uri ?? vscode.window.activeTextEditor?.document.uri;
      const view = target && TEXT_VIEW[extOf(target)];
      if (view) vscode.commands.executeCommand('vscode.openWith', target, view);
    }),
  );
}

export function deactivate() {}
