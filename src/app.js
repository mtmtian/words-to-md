import { zipSync, strToU8 } from 'fflate';
import { MAX_FILE_BYTES, parseDocx, mergeDocuments } from './parser.js';
import demoDocuments from './demos.json';

const $ = id => document.getElementById(id);
const state = { docs: [], errors: [], busy: false, view: 'preview', nextId: 1, markdown: '' };
const blobUrls = new Set();
let toastTimer;

function toast(message) {
  $('toast').textContent = message;
  $('toast').hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { $('toast').hidden = true; }, 4200);
}
function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}
function action(label, symbol, callback, disabled = false, extraClass = '') {
  const button = element('button', `icon-button ${extraClass}`, symbol);
  button.type = 'button';
  button.setAttribute('aria-label', label);
  button.title = label;
  button.disabled = disabled || state.busy;
  button.addEventListener('click', callback);
  return button;
}

function renderQueue() {
  const list = $('file-list');
  list.replaceChildren();
  state.docs.forEach((doc, index) => {
    const row = element('li', 'file-row');
    row.dataset.id = String(doc.id);
    const top = element('div', 'file-top');
    top.append(element('span', 'file-number', String(index + 1).padStart(2, '0')));
    const info = element('div', 'file-info');
    info.append(element('div', 'file-title', doc.title), element('div', 'file-name', doc.filename));
    top.append(info);
    const bottom = element('div', 'file-bottom');
    bottom.append(element('span', 'file-meta', `${doc.titleSource === 'font' ? '已识别标题' : '使用文件名'} · ${doc.headings} 个章节`));
    const buttons = element('div', 'file-actions');
    const move = offset => {
      const next = index + offset;
      [state.docs[index], state.docs[next]] = [state.docs[next], state.docs[index]];
      render();
      const selector = `[data-id="${doc.id}"] button[aria-label^="${offset < 0 ? '上移' : '下移'}"]`;
      const button = document.querySelector(selector);
      if (button && !button.disabled) button.focus();
      else document.querySelector(`[data-id="${doc.id}"] button:not(:disabled)`)?.focus();
      toast(`已调整「${doc.title}」的合并顺序`);
    };
    buttons.append(action(`上移 ${doc.filename}`, '↑', () => move(-1), index === 0));
    buttons.append(action(`下移 ${doc.filename}`, '↓', () => move(1), index === state.docs.length - 1));
    buttons.append(action(`移除 ${doc.filename}`, '×', () => {
      state.docs.splice(index, 1);
      render();
      $('clear-all').disabled ? $('file-input').focus() : $('clear-all').focus();
    }, false, 'remove'));
    bottom.append(buttons);
    row.append(top, bottom);
    if (doc.warnings.length) {
      const details = element('details', 'file-warning');
      details.append(element('summary', '', `${doc.warnings.length} 条提示 · 请核对`));
      const warnings = element('ul');
      doc.warnings.forEach(text => warnings.append(element('li', '', text)));
      details.append(warnings);
      row.append(details);
    }
    list.append(row);
  });
  $('file-count').textContent = state.docs.length;
  $('empty-queue').hidden = state.docs.length > 0;
  const errors = $('import-errors');
  errors.replaceChildren();
  errors.hidden = !state.errors.length;
  if (state.errors.length) {
    errors.append(element('strong', '', `${state.errors.length} 项未加入合并`));
    state.errors.forEach(error => errors.append(element('p', '', `${error.filename}：${error.message}`)));
    const dismiss = element('button', 'text-button', '收起错误记录');
    dismiss.disabled = state.busy;
    dismiss.addEventListener('click', () => { state.errors = []; renderQueue(); });
    errors.append(dismiss);
  }
}

function renderPreview() {
  const preview = $('preview');
  preview.replaceChildren();
  state.docs.forEach((doc, index) => {
    if (index) preview.append(element('hr'));
    preview.append(element('h1', '', doc.title.replace(/\s+/g, ' ').trim()));
    doc.paragraphs.forEach((p, pIndex) => {
      if (pIndex === doc.titleIndex) return;
      preview.append(element(p.isHeading ? 'h2' : 'p', '', p.isHeading ? p.text.replace(/\s+/g, ' ') : p.text));
    });
  });
}

function updateView() {
  const hasDocs = state.docs.length > 0;
  const isPreview = state.view === 'preview';
  $('output-empty').hidden = hasDocs || !isPreview;
  $('preview').hidden = !hasDocs || !isPreview;
  $('source-panel').hidden = isPreview;
  ['preview', 'source'].forEach(view => {
    $(`tab-${view}`).setAttribute('aria-selected', String(state.view === view));
    $(`tab-${view}`).tabIndex = state.view === view ? 0 : -1;
  });
}

function render() {
  state.markdown = mergeDocuments(state.docs);
  renderQueue();
  renderPreview();
  $('markdown-output').value = state.markdown;
  $('result-stats').textContent = state.docs.length ? `${state.docs.length} 份文档 · ${state.docs.reduce((sum, doc) => sum + doc.headings, 0)} 个章节 · ${[...state.markdown].length.toLocaleString('zh-CN')} 字符` : '等待导入文档';
  $('clear-all').disabled = state.busy || !(state.docs.length || state.errors.length);
  $('load-demo').disabled = state.busy;
  $('file-input').disabled = state.busy;
  for (const id of ['copy-output', 'open-output', 'download-output']) $(id).disabled = state.busy || !state.docs.length;
  $('download-output').setAttribute('aria-label', state.docs.length ? `下载 ${state.docs.length} 份文档合并的 Markdown` : '下载合并的 Markdown');
  updateView();
}

async function addFiles(files) {
  if (state.busy) { toast('正在解析，请稍后再添加文档。'); return; }
  if (!files.length) return;
  state.busy = true;
  const startCount = state.docs.length;
  const allowed = 100 - state.docs.length;
  const pending = [...files].slice(0, allowed);
  if (files.length > allowed) state.errors.push({ filename: '批量导入', message: `最多保留 100 份文档，本次超出的 ${files.length - allowed} 份未处理。` });
  $('progress').hidden = false;
  render();
  try {
    for (let i = 0; i < pending.length; i++) {
      const file = pending[i];
      $('progress').textContent = `正在解析 ${i + 1} / ${pending.length}：${file.name}`;
      await new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
      try {
        if (state.docs.some(doc => doc.filename === file.name && doc.size === file.size && doc.modified === file.lastModified)) {
          state.errors.push({ filename: file.name, message: '已导入相同名称、大小和修改时间的文件，已跳过。' });
          continue;
        }
        if (file.size > MAX_FILE_BYTES) throw new Error('超过 30 MB，请拆分后导入。');
        const parsed = parseDocx(await file.arrayBuffer(), file.name);
        state.docs.push({ ...parsed, id: state.nextId++, size: file.size, modified: file.lastModified });
      } catch (error) {
        state.errors.push({ filename: file.name, message: error.message || '无法解析此文档。' });
      }
    }
  } finally {
    state.busy = false;
    $('progress').hidden = true;
    render();
    toast(`已加入 ${state.docs.length - startCount} 份文档${state.errors.length ? '，请查看导入提示' : '，可以预览或保存了'}`);
  }
}

function bytesFromBase64(data) { return Uint8Array.from(atob(data), char => char.charCodeAt(0)); }
function blobUrl(blob) {
  const url = URL.createObjectURL(blob);
  blobUrls.add(url);
  return url;
}
function download(blob, filename) {
  const url = blobUrl(blob);
  const link = element('a');
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => { URL.revokeObjectURL(url); blobUrls.delete(url); }, 60_000);
}

$('file-input').addEventListener('change', event => {
  const files = [...event.target.files];
  event.target.value = '';
  addFiles(files);
});
let dragDepth = 0;
const zone = $('drop-zone');
for (const event of ['dragenter', 'dragover', 'dragleave', 'drop']) {
  window.addEventListener(event, e => { e.preventDefault(); });
}
zone.addEventListener('dragenter', () => { dragDepth++; zone.classList.add('drag-over'); });
zone.addEventListener('dragleave', () => { if (--dragDepth <= 0) zone.classList.remove('drag-over'); });
zone.addEventListener('drop', event => {
  dragDepth = 0;
  zone.classList.remove('drag-over');
  addFiles([...event.dataTransfer.files]);
});
$('clear-all').addEventListener('click', () => { state.docs = []; state.errors = []; render(); $('file-input').focus(); });
$('load-demo').addEventListener('click', () => addFiles(demoDocuments.map(doc => new File([bytesFromBase64(doc.data)], doc.name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', lastModified: 0 }))));
$('download-demo').addEventListener('click', () => {
  const files = Object.fromEntries(demoDocuments.map(doc => [doc.name, bytesFromBase64(doc.data)]));
  files['示例说明.txt'] = strToU8('四份示例分别覆盖：直接字体、样式继承、主题字体、缺失标题与混合字体。\n按文件名顺序导入，检查每份文档的一级标题与全黑体二级标题。\n文档字体名称写入文件；本机未安装对应字体时，Word 显示可能使用替代字体，不影响本工具读取字体元数据。\n');
  download(new Blob([zipSync(files)], { type: 'application/zip' }), 'Word转Markdown_示例文档.zip');
});
for (const view of ['preview', 'source']) {
  $(`tab-${view}`).addEventListener('click', () => { state.view = view; updateView(); });
  $(`tab-${view}`).addEventListener('keydown', event => {
    if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) {
      event.preventDefault();
      const next = event.key === 'Home' ? 'preview' : event.key === 'End' ? 'source' : view === 'preview' ? 'source' : 'preview';
      state.view = next;
      updateView();
      $(`tab-${next}`).focus();
    }
  });
}
$('download-output').addEventListener('click', () => download(new Blob([state.markdown], { type: 'text/markdown;charset=utf-8' }), '合并文档.md'));
$('open-output').addEventListener('click', () => {
  const popup = window.open('', '_blank');
  if (!popup) { toast('浏览器拦截了新标签页，请允许弹出窗口，或使用“下载 .md”。'); return; }
  popup.opener = null;
  popup.location.replace(blobUrl(new Blob([state.markdown], { type: 'text/plain;charset=utf-8' })));
  toast('已在新标签页打开纯文本，可使用 Ctrl / ⌘ + S 保存。');
});
$('copy-output').addEventListener('click', async () => {
  try {
    if (!navigator.clipboard?.writeText) throw new Error('Clipboard unavailable');
    await navigator.clipboard.writeText(state.markdown);
    toast('Markdown 已复制');
  } catch {
    state.view = 'source';
    updateView();
    const output = $('markdown-output');
    output.focus();
    output.select();
    try {
      if (!document.execCommand('copy')) throw new Error('Copy failed');
      toast('Markdown 已复制');
    } catch { toast('已选中全部 Markdown，请按 Ctrl / ⌘ + C 复制。'); }
  }
});
window.addEventListener('pagehide', () => { for (const url of blobUrls) URL.revokeObjectURL(url); });
// Small documented API for reuse and deterministic browser regression checks.
window.WordMD = Object.freeze({ parseDocx, mergeDocuments, version: '1.0.0' });
render();
