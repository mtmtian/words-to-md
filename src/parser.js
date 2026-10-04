import { unzipSync } from 'fflate';

const W = new Set(['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main']);
const MiB = 1024 * 1024;
export const MAX_FILE_BYTES = 30 * MiB;
// 方正小标宋_GBK and its English name. FZXiaoBiaoSong-B05S is 方正小标宋简体, a separate GB2312 font.
const TITLE_FONTS = new Set(['方正小标宋gbk', 'fzxiaobiaosongb05']);
const HEI_FONTS = new Set(['黑体', 'simhei']);
const normalizeFont = name => (name || '').normalize('NFKC').toLowerCase().replace(/[@\s_\-]/g, '');
const isW = (node, name) => node?.nodeType === 1 && W.has(node.namespaceURI) && (!name || node.localName === name);
const kids = node => [...(node?.children || [])];
const child = (node, name) => kids(node).find(el => isW(el, name));
const attr = (node, name) => {
  if (!node) return '';
  for (const ns of W) if (node.hasAttributeNS(ns, name)) return node.getAttributeNS(ns, name);
  return '';
};
const descendants = (node, name) => [...(node?.getElementsByTagNameNS('*', name) || [])].filter(el => W.has(el.namespaceURI));
const on = node => !!node && !['0', 'false', 'off'].includes(attr(node, 'val'));

function xml(bytes, label) {
  if (!bytes) return null;
  const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le' : bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8';
  const source = new TextDecoder(encoding, { fatal: true }).decode(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error(`${label} 含不支持的 XML 实体声明。`);
  const doc = new DOMParser().parseFromString(source, 'application/xml');
  if (doc.getElementsByTagNameNS('*', 'parsererror').length) throw new Error(`${label} 的 XML 已损坏。`);
  return doc;
}

function openPackage(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (bytes.byteLength > MAX_FILE_BYTES) throw new Error('文件超过 30 MB，请拆分后导入。');
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf) throw new Error('这是旧版 DOC 或加密 Word 文件。请在 Word 中解密并另存为 .docx。');
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('不是有效的 DOCX 文件，请勿只修改文件扩展名。');
  let total = 0;
  // Inflates only the named parts; the decompressed-size budget spans every call for this file.
  return names => {
    let count = 0;
    try {
      return unzipSync(bytes, { filter(entry) {
        if (++count > 10000) throw new Error('文档包含过多内部文件。');
        if (!names.includes(entry.name)) return false;
        total += entry.originalSize;
        if (entry.originalSize > 12 * MiB || total > 32 * MiB) throw new Error('文档解压后的正文或样式过大，请拆分后导入。');
        return true;
      }});
    } catch (error) {
      if (/过大|过多/.test(error.message)) throw error;
      throw new Error('DOCX 压缩包损坏或格式不受支持，请重新保存后导入。');
    }
  };
}

// Internal relationship targets as package paths, keyed by the last segment of the type (officeDocument, styles, theme, …).
function relationshipTargets(relsDoc, sourcePart) {
  const targets = {};
  for (const rel of relsDoc?.documentElement.children || []) {
    const type = (rel.getAttribute('Type') || '').split('/').pop();
    const target = rel.getAttribute('Target');
    if (!type || !target || targets[type] || rel.getAttribute('TargetMode') === 'External') continue;
    const path = new URL(target, `https://local.invalid/${sourcePart}`).pathname.slice(1);
    try { targets[type] = decodeURIComponent(path); } catch { targets[type] = path; }
  }
  return targets;
}

function readRunProps(rPr) {
  const result = { fonts: {} };
  const fonts = child(rPr, 'rFonts');
  for (const slot of ['ascii', 'hAnsi', 'eastAsia', 'cs']) {
    const theme = attr(fonts, slot === 'cs' ? 'cstheme' : `${slot}Theme`);
    const face = attr(fonts, slot);
    if (theme || face) result.fonts[slot] = theme ? { theme } : { face };
  }
  const lang = child(rPr, 'lang');
  if (attr(lang, 'eastAsia')) result.lang = attr(lang, 'eastAsia');
  if (attr(fonts, 'hint')) result.hint = attr(fonts, 'hint');
  for (const name of ['vanish', 'webHidden', 'cs', 'rtl']) {
    const el = child(rPr, name);
    if (el) result[name] = on(el);
  }
  return result;
}

function overlay(...layers) {
  const result = { fonts: {} };
  for (const layer of layers) if (layer) {
    const previous = result.fonts;
    Object.assign(result, layer);
    result.fonts = { ...previous, ...layer.fonts };
  }
  return result;
}

function makeStyles(stylesDoc, warn) {
  const styles = new Map();
  let defaultParagraph = '', defaultCharacter = '';
  for (const el of descendants(stylesDoc, 'style')) {
    const id = attr(el, 'styleId');
    const props = readRunProps(child(el, 'rPr'));
    const numbering = child(child(el, 'pPr'), 'numPr');
    if (numbering) props.numbered = attr(child(numbering, 'numId'), 'val') !== '0';
    styles.set(id, { basedOn: attr(child(el, 'basedOn'), 'val'), props });
    if (onDefault(el)) {
      if (attr(el, 'type') === 'paragraph') defaultParagraph = id;
      if (attr(el, 'type') === 'character') defaultCharacter = id;
    }
  }
  const defaults = readRunProps(child(child(descendants(stylesDoc, 'docDefaults')[0], 'rPrDefault'), 'rPr'));
  const cache = new Map();
  const resolve = (id, visited = new Set()) => {
    if (!id) return { fonts: {} };
    if (cache.has(id)) return cache.get(id);
    if (visited.has(id) || visited.size > 80) { warn('样式继承存在循环或层级过深，部分字体无法确定。'); return { fonts: {} }; }
    const style = styles.get(id);
    if (!style) { warn(`未找到样式 ${id}，部分字体可能无法确定。`); return { fonts: {} }; }
    visited.add(id);
    const resolved = overlay(resolve(style.basedOn, visited), style.props);
    cache.set(id, resolved);
    return resolved;
  };
  return { defaults, defaultParagraph, defaultCharacter, resolve };
}
function onDefault(el) { return ['1', 'true', 'on'].includes(attr(el, 'default')); }

function themeFont(themeDoc, key, lang) {
  const group = key?.startsWith('major') ? 'majorFont' : 'minorFont';
  const font = themeDoc?.getElementsByTagNameNS('*', group)[0];
  if (!font) return '';
  const type = key?.endsWith('EastAsia') ? 'ea' : key?.endsWith('Bidi') ? 'cs' : 'latin';
  const direct = kids(font).find(el => el.localName === type)?.getAttribute('typeface');
  if (direct) return direct;
  if (type !== 'ea') return '';
  const script = /^ja/i.test(lang) ? 'Jpan' : /^ko/i.test(lang) ? 'Hang' : /(?:TW|HK|MO|Hant)/i.test(lang) ? 'Hant' : 'Hans';
  return kids(font).find(el => el.localName === 'font' && el.getAttribute('script') === script)?.getAttribute('typeface') || '';
}

// Unicode block → font slot as Word applies it ([MS-OI29500] 2.1.88, ISO/IEC 29500-1 §17.3.2.26); unlisted code points use hAnsi.
// "hint" means eastAsia only under w:hint="eastAsia"; "hintZh" additionally requires a Chinese run language.
const SLOT_BLOCKS = [
  [0x0000, 0x007f, 'ascii'], [0x00a0, 0x00ff, 'latin1'], [0x0100, 0x02af, 'hintZh'], [0x02b0, 0x03cf, 'hint'],
  [0x0400, 0x04ff, 'hint'], [0x0590, 0x07bf, 'ascii'], [0x1100, 0x11ff, 'eastAsia'], [0x1e00, 0x1eff, 'hintZh'],
  [0x2000, 0x27bf, 'hint'], [0x2e80, 0x2eff, 'hint'], [0x2f00, 0x2fdf, 'eastAsia'], [0x2ff0, 0x319f, 'eastAsia'],
  [0x3200, 0x4dbf, 'eastAsia'], [0x4e00, 0x9faf, 'eastAsia'], [0xa000, 0xa4cf, 'eastAsia'], [0xac00, 0xd7af, 'eastAsia'],
  [0xd800, 0xdfff, 'eastAsia'], [0xe000, 0xf8ff, 'hint'], [0xf900, 0xfaff, 'eastAsia'], [0xfb00, 0xfb1c, 'hint'],
  [0xfb1d, 0xfdff, 'ascii'], [0xfe30, 0xfe6f, 'eastAsia'], [0xfe70, 0xfefe, 'ascii'], [0xff00, 0xffef, 'eastAsia']
];
const LATIN1_HINT = '¡¤§¨ª\u00ad¯°±²³´¶·¸¹º¼½¾¿×÷';
const LATIN1_HINT_ZH = 'àáèéêìíòóùúü';

function charSlot(char, hint, zh) {
  const code = char.codePointAt(0);
  // Supplementary-plane characters are stored as UTF-16 surrogates, which the table maps to eastAsia.
  const rule = code > 0xffff ? 'eastAsia' : SLOT_BLOCKS.find(([start, end]) => code >= start && code <= end)?.[2] || 'hAnsi';
  if (rule === 'latin1') return hint && (LATIN1_HINT.includes(char) || zh && LATIN1_HINT_ZH.includes(char)) ? 'eastAsia' : 'hAnsi';
  if (rule === 'hint') return hint ? 'eastAsia' : 'hAnsi';
  if (rule === 'hintZh') return hint && zh ? 'eastAsia' : 'hAnsi';
  return rule;
}

const sameSpec = (a, b) => !!a && !!b && a.theme === b.theme && normalizeFont(a.face) === normalizeFont(b.face);

// Every character of a run shares one props object, so its slot rules and fonts are resolved once per run instead of once per character.
function fontResolver(theme, themeLang, aliases) {
  const runs = new WeakMap();
  const describe = props => {
    const { ascii, hAnsi, eastAsia } = props.fonts;
    const lang = props.lang || themeLang;
    // Complex-script runs, and a "Times New Roman" eastAsia face with equal ascii/hAnsi, bypass the Unicode table.
    const fixed = props.cs || props.rtl ? 'cs' : normalizeFont(eastAsia?.face) === 'timesnewroman' && sameSpec(ascii, hAnsi) ? 'ascii' : '';
    return { lang, fixed, hint: props.hint === 'eastAsia', zh: /^zh/i.test(lang), fonts: {} };
  };
  return (char, props) => {
    let run = runs.get(props);
    if (!run) runs.set(props, run = describe(props));
    const slot = run.fixed || charSlot(char, run.hint, run.zh);
    if (!(slot in run.fonts)) {
      const spec = props.fonts[slot];
      const name = spec ? normalizeFont(spec.face || themeFont(theme, spec.theme, run.lang)) : '';
      run.fonts[slot] = aliases.get(name) || name;
    }
    return run.fonts[slot];
  };
}

function childElementsVisible(node) {
  if (node.localName === 'AlternateContent') {
    const children = kids(node);
    const chosen = children.find(el => el.localName === 'Choice') || children.find(el => el.localName === 'Fallback');
    return chosen ? kids(chosen) : [];
  }
  return kids(node);
}

export function escapeMarkdown(text) {
  return text.replace(/\\/g, '\\\\').replace(/([`*_{}\[\]<>&#!|~+\-=])/g, '\\$1').replace(/^([ \t]*\d+)([.)])(?=\s)/gm, '$1\\$2');
}

const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Bopomofo}\u3000-\u303f\uff00-\uffef]/u;
// Collapses heading whitespace; a line break next to a CJK character is dropped, because Chinese text takes no space.
export function headingText(text) {
  return text.trim().replace(/\s*\n\s*/g, (gap, offset, whole) => {
    const before = whole.slice(0, offset).match(/.$/u)?.[0] || '', after = whole.slice(offset + gap.length).match(/^./u)?.[0] || '';
    return CJK.test(before) || CJK.test(after) ? '' : ' ';
  }).replace(/\s+/g, ' ');
}

export function documentToMarkdown(doc) {
  const lines = [`# ${escapeMarkdown(headingText(doc.title))}`];
  for (const [index, p] of doc.paragraphs.entries()) {
    if (index === doc.titleIndex) continue;
    lines.push(p.isHeading ? `## ${escapeMarkdown(headingText(p.text))}` : escapeMarkdown(p.text).replace(/\n/g, '  \n'));
  }
  return `${lines.join('\n\n')}\n`;
}

export function mergeDocuments(docs) {
  return docs.length ? `${docs.map(doc => documentToMarkdown(doc).trimEnd()).join('\n\n---\n\n')}\n` : '';
}

/** Parse visible main-body paragraphs. Original page layout is intentionally not reproduced. */
export function parseDocx(input, filename = '未命名.docx') {
  if (/\.doc$/i.test(filename)) throw new Error('暂不支持旧版 .doc。请在 Word 或 WPS 中另存为 .docx 后导入。');
  if (!/\.docx$/i.test(filename)) throw new Error('请选择 .docx 格式的 Word 文档。');
  // Word for the web saves the main part as word/document2.xml, so follow the package relationships instead of fixed names.
  const extract = openPackage(input);
  const packageRels = relationshipTargets(xml(extract(['_rels/.rels'])['_rels/.rels'], '关系'), '');
  const mainPath = packageRels.officeDocument || 'word/document.xml';
  const mainRelsPath = mainPath.replace(/[^/]*$/, name => `_rels/${name}.rels`);
  const main = extract([mainPath, mainRelsPath]);
  if (!main[mainPath]) throw new Error('压缩包中没有 Word 正文，或使用了不支持的文档部件路径。');
  const document = xml(main[mainPath], '正文');
  const body = child(document.documentElement, 'body');
  if (!body) throw new Error('未找到 Word 正文。');
  const warnings = new Set();
  const warn = text => warnings.add(text);
  const related = relationshipTargets(xml(main[mainRelsPath], '关系'), mainPath);
  const folder = mainPath.replace(/[^/]*$/, '');
  const paths = { styles: 'styles.xml', settings: 'settings.xml', fontTable: 'fontTable.xml', theme: 'theme/theme1.xml' };
  for (const [type, fallback] of Object.entries(paths)) paths[type] = related[type] || folder + fallback;
  const parts = extract(Object.values(paths));
  const styles = makeStyles(xml(parts[paths.styles], '样式'), warn);
  const settings = xml(parts[paths.settings], '设置');
  const themeLang = attr(descendants(settings, 'themeFontLang')[0], 'eastAsia') || 'zh-CN';
  const theme = xml(parts[paths.theme], '主题');
  const fontTable = xml(parts[paths.fontTable], '字体表');
  const aliases = new Map();
  for (const font of descendants(fontTable, 'font')) {
    const name = normalizeFont(attr(font, 'name'));
    const alternate = normalizeFont(attr(child(font, 'altName'), 'val'));
    if (TITLE_FONTS.has(alternate) || HEI_FONTS.has(alternate)) aliases.set(name, alternate);
  }
  const fontOf = fontResolver(theme, themeLang, aliases);
  const paragraphs = [];
  let unknownFonts = false, hiddenText = false, deletedText = false;
  const fieldStack = [];

  function readParagraph(p, inTable) {
    const pPr = child(p, 'pPr');
    const base = overlay(styles.defaults, styles.resolve(attr(child(pPr, 'pStyle'), 'val') || styles.defaultParagraph));
    const segments = [];
    function visit(node, inherited) {
      if (isW(node, 'del') || isW(node, 'moveFrom')) { deletedText = true; return; }
      if (isW(node, 'pPr') || isW(node, 'rPr')) return;
      if (isW(node, 'p') && node !== p) return;
      if (isW(node, 'drawing') || isW(node, 'pict') || isW(node, 'object')) {
        warn('图片、文本框和嵌入对象未提取；不包含 OCR 识别。'); return;
      }
      if (isW(node, 'r')) {
        const rPr = child(node, 'rPr');
        const props = overlay(base, styles.resolve(attr(child(rPr, 'rStyle'), 'val') || styles.defaultCharacter), readRunProps(rPr));
        if (props.vanish || props.webHidden) { hiddenText = true; return; }
        for (const next of childElementsVisible(node)) visit(next, props);
        return;
      }
      if (isW(node, 'fldChar')) {
        const type = attr(node, 'fldCharType');
        if (type === 'begin') fieldStack.push(false);
        else if (type === 'separate' && fieldStack.length) fieldStack[fieldStack.length - 1] = true;
        else if (type === 'end') fieldStack.pop();
        return;
      }
      if (fieldStack.some(show => !show)) return;
      if (isW(node, 't')) { segments.push({ text: node.textContent || '', props: inherited || base }); return; }
      if (isW(node, 'tab')) { segments.push({ text: '\t', props: inherited || base }); return; }
      if (isW(node, 'br') || isW(node, 'cr')) { segments.push({ text: '\n', props: inherited || base }); return; }
      if (isW(node, 'noBreakHyphen')) { segments.push({ text: '‑', props: inherited || base }); return; }
      if (isW(node, 'softHyphen')) { segments.push({ text: '\u00ad', props: inherited || base }); return; }
      if (isW(node, 'sym')) { warn('文档含符号字体字符，部分符号未提取。'); return; }
      for (const next of childElementsVisible(node)) visit(next, inherited);
    }
    visit(p, base);
    const text = segments.map(part => part.text).join('').replace(/\r/g, '').trim();
    if (!text) return;
    let hasTitleFont = false, allHei = true, visible = 0;
    const fonts = new Set();
    for (const segment of segments) for (const char of segment.text) {
      if (/[\s\u200b\u200c\u200d\ufeff]/u.test(char)) continue;
      visible++;
      const font = fontOf(char, segment.props);
      if (!font) unknownFonts = true;
      fonts.add(font || '未知');
      hasTitleFont ||= TITLE_FONTS.has(font);
      allHei &&= HEI_FONTS.has(font);
    }
    const directNum = child(pPr, 'numPr');
    if (directNum ? attr(child(directNum, 'numId'), 'val') !== '0' : base.numbered) warn('自动列表编号未还原，仅保留段落文字；手动输入的编号会保留。');
    paragraphs.push({ text, hasTitleFont, isHeading: !inTable && visible > 0 && allHei, fonts: [...fonts] });
  }
  // Table header cells are often set in 黑体, so paragraphs inside tables never become headings.
  function walk(node, inTable = false) {
    if (isW(node, 'del') || isW(node, 'moveFrom')) { deletedText = true; return; }
    if (isW(node, 'p')) { readParagraph(node, inTable); return; }
    if (isW(node, 'tbl')) { warn('表格已按行、单元格顺序展开为段落，不保留表格布局及表格样式字体；表格内的黑体段落不作为二级标题。'); inTable = true; }
    if (isW(node, 'altChunk')) { warn('文档含外部插入内容（altChunk），该部分未提取。'); return; }
    for (const next of childElementsVisible(node)) walk(next, inTable);
  }
  walk(body);
  if (!paragraphs.length) throw new Error('文档中没有可提取的正文文字，可能是扫描件或仅包含图片。');
  const titleIndex = paragraphs.findIndex(p => p.hasTitleFont);
  const title = titleIndex >= 0 ? headingText(paragraphs[titleIndex].text) : filename.replace(/\.docx$/i, '');
  if (titleIndex < 0) warn('未找到方正小标宋_GBK 段落，已使用文件名作为一级标题。');
  if (unknownFonts) warn('部分文字未声明可解析的字体，已保留文字；未知字体不用于识别标题。');
  if (hiddenText || deletedText) warn('已忽略隐藏文字和删除的修订，保留插入的修订。');
  if (descendants(document, 'headerReference').length || descendants(document, 'footerReference').length) warn('页眉和页脚未提取。');
  if (descendants(document, 'footnoteReference').length || descendants(document, 'endnoteReference').length) warn('脚注和尾注未提取。');
  if (descendants(document, 'oMath').length || document.getElementsByTagNameNS('*', 'oMath').length) warn('数学公式未转换，可能缺少公式内容。');
  return { filename, title, titleIndex, titleSource: titleIndex < 0 ? 'filename' : 'font', paragraphs, headings: paragraphs.filter((p, i) => p.isHeading && i !== titleIndex).length, warnings: [...warnings] };
}
