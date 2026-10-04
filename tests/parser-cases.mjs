import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { readFile } from 'node:fs/promises';

export async function parserCases(root) {
  const base = unzipSync(await readFile(new URL('demos/01_直接字体.docx', root)));
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = (text, props = '') => `<w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const f = name => `<w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:eastAsia="${name}" w:cs="${name}"/>`;
  const p = runs => `<w:p>${runs}</w:p>`;
  const styles = `<w:styles xmlns:w="${W}"/>`;
  function pack(body, overrides = {}) {
    const parts = { ...base, 'word/document.xml': strToU8(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`), 'word/styles.xml': strToU8(styles), ...overrides };
    return Buffer.from(zipSync(parts)).toString('base64');
  }
  function body(text, face = '宋体') { return p(r(text, f(face))); }
  const cases = [
    { name: '完整标题来自首次匹配段落', data: pack(p(r('完整', f('宋体')) + r('标题', f('方正小标宋_gbk')) + r('后缀', f('宋体')))), title: '完整标题后缀', headings: 0 },
    { name: '多 run 黑体与空白', data: pack(p(r('甲', f('黑体')) + r('  ', f('宋体')) + r('乙', f('SimHei')))), headings: 1 },
    { name: '部分黑体不升级', data: pack(p(r('甲', f('黑体')) + r('乙', f('宋体')))), headings: 0 },
    { name: '加粗不等于黑体', data: pack(p(r('粗体', f('宋体') + '<w:b/>'))), headings: 0 },
    { name: '中文和 ASCII 分别识别字体', data: pack(p(r('中文1', '<w:rFonts w:ascii="Calibri" w:eastAsia="黑体"/>'))), headings: 0 },
    { name: 'document defaults 生效', data: pack(p(r('默认字体')), { 'word/styles.xml': strToU8(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>${f('黑体')}</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`) }), headings: 1 },
    { name: '未知字体不会猜成标题', data: pack(p(r('没有字体'))), headings: 0, warning: '未声明' },
    { name: '字符样式作为一级标题', data: pack(p(r('字符样式标题', '<w:rStyle w:val="TitleChild"/>')), { 'word/styles.xml': strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="character" w:styleId="TitleBase"><w:rPr>${f('方正小标宋_GBK')}</w:rPr></w:style><w:style w:type="character" w:styleId="TitleChild"><w:basedOn w:val="TitleBase"/></w:style></w:styles>`) }), title: '字符样式标题', headings: 0 },
    { name: '循环样式不会无限递归', data: pack(`<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr>${r('循环')}</w:p>`), title: 'test', warning: '未找到样式' },
    { name: '空白段落不成为标题', data: pack(body('  ', '方正小标宋_GBK') + body('正文')), title: 'test', headings: 0 },
    { name: '隐藏字与删除修订排除', data: pack(p(r('隐藏', f('方正小标宋_GBK') + '<w:vanish/>')) + `<w:del>${body('删除', '方正小标宋_GBK')}</w:del>` + body('保留')), title: 'test', text: ['保留'] },
    { name: '复杂域保留结果不保留代码', data: pack(`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>HYPERLINK secret</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r('结果', f('宋体'))}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`), text: ['结果'] },
    { name: '替代内容只读取一个分支', data: pack(`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="w">${body('首选')}</mc:Choice><mc:Fallback>${body('重复的回退')}</mc:Fallback></mc:AlternateContent>`), text: ['首选'] },
    { name: '字体别名表支持', data: pack(body('别名标题', 'CustomSmallSong'), { 'word/fontTable.xml': strToU8(`<w:fonts xmlns:w="${W}"><w:font w:name="CustomSmallSong"><w:altName w:val="方正小标宋_GBK"/></w:font></w:fonts>`) }), title: '别名标题' },
    { name: 'XML 实体声明拒绝', data: pack(body('正文'), { 'word/document.xml': strToU8(`<!DOCTYPE w:document [<!ENTITY a "abc">]><w:document xmlns:w="${W}"><w:body/></w:document>`) }), error: '实体' },
    { name: 'XML 损坏拒绝', data: pack('', { 'word/document.xml': strToU8('<w:document>') }), error: 'XML' },
    { name: '纯图片或空正文给出提示', data: pack(''), error: '没有可提取' },
    { name: '非 Word ZIP 拒绝', data: Buffer.from(zipSync({ 'readme.txt': strToU8('not word') })).toString('base64'), error: '没有 Word 正文' },
    { name: '伪造扩展名拒绝', data: Buffer.from('plain text').toString('base64'), error: '不是有效' },
    { name: '截断 ZIP 拒绝', data: Buffer.from([0x50, 0x4b, 0x03, 0x04]).toString('base64'), error: '压缩包' },
    { name: 'OLE 加密容器拒绝', data: Buffer.from([0xd0, 0xcf, 0x11, 0xe0]).toString('base64'), error: '加密' },
    { name: '旧版 DOC 明确拒绝', data: Buffer.from([0xd0, 0xcf]).toString('base64'), filename: 'old.doc', error: '旧版' },
    { name: '非 Word 扩展名拒绝', data: pack(body('正文')), filename: 'data.pdf', error: '.docx' },
    { name: '解压正文超限拒绝', data: pack('', { 'word/document.xml': strToU8('x'.repeat(12 * 1024 * 1024 + 1)) }), error: '解压' },
    { name: '标题字体反向验证', data: pack(body('看起来像标题但字体是宋体', '宋体')), title: 'test', headings: 0 },
    { name: 'Markdown 特殊字符保持字面含义', data: pack(body('---') + body('===') + body('1. 正文') + body('## 伪标题') + body('&lt;img src=x onerror=alert(1)&gt;') + body('A &amp;amp; B')), markdownIncludes: ['\\-\\-\\-', '\\=\\=\\=', '1\\. 正文', '\\#\\# 伪标题', '\\<img src\\=x onerror\\=alert(1)\\>', 'A \\&amp; B'] }
  ];
  // Real cyclical inheritance, distinct from the missing-style case.
  cases.push({ name: '样式继承循环有界', data: pack(`<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr>${r('循环')}</w:p>`, { 'word/styles.xml': strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style><w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style></w:styles>`) }), warning: '循环' });
  const strict = pack(body('严格格式', '方正小标宋_GBK'));
  const strictParts = unzipSync(Buffer.from(strict, 'base64'));
  for (const key of ['word/document.xml', 'word/styles.xml']) strictParts[key] = strToU8(strFromU8(strictParts[key]).replaceAll(W, 'http://purl.oclc.org/ooxml/wordprocessingml/main'));
  cases.push({ name: 'Strict OOXML 命名空间', data: Buffer.from(zipSync(strictParts)).toString('base64'), title: '严格格式' });
  const themeBody = p(r('主题中文', '<w:rFonts w:eastAsiaTheme="majorEastAsia"/>'));
  cases.push({ name: '主题字体 Hans 补充映射', data: pack(themeBody, { 'word/settings.xml': strToU8(`<w:settings xmlns:w="${W}"><w:themeFontLang w:eastAsia="zh-CN"/></w:settings>`), 'word/theme/theme1.xml': strToU8('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Demo"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hans" typeface="方正小标宋_GBK"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>') }), title: '主题中文' });
  // Font slots follow the Unicode table Word uses ([MS-OI29500] 2.1.88): CJK text in 黑体, Latin text in Times New Roman.
  const mixed = (hint, lang = '') => `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"${hint ? ' w:hint="eastAsia"' : ''}/>${lang ? `<w:lang w:eastAsia="${lang}"/>` : ''}`;
  const titled = paragraphs => pack(body('标题', '方正小标宋_GBK') + paragraphs);
  cases.push(
    { name: 'hint=eastAsia 时 ①、→、· 使用中文字体', data: titled(p(r('①工作目标', mixed(true))) + p(r('→工作目标', mixed(true))) + p(r('张三·李四', mixed(true)))), headings: 3 },
    { name: '没有 hint 时 ① 使用西文字体', data: titled(p(r('①工作目标', mixed(false)))), headings: 0 },
    { name: '㈠ 等中日韩带圈字符总用中文字体', data: titled(p(r('㈠工作目标', mixed(false)))), headings: 1 },
    { name: '拼音声调字母仅在中文 run 中随中文字体', data: titled(p(r('注音ā', mixed(true, 'zh-CN'))) + p(r('注音ā', mixed(true, 'ja-JP')))), headings: 1 },
    { name: 'eastAsia 为 Times New Roman 且西文字体相同时整段用 ascii', data: titled(p(r('工作目标', '<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="Times New Roman"/>'))), headings: 1 },
    { name: 'eastAsia 为 Times New Roman 但西文字体不同时仍按表', data: titled(p(r('工作目标', '<w:rFonts w:ascii="黑体" w:hAnsi="Arial" w:eastAsia="Times New Roman"/>'))), headings: 0 }
  );
  const row = (cells, face) => `<w:tr>${cells.map(text => `<w:tc>${body(text, face)}</w:tc>`).join('')}</w:tr>`;
  // Word for the web stores the main part as word/document2.xml; every part must be found through relationships.
  const relsXml = list => strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map(([type, target], i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join('')}</Relationships>`);
  const webParts = { ...base };
  for (const name of ['word/document.xml', 'word/styles.xml', 'word/_rels/document.xml.rels']) delete webParts[name];
  Object.assign(webParts, {
    '_rels/.rels': relsXml([['officeDocument', 'word/document2.xml']]),
    'word/_rels/document2.xml.rels': relsXml([['styles', 'styles2.xml'], ['theme', 'theme/theme1.xml']]),
    'word/document2.xml': strToU8(`<w:document xmlns:w="${W}"><w:body>${body('网页版标题', '方正小标宋_GBK')}<w:p><w:pPr><w:pStyle w:val="Hei"/></w:pPr>${r('一、工作目标')}</w:p></w:body></w:document>`),
    'word/styles2.xml': strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Hei"><w:rPr>${f('黑体')}</w:rPr></w:style></w:styles>`)
  });
  cases.push({ name: '按关系文件定位 document2.xml 及其样式部件', data: Buffer.from(zipSync(webParts)).toString('base64'), title: '网页版标题', headings: 1 });
  // OPC keeps part names percent-encoded in ZIP item names, so relationship targets must not be decoded.
  const encodedParts = { ...webParts, 'word/_rels/document2.xml.rels': relsXml([['styles', 'styles%202.xml']]), 'word/styles%202.xml': webParts['word/styles2.xml'] };
  delete encodedParts['word/styles2.xml'];
  cases.push({ name: '关系目标保持百分号编码匹配 ZIP 条目名', data: Buffer.from(zipSync(encodedParts)).toString('base64'), title: '网页版标题', headings: 1 });
  const br = face => `<w:r><w:rPr>${f(face)}</w:rPr><w:br/></w:r>`;
  cases.push(
    { name: '标题与二级标题内换行在中文之间直接拼接', data: pack(p(r('关于推进文档数字化', f('方正小标宋_GBK')) + br('方正小标宋_GBK') + r('整理工作的通知', f('方正小标宋_GBK'))) + p(r('一、工作', f('黑体')) + br('黑体') + r('目标', f('黑体')))), title: '关于推进文档数字化整理工作的通知', headings: 1, markdownIncludes: ['# 关于推进文档数字化整理工作的通知\n', '## 一、工作目标\n'] },
    { name: '西文标题换行合并为一个空格', data: pack(p(r('Annual', f('方正小标宋_GBK')) + br('方正小标宋_GBK') + r('Report', f('方正小标宋_GBK')))), title: 'Annual Report' },
    { name: '方正小标宋_GBK 英文名 FZXiaoBiaoSong-B05', data: pack(body('英文名标题', 'FZXiaoBiaoSong-B05')), title: '英文名标题' },
    { name: '方正小标宋简体及其英文名不是标题字体', data: pack(body('简体中文名', '方正小标宋简体') + body('简体英文名', 'FZXiaoBiaoSong-B05S')), title: 'test', warning: '未找到方正小标宋' }
  );
  cases.push({ name: '表格内黑体单元格不作为二级标题', data: titled(body('一、工作安排', '黑体') + `<w:tbl>${row(['序号', '任务'], '黑体')}${row(['1', '收集'])}</w:tbl>`), headings: 1, text: ['标题', '一、工作安排', '序号', '任务', '1', '收集'] });
  // Word's default Normal style resolves through theme fonts; 20,000 such paragraphs took 3-4 s before per-run caching.
  const themeDefaults = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;
  const longBody = body('长文档标题', '方正小标宋_GBK') + Array.from({ length: 20000 }, (_, i) => p(r(`第${i + 1}段：为落实文档数字化整理要求，各部门应按期完成材料收集、归档与核对工作。`))).join('');
  cases.push({ name: '主题字体长文档解析耗时有上限', data: pack(longBody, { 'word/styles.xml': strToU8(themeDefaults) }), title: '长文档标题', maxMs: 1500 });
  return cases;
}
