# Words to MD

纯前端的 Word → Markdown 批量合并工具。解析和导出都在浏览器本地完成，没有服务端、账号、分析埋点或运行时 CDN 依赖。

[单文件 HTML](dist/word-to-markdown.html) · [示例 Word](demos/) · [示例合并结果](tests/expected.md)

![文档归并界面](docs/preview.png)

## 直接使用

1. 下载 `dist/word-to-markdown.html`，用浏览器打开，无需安装或启动服务器。在 GitHub 文件页点击 “Download raw file” 即可下载。
2. 选择或拖入多份 `.docx`，按 ↑ / ↓ 调整合并顺序。点击“试用示例文档”可直接导入内置的四份真实 Word 文件。
3. 查看逐文件提示及结果，点击“复制”“下载 .md”或“新标签页”。下载得到 UTF-8 编码的 Markdown 文件。

新标签页的内容类型是 `text/plain`。使用 Ctrl / ⌘ + S 可以另存为本地文件，浏览器可能默认给出 `.txt` 后缀，需要手动改成 `.md`。直接使用“下载 .md”可获得正确文件名。页面刷新或关闭会清空导入内容。

## 标题规则

- 每份 Word 文档恰好一个一级标题：取正文中第一个含“方正小标宋_GBK”可见文字的完整段落，移到该文档开头，不重复输出。该段落之前的正文仍按原顺序保留。标题和二级标题内的换行，与中文相邻时直接拼接，其余合并为一个空格。
- 支持字体名称大小写、空格、下划线差异及英文名 `FZXiaoBiaoSong-B05`。方正小标宋简体（英文名 `FZXiaoBiaoSong-B05S`）是另一款 GB2312 字体，不作为标题字体。字体判断读取文件中的元数据，无需本机安装该字体。
- 其他段落的所有非空白文字，包括数字和标点，都使用“黑体 / SimHei”时，输出二级标题。宋体加粗、黑体宋体混排、中文黑体但数字为 Calibri 的段落均保留为正文。表格单元格内的段落不作为二级标题，因为表头常用黑体。
- 解析直接字体、文档默认字体、段落样式及 basedOn 继承、字符样式及继承、常见中西文主题字体、字体表 alternate name。支持 Transitional 和 Strict WordprocessingML 命名空间。
- 找不到目标字体时使用文件名作为一级标题，并提示用户。无法确定的字体不会猜测成目标字体。

## 内容与边界

- 支持 `.docx`；旧版二进制 `.doc`、加密 Word、其他扩展名会被明确拒绝。请先在 Word / WPS 中另存为未加密 `.docx`。
- 正文、样式、设置、字体表和主题均按 `_rels/.rels` 及正文的关系文件定位，兼容 Word 网页版保存的 `word/document2.xml` 等非默认部件名。
- 提取正文段落、超链接显示文字、软换行、制表符，以及按行和单元格顺序展开的表格文字。Markdown 特殊符号会转义，防止正文被误读成标题、HTML 或列表。
- 不保留原始分页、版式、表格结构、表格条件样式字体、自动列表编号、图片、文本框、嵌入对象、页眉页脚、脚注尾注或公式。检测到相应结构时会提供提示。没有 OCR。
- 按“接受修订后的正文”处理：保留插入、排除删除及 moveFrom；隐藏文字不会参与标题识别。不会修改原 Word 文件。
- 每个字符使用西文、中文还是复杂文种字体，按 Word 的 Unicode 区块表判断（[MS-OI29500] 2.1.88），包括 `w:hint="eastAsia"` 下的 ①、→、· 等符号，以及 Times New Roman 中文字体占位规则；表中“字体字符集”条件以 run 的东亚语言近似。这不等同于 Word 的完整排版引擎，特殊字体替代或第三方软件生成的非常规 OOXML，请核对原文。
- 单个压缩文件上限 30 MB，正文 / 单个 XML 上限 12 MB，选中 XML 解压总量上限 32 MB，单次会话最多 100 份文档。不解压图片等无关二进制资源。
- 错误文件不会阻止本批其余文档。名称、大小及修改时间均相同的已导入文件会跳过，并明确提示；这不是基于内容哈希的去重。

## 实现与开发

使用原生 HTML / CSS / JavaScript、`fflate 0.8.2` 和浏览器 `DOMParser`。只增加一个运行时依赖用于 ZIP 解压与 demo 打包；字体识别直接读取 OOXML，避免普通文本转换丢失字体信息。`esbuild 0.25.12` 仅用于构建单文件 HTML。

```sh
npm ci
npm run build
# 可选，仅提供静态本地预览；转换不使用此服务
npm run serve
```

`serve` 只监听 `127.0.0.1:4173`。生产使用只需复制构建出的 HTML 文件。页面内置 CSP，禁止网络连接、外部脚本、字体及对象；预览使用 `textContent` 创建 DOM，导入文字不会作为 HTML 执行。文件不会写入 localStorage 或 IndexedDB。

浏览器提供一个可复用的小型 API：

```js
const bytes = new Uint8Array(await selectedFile.arrayBuffer());
const parsed = WordMD.parseDocx(bytes, selectedFile.name);
// parsed: { filename, title, titleIndex, titleSource, paragraphs, headings, warnings }，title 已合并换行
// paragraphs: [{ text, hasTitleFont, isHeading, fonts }]
// 失败时抛出带中文说明的 Error；没有匹配标题时返回 filename 兜底及 warnings。
const markdown = WordMD.mergeDocuments([parsed]);
```

主要文件：

- `src/parser.js`：读取压缩包、样式及主题，解析有效字体，生成 Markdown。
- `src/app.js`：批量导入、顺序管理、预览及三种导出方式。
- `src/index.html`、`src/style.css`：交互结构与响应式布局。
- `src/demos.json`：四份 demo 的原始 DOCX 字节，支持离线试用及下载。
- `scripts/generate_demos.py`：使用 python-docx 生成真实 Word 示例。开发者如需重生成，安装 `python-docx>=1.2`，运行此脚本后重新 build；日常使用不需要 Python。

## 验证

`tests/acceptance.md` 在实现前定义验收行为；`tests/expected.md` 是人工写定的合并期望。四份真实 Word 覆盖直接字体、段落及字符样式继承、主题字体、缺失标题、混合字体、修订和表格文字。

本次验证在 macOS 的 ego-lite Chromium 中进行，使用 `file://` 打开，验证期间浏览器断网。测试包括字面文本安全、样式循环、无效 / 加密 / 超限文件、Word 字体槽表（①、→、㈠、拼音字母、Times New Roman 占位）、表格内黑体单元格、按关系文件定位的 `document2.xml`、标题内换行、主题字体长文档解析耗时上限（2 万段 1.5 s）、真实文件选择、暂停 `requestAnimationFrame` 时的批量导入、排序删除、混合失败批次、内置 demo、真实下载、剪贴板粘贴核对、新标签页纯文本、键盘切换和 390px 窄屏。结果及文件 SHA-256 保存在 `verification/results.json`。

安装并启动 ego-browser 后，可重复运行：

```sh
npm run build
npm test
# 已有当前任务独占的隔离空间时，可复用：
# EGO_SPACE_ID=你的空间编号 npm test
```

测试生成 `test-artifacts/results.json`、下载文件与截图；失败时保留 `failure.png`。无需浏览器工具也可手工将 `demos/` 的四份 Word 按文件名顺序导入，与 `tests/expected.md` 比较。

四份 demo 已经过 LibreOffice 渲染及逐页检查；此后 02 号只把样式字体名由 `FZXiaoBiaoSong-B05S` 改为 `FZXiaoBiaoSong-B05`，未重新渲染。渲染验证在临时 fontconfig 中使用本机宋体 / 黑体替代字形；没有改写 DOCX 的字体名称，也没有打包商业字体。尚未使用用户的实际 Word 样本，Safari / Firefox 及原生移动浏览器未实测。

技术依据：[Microsoft Open XML RunFonts](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts?view=openxml-3.0.1)、[MS-OI29500 2.1.88 rFonts](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/aef3c9a6-5d6c-434b-90b7-85e761fd8e62)、[MDN Blob URL](https://developer.mozilla.org/en-US/docs/Web/API/URL/createObjectURL_static)。
