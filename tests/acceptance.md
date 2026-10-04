# 验收场景（实现前定义）

1. Given 一篇含前言、两次小标宋和黑体段落的文档，When 导入，Then 首次小标宋所在完整段落成为唯一 H1，标题从原位置移除，前言仍在正文，后续小标宋不升级。
2. Given 一个段落分成多个黑体 run，包含空白 run，When 解析，Then 输出 H2。Given 黑体与宋体混排，或宋体仅加粗，Then 保留正文。
3. Given 字体来自段落样式、basedOn 父样式、字符样式或主题，When 没有 run 直接字体，Then 按有效字体识别；直接字体覆盖继承值。
4. Given 中西文字体不同，When 汉字是黑体但 ASCII 字符不是，Then 段落不是全黑体。未知字体不猜测为标题。
5. Given 缺失小标宋，When 解析，Then 用不含扩展名的文件名做 H1 并显示警告，不丢正文。
6. Given 表格、换行、制表符、链接文本、已删除修订、隐藏文字，When 解析，Then 保留可见正文和表格单元格阅读顺序，忽略删除与隐藏文字，提示有损范围。
7. Given 两篇以上文档，When 导入、调整顺序、删除，Then 合并输出与可见顺序一致，失败文件不会阻止其他文件成功。
8. Given 正文含 #、星号、反引号、HTML 标签，When 预览或导出，Then 不会产生伪标题或执行 HTML，文字内容保留。
9. Given 旧版 DOC、损坏 ZIP、无 Word 正文的 ZIP、加密容器、超限文件，When 导入，Then 给出可理解错误，不输出伪造正文。
10. Given 已合并内容，When 下载、复制或在新标签页打开，Then UTF-8 Markdown 与页面内容一致，下载扩展名为 .md；新页是纯文本。
11. Given 单文件 HTML 在 file:// 或断网环境，When 载入 demo，Then 解析不需要网络和后端。
12. Given 窄屏和键盘操作，When 添加文件、切换预览、调整顺序，Then 关键操作可达，没有页面横向溢出。

## 补充场景（2026-10-04）

13. Given 表格表头为黑体，When 解析，Then 表格内段落保留为正文，不产生二级标题。
14. Given 黑体段落中含 hint=eastAsia 的 ①、→、· 或 ㈠ 等中日韩带圈字符，When 解析，Then 按 Word 的 Unicode 区块表（MS-OI29500 2.1.88）判定字体，整段仍为二级标题；没有 hint 的 ① 走西文字体。
15. Given Word 网页版保存的 word/document2.xml，When 导入，Then 按关系文件定位正文与样式，正常识别标题。
16. Given 标题或二级标题段落内有换行，When 输出，Then 与中文相邻的换行直接拼接，西文之间保留一个空格。
17. Given 方正小标宋简体或其英文名 FZXiaoBiaoSong-B05S，When 解析，Then 不作为标题字体；方正小标宋_GBK 的英文名 FZXiaoBiaoSong-B05 视同目标字体。
18. Given 浏览器暂停 requestAnimationFrame（后台标签页），When 批量导入，Then 导入照常完成；2 万段主题字体文档在 1.5 s 内解析完成。
