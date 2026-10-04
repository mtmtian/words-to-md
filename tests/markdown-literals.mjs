// Text that CommonMark/GFM would read as structure if left unescaped. Each entry must render back as the same literal text,
// both as a body paragraph and as a heading. Entries never start with whitespace, because extracted paragraphs are trimmed.
export const markdownLiterals = [
  '# 不是标题', '###### 六级', '####### 七个井号', '#标签', 'C# 指南', '标题 #', '标题 ##', '#',
  '> 不是引用', '>引用', '- 不是列表', '+ 不是列表', '* 不是列表', '-', '--', '---', '- - -', '-- -', '***', '___', '_ _ _', '===', '=',
  '1. 不是列表', '1) 不是列表', '2026. 年', '123456789. 九位', '1234567890. 十位', '1.5 倍',
  '*单个星号', '两个*星号*成对', '**粗体**', '__下划线__', '_注_', '_前缀', 'user_name 与 file_path', 'snake_case_name',
  '1~3月', '1~3月与5~6月', '~~删除~~', '单个`反引号', '成对`代码`', '```', '~~~',
  '[链接](https://example.com)', '![图片](x.png)', '[1]: https://example.com', '[^1]', 'a]b',
  '<script>alert(1)</script>', '<b>粗</b>', 'a<b', 'a < b', '<http://example.com>', '<!-- 注释 -->', '<?php ?>',
  'A &amp; B', '&nbsp;', '&#123;', '&#x4e00;', 'R&D', 'a & b', '反斜杠\\结尾\\', '\\*', 'https://example.com/a_b?x=1&y=2',
  '多行\n# 第二行像标题', '多行\n- 第二行像列表', '多行\n1. 第二行像列表', '多行\n2. 第二行', '多行\n> 第二行像引用',
  '多行\n===', '多行\n---', '多行\n--', '多行\n-', '多行\n```', '多行\n    缩进四格', '多行\n\t制表符开头',
  'a | b\n--- | ---', '| a | b |\n|---|---|\n| 1 | 2 |', '单行 | 管道 | 符号', '中文——破折号…省略号「引号」', '$100 与 $200'
];
