// Both runners use these acceptance checks against the built file:// HTML.
// The caller owns the isolated browser, CDP session and diagnostic snapshots.
export async function runBrowserTests({ root, page, cdp, snapshotPage, browser }) {
  const fs = await import('node:fs/promises');
  const { pathToFileURL } = await import('node:url');
  const path = await import('node:path');
  const assert = (await import('node:assert/strict')).default;
  const { createHash } = await import('node:crypto');
  const rootUrl = pathToFileURL(root.endsWith('/') ? root : root + '/');
  const { parserCases } = await import(new URL('tests/parser-cases.mjs', rootUrl).href);
  const cases = await parserCases(rootUrl);
  const expected = await fs.readFile(path.join(root, 'tests/expected.md'), 'utf8');
  const artifactDir = path.join(root, 'test-artifacts');
  await fs.mkdir(artifactDir, { recursive: true });
  const checks = [];
  function check(name, condition) { assert.ok(condition, name); checks.push({ name, result: 'pass' }); }
  async function snapshot() { console.log(await snapshotPage()); }
  async function waitForFiles(count) { await page.waitForFunction(n => document.querySelector('#file-count').textContent === String(n) && document.querySelector('#progress').hidden, count, { timeout: 15000 }); }
  const names = ['01_直接字体.docx', '02_样式继承.docx', '03_主题字体.docx', '04_缺失标题与混合字体.docx'];
  const fixtures = names.map(name => path.join(root, 'demos', name));
  let failure = null;
  try {
    await cdp('Network.enable');
    await cdp('Network.emulateNetworkConditions', { offline: true, latency: 0, downloadThroughput: -1, uploadThroughput: -1 });
    await page.goto(pathToFileURL(path.join(root, 'dist/word-to-markdown.html')).href);
    await snapshot();
    check('file:// 断网加载且初始下载禁用', await page.evaluate(() => !!window.WordMD && document.querySelector('#download-output').disabled));
    const results = await page.evaluate(cases => cases.map(test => {
      const bytes = Uint8Array.from(atob(test.data), c => c.charCodeAt(0));
      try {
        const started = performance.now();
        const doc = window.WordMD.parseDocx(bytes, test.filename || 'test.docx');
        const ms = Math.round(performance.now() - started);
        const md = window.WordMD.mergeDocuments([doc]);
        return { name: test.name, title: doc.title, headings: doc.headings, text: doc.paragraphs.map(p => p.text), warnings: doc.warnings, md, ms };
      } catch (error) { return { name: test.name, error: error.message }; }
    }), cases);
    for (let i = 0; i < cases.length; i++) {
      const test = cases[i], result = results[i];
      if (test.error) assert.ok(result.error?.includes(test.error), `${test.name}: ${JSON.stringify(result)}`);
      else {
        assert.equal(result.error, undefined, test.name);
        if (test.title !== undefined) assert.equal(result.title, test.title, test.name);
        if (test.headings !== undefined) assert.equal(result.headings, test.headings, test.name);
        if (test.text) assert.deepEqual(result.text, test.text, test.name);
        if (test.warning) assert.ok(result.warnings.some(w => w.includes(test.warning)), test.name);
        for (const literal of test.markdownIncludes || []) assert.ok(result.md.includes(literal), `${test.name}: ${literal}`);
        if (test.maxMs) assert.ok(result.ms < test.maxMs, `${test.name}: ${result.ms} ms`);
      }
      checks.push({ name: test.name, result: 'pass' });
    }
    check('文件大小 30 MB 上限', await page.evaluate(() => {
      try { window.WordMD.parseDocx(new Uint8Array(30 * 1024 * 1024 + 1), 'large.docx'); return false; }
      catch (e) { return e.message.includes('30 MB'); }
    }));

    await page.setInputFiles('#file-input', fixtures);
    await waitForFiles(4);
    await snapshot();
    const imported = await page.evaluate(() => ({
      md: document.querySelector('#markdown-output').value,
      h1: document.querySelectorAll('#preview h1').length,
      h2: document.querySelectorAll('#preview h2').length,
      scripts: document.querySelectorAll('#preview script, #preview img').length,
      resources: performance.getEntriesByType('resource').length,
      text: document.querySelector('#preview').textContent,
      warnings: document.querySelector('#file-list').textContent
    }));
    assert.equal(imported.md, expected, '合并内容逐字符等于人工预期');
    checks.push({ name: '四份真实 DOCX 合并逐字符比对', result: 'pass' });
    check('一级标题 4 个与二级标题 7 个', imported.h1 === 4 && imported.h2 === 7);
    check('HTML 原文不会创建可执行 DOM', imported.scripts === 0 && imported.text.includes('<script>alert(1)</script>'));
    check('没有匹配标题时明确提示', imported.warnings.includes('使用文件名') && imported.warnings.includes('未找到方正小标宋'));
    check('离线导入没有子资源请求', imported.resources === 0);
    check('长文档的下载按钮留在可见区域', await page.evaluate(() => {
      const rect = document.querySelector('#download-output').getBoundingClientRect();
      return rect.top >= 0 && rect.bottom <= innerHeight;
    }));

    await page.click('#tab-source');
    check('Markdown 标签显示源码', await page.evaluate(() => !document.querySelector('#source-panel').hidden && document.querySelector('#preview').hidden));
    await page.press('#tab-source', 'ArrowLeft');
    check('键盘切换阅读视图', await page.evaluate(() => document.querySelector('#tab-preview').getAttribute('aria-selected') === 'true'));
    await page.click('button[aria-label="下移 01_直接字体.docx"]');
    check('调整顺序同步合并正文', await page.evaluate(() => document.querySelector('#preview h1').textContent === '公共服务资料归档工作指引' && document.querySelector('#markdown-output').value.startsWith('# 公共服务资料归档工作指引')));
    await snapshot();
    await page.click('button[aria-label="移除 02_样式继承.docx"]');
    await waitForFiles(3);
    check('删除文档同步结果', await page.evaluate(() => !document.querySelector('#markdown-output').value.includes('公共服务资料归档工作指引')));
    await page.click('#clear-all');
    check('清空恢复初始状态', await page.evaluate(() => document.querySelector('#file-count').textContent === '0' && document.querySelector('#download-output').disabled));
    // Background tabs never run requestAnimationFrame callbacks; importing must not depend on them.
    await page.evaluate(() => { window.__requestAnimationFrame = window.requestAnimationFrame; window.requestAnimationFrame = () => 0; });
    await page.click('#load-demo');
    await waitForFiles(4);
    await page.evaluate(() => { window.requestAnimationFrame = window.__requestAnimationFrame; });
    check('暂停 requestAnimationFrame（如后台标签页）时导入仍能完成', true);
    check('内置示例使用同一解析流程', await page.evaluate(expected => document.querySelector('#markdown-output').value === expected, expected));

    const badDoc = path.join(artifactDir, 'old.doc'), badZip = path.join(artifactDir, 'damaged.docx');
    await fs.writeFile(badDoc, Buffer.from([0xd0, 0xcf, 0x11, 0xe0]));
    await fs.writeFile(badZip, 'not a zip');
    await page.setInputFiles('#file-input', [badDoc, badZip]);
    await waitForFiles(4);
    check('错误文件不破坏已合并内容', await page.evaluate(expected => document.querySelector('#markdown-output').value === expected && document.querySelector('#import-errors').textContent.includes('2 项未加入合并'), expected));
    await snapshot();
    await page.click('#clear-all');
    await page.setInputFiles('#file-input', [badZip, fixtures[0]]);
    await waitForFiles(1);
    check('混合成功失败批次继续解析', await page.evaluate(() => document.querySelector('#import-errors').textContent.includes('damaged.docx') && document.querySelector('#preview h1').textContent === '关于推进文档数字化整理的通知'));
    await page.setInputFiles('#file-input', [fixtures[0]]);
    await waitForFiles(1);
    check('重复文件有明确跳过提示', await page.evaluate(() => document.querySelector('#import-errors').textContent.includes('修改时间')));
    await page.click('#clear-all');
    await page.click('#load-demo');
    await waitForFiles(4);
    await snapshot();

    const downloadWait = page.waitForEvent('download', { timeout: 15000 });
    await page.click('#download-output');
    const downloaded = await downloadWait;
    const downloadPath = path.join(artifactDir, 'downloaded.md');
    await downloaded.saveAs(downloadPath);
    assert.equal(await fs.readFile(downloadPath, 'utf8'), expected);
    check('浏览器真实下载 UTF-8 .md 与预览一致', downloaded.suggestedFilename() === '合并文档.md');

    const demoWait = page.waitForEvent('download', { timeout: 15000 });
    await page.click('#download-demo');
    const demoDownload = await demoWait;
    await demoDownload.saveAs(path.join(artifactDir, 'demo-download.zip'));
    const { unzipSync } = await import(new URL('node_modules/fflate/esm/index.mjs', rootUrl).href);
    const demoZip = unzipSync(await fs.readFile(path.join(artifactDir, 'demo-download.zip')));
    check('浏览器下载示例包包含四份原始 Word', names.every((name, i) => Buffer.from(demoZip[name]).equals(Buffer.from([])) === false) && names.length === Object.keys(demoZip).filter(x => x.endsWith('.docx')).length);
    for (const name of names) assert.deepEqual(Buffer.from(demoZip[name]), await fs.readFile(path.join(root, 'demos', name)));

    const popupWait = page.waitForEvent('popup', { timeout: 10000 });
    await page.click('#open-output');
    const popup = await popupWait;
    await popup.waitForFunction(() => location.protocol === 'blob:' && !!document.querySelector('pre'), undefined, { timeout: 10000 });
    const popupContent = await popup.evaluate(() => ({ text: document.querySelector('pre')?.textContent, type: document.contentType, opener: !!window.opener }));
    assert.equal(popupContent.text, expected);
    check('新标签页为相同纯文本且断开 opener', popupContent.type === 'text/plain' && !popupContent.opener);
    await popup.close();

    await page.click('#copy-output');
    await page.waitForFunction(() => document.querySelector('#toast').textContent.includes('已复制'), undefined, { timeout: 5000 });
    check('用户点击复制返回成功状态', true);
    await page.evaluate(() => {
      const sink = document.createElement('textarea');
      sink.id = 'clipboard-verification';
      sink.setAttribute('aria-label', '测试剪贴板内容');
      sink.style.cssText = 'position:fixed;inset:20px;width:300px;height:60px;z-index:50';
      document.body.append(sink);
    });
    await page.focus('#clipboard-verification');
    await page.press('#clipboard-verification', 'ControlOrMeta+V');
    check('粘贴验证实际剪贴板内容一致', await page.evaluate(expected => document.querySelector('#clipboard-verification').value === expected, expected));
    await page.evaluate(() => document.querySelector('#clipboard-verification').remove());

    await cdp('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true });
    check('390px 窄屏无横向溢出', await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    await snapshot();
    await page.click('#tab-source');
    check('窄屏源码与下载可用', await page.evaluate(() => !document.querySelector('#source-panel').hidden && !document.querySelector('#download-output').disabled));
    await page.evaluate(() => document.querySelector('.output-panel').scrollIntoView({ block: 'start' }));
    await page.screenshot({ path: path.join(artifactDir, 'mobile.png') });
    await cdp('Emulation.clearDeviceMetricsOverride');
    await page.click('#tab-preview');
    await page.evaluate(() => { window.scrollTo(0, 0); document.querySelector('#preview').scrollTop = 0; });
    await page.screenshot({ path: path.join(artifactDir, 'desktop.png') });
    await snapshot();
  } catch (error) {
    failure = error.stack || error.message;
    console.error(failure);
    await page.screenshot({ path: path.join(artifactDir, 'failure.png') }).catch(() => {});
  } finally {
    await cdp('Emulation.clearDeviceMetricsOverride').catch(() => {});
    await cdp('Network.emulateNetworkConditions', { offline: false, latency: 0, downloadThroughput: -1, uploadThroughput: -1 }).catch(() => {});
    const inputs = [];
    for (const filename of fixtures) inputs.push({ name: path.basename(filename), sha256: createHash('sha256').update(await fs.readFile(filename)).digest('hex') });
    await fs.writeFile(path.join(artifactDir, 'results.json'), JSON.stringify({ checkedAt: new Date().toISOString(), browser, runtime: { node: process.version, platform: process.platform, arch: process.arch }, environment: 'file://, offline during checks', inputs, htmlSha256: createHash('sha256').update(await fs.readFile(path.join(root, 'dist/word-to-markdown.html'))).digest('hex'), checks, passed: checks.length, failure }, null, 2));
  }
  console.log(JSON.stringify({ passed: checks.length, failure, artifacts: artifactDir }));
  if (failure) throw new Error('Regression checks failed');
}
