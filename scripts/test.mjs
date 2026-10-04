import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('..', import.meta.url));
const config = { root, spaceId: process.env.EGO_SPACE_ID ? Number(process.env.EGO_SPACE_ID) : null };
const result = spawnSync('ego-browser', ['nodejs'], {
  input: `
    const testConfig = ${JSON.stringify(config)};
    const { runBrowserTests } = await import(${JSON.stringify(new URL('../tests/browser.mjs', import.meta.url).href)});
    const task = await taskSpace(testConfig.spaceId || 'Word Markdown regression');
    console.log({ spaceId: task.spaceId });
    const page = task.page('p1');
    await runBrowserTests({
      root: testConfig.root,
      page,
      cdp: (method, params) => page.cdp(method, params),
      snapshotPage: () => page.snapshot(),
      browser: 'ego-lite Chromium'
    });
    if (!testConfig.spaceId) await task.finish({ keep: [] });
  `,
  cwd: root,
  env: process.env,
  encoding: 'utf8',
  maxBuffer: 8 * 1024 * 1024,
  timeout: 180_000
});
if (result.stdout) process.stdout.write(result.stdout);
if (result.stderr) process.stderr.write(result.stderr);
if (result.error) console.error(result.error.message);
process.exit(result.status ?? 1);
