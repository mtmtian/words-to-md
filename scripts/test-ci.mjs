import { chromium } from 'playwright';
import { mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { runBrowserTests } from '../tests/browser.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const artifactDir = path.join(root, 'test-artifacts');
await mkdir(artifactDir, { recursive: true });
const browser = await chromium.launch({ headless: true });
try {
  const context = await browser.newContext({
    viewport: { width: 1280, height: 900 },
    locale: 'zh-CN',
    acceptDownloads: true,
    permissions: ['clipboard-read', 'clipboard-write']
  });
  context.setDefaultTimeout(15_000);
  await context.tracing.start({ screenshots: true, snapshots: true, sources: true });
  try {
    const page = await context.newPage();
    const session = await context.newCDPSession(page);
    await runBrowserTests({
      root,
      page,
      cdp: (method, params) => session.send(method, params),
      snapshotPage: () => page.locator('body').ariaSnapshot(),
      browser: `Playwright Chromium ${browser.version()}`
    });
  } finally {
    await context.tracing.stop({ path: path.join(artifactDir, 'trace.zip') });
  }
} finally {
  await browser.close();
}
