import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('..', import.meta.url));
const config = { root, spaceId: process.env.EGO_SPACE_ID ? Number(process.env.EGO_SPACE_ID) : null };
const result = spawnSync('ego-browser', ['nodejs'], {
  input: `const testConfig = ${JSON.stringify(config)};\n` + readFileSync(new URL('../tests/browser.mjs', import.meta.url), 'utf8'),
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
