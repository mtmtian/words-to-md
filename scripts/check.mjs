import { glob } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
let checked = 0;
for await (const file of glob(['src/**/*.js', 'scripts/**/*.mjs', 'tests/**/*.mjs'], { cwd: root })) {
  const result = spawnSync(process.execPath, ['--check', file], { cwd: root, stdio: 'inherit' });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
  checked++;
}
console.log(`JavaScript syntax checked: ${checked} files.`);
