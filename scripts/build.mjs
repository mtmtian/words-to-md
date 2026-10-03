import { build } from 'esbuild';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = fileURLToPath(new URL('..', import.meta.url));
const app = await build({ entryPoints: [path.join(root, 'src/app.js')], bundle: true, write: false, minify: true, target: ['es2022'], legalComments: 'inline' });
const css = await readFile(path.join(root, 'src/style.css'), 'utf8');
const template = await readFile(path.join(root, 'src/index.html'), 'utf8');
const js = app.outputFiles[0].text.replace(/<\/script/gi, '<\\/script');
const license = await readFile(path.join(root, 'node_modules/fflate/LICENSE'), 'utf8');
const html = template.replace('/* STYLES */', () => css).replace('/* APP */', () => js).replace('</head>', () => `<!-- Bundled dependency: fflate 0.8.2\n${license.replace(/--/g, '—')}\n-->\n</head>`);
await mkdir(path.join(root, 'dist'), { recursive: true });
await writeFile(path.join(root, 'dist/word-to-markdown.html'), html);
console.log(`Built dist/word-to-markdown.html (${Buffer.byteLength(html).toLocaleString()} bytes); fully offline, no external assets.`);
