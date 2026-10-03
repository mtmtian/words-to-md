import http from 'node:http';
import { readFile } from 'node:fs/promises';
const port = Number(process.env.PORT || 4173);
const file = new URL('../dist/word-to-markdown.html', import.meta.url);
http.createServer(async (req, res) => {
  if (!['GET', 'HEAD'].includes(req.method)) { res.writeHead(405); return res.end(); }
  if (!['/', '/word-to-markdown.html'].includes(req.url?.split('?')[0])) { res.writeHead(404); return res.end('Not found'); }
  try { const data = await readFile(file); res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' }); res.end(req.method === 'HEAD' ? undefined : data); }
  catch { res.writeHead(500); res.end('Run npm run build first.'); }
}).listen(port, '127.0.0.1', () => console.log(`Local preview: http://127.0.0.1:${port}`));
