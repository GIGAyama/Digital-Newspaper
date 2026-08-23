const http = require('http'), fs = require('fs'), path = require('path'), url = require('url');
const root = path.resolve(process.argv[2] || '.');
const port = Number(process.argv[3] || 8098);
const MIME = { '.html':'text/html; charset=utf-8', '.js':'text/javascript; charset=utf-8', '.css':'text/css; charset=utf-8',
  '.json':'application/json; charset=utf-8', '.webmanifest':'application/manifest+json', '.png':'image/png',
  '.woff2':'font/woff2', '.svg':'image/svg+xml', '.jpg':'image/jpeg' };
http.createServer((req, res) => {
  let p = decodeURIComponent(url.parse(req.url).pathname);
  if (p.endsWith('/')) p += 'index.html';
  const f = path.resolve(path.join(root, p));
  // ディレクトリの外へ出る要求は断る（../ を含む URL）
  if (f !== root && !f.startsWith(root + path.sep)) { res.writeHead(403).end(); return; }
  fs.readFile(f, (e, b) => {
    if (e) { res.writeHead(404).end('not found'); return; }
    res.writeHead(200, { 'content-type': MIME[path.extname(f)] || 'application/octet-stream', 'cache-control': 'no-store' });
    res.end(b);
  });
}).listen(port, '127.0.0.1', () => console.log('serving', root, 'on', port));
