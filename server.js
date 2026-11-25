/**
 * Concierto PPT Generator — Local Server
 *
 * Run:  node server.js
 * Then: open http://localhost:3000 in your browser
 *
 * This server does two things:
 *   1. Serves the HTML tool so "Create TCO Links" can work (browser requires
 *      an http:// origin, not file://, to proxy requests correctly)
 *   2. Proxies POST requests to the AWS Calculator saveAs API with the correct
 *      Origin + Referer headers that the browser cannot set on its own
 */

const http  = require('http');
const https = require('https');
const fs    = require('fs');
const path  = require('path');

const PORT     = 3000;
const HTML_FILE = path.join(__dirname, 'Concierto_PPT_Generator.html');
const SAVEAS_HOST = 'dnd5zrqcec4or.cloudfront.net';
const SAVEAS_PATH = '/Prod/v2/saveAs';

const { exec } = require('child_process');

const server = http.createServer((req, res) => {
  // CORS — allow the browser to call /proxy/saveas from the served page
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // ── RDS Rates: GET /proxy/rds-rates ──────────────────────────────────────
  // Fetches the 38 MB vantage.sh RDS JSON server-side and returns only the
  // fields needed for TCO pricing (~50 KB), avoiding browser fetch/parse limits.
  if (req.method === 'GET' && req.url === '/proxy/rds-rates') {
    const RDS_HOST = 'instances.vantage.sh';
    const RDS_PATH = '/rds/instances.json';
    const rdsOpts = {
      hostname: RDS_HOST,
      path:     RDS_PATH,
      method:   'GET',
      headers:  { 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' },
    };
    console.log('  [rds-rates] → fetching vantage.sh RDS pricing...');
    const rdsReq = https.request(rdsOpts, (upstream) => {
      let raw = '';
      upstream.on('data', c => { raw += c; });
      upstream.on('end', () => {
        try {
          const all = JSON.parse(raw);
          // Keep only the fields needed: instance_type + pricing (ondemand + reserved + multi_az.ondemand)
          const slim = all.map(inst => {
            const slimPricing = {};
            for (const [region, engines] of Object.entries(inst.pricing || {})) {
              slimPricing[region] = {};
              for (const [eng, ep] of Object.entries(engines)) {
                if (typeof ep !== 'object') continue;
                const od = ep.ondemand || 0;
                const res = ep.reserved || {};
                const mazOd = (ep.multi_az && ep.multi_az.ondemand) || 0;
                const mazRes = (ep.multi_az && ep.multi_az.reserved) || {};
                slimPricing[region][eng] = {
                  ondemand: od,
                  reserved: {
                    'yrTerm1Standard.noUpfront': res['yrTerm1Standard.noUpfront'] || 0,
                    'yrTerm3Standard.noUpfront': res['yrTerm3Standard.noUpfront'] || 0,
                    'yrTerm1DatabaseSavings.noUpfront': res['yrTerm1DatabaseSavings.noUpfront'] || 0,
                    'yrTerm3DatabaseSavings.noUpfront': res['yrTerm3DatabaseSavings.noUpfront'] || 0,
                  },
                  multi_az: mazOd > 0 ? { ondemand: mazOd, reserved: {
                    'yrTerm1Standard.noUpfront': mazRes['yrTerm1Standard.noUpfront'] || 0,
                    'yrTerm3Standard.noUpfront': mazRes['yrTerm3Standard.noUpfront'] || 0,
                  }} : undefined,
                };
              }
            }
            return { instance_type: inst.instance_type, pricing: slimPricing };
          });
          const out = JSON.stringify(slim);
          console.log(`  [rds-rates] ← OK, slimmed ${(raw.length/1024/1024).toFixed(1)}MB → ${(out.length/1024).toFixed(0)}KB`);
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(out);
        } catch (e) {
          console.error('  [rds-rates] parse error:', e.message);
          res.writeHead(502, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: e.message }));
        }
      });
    });
    rdsReq.on('error', (err) => {
      console.error('  [rds-rates] fetch error:', err.message);
      res.writeHead(502, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    });
    rdsReq.end();
    return;
  }

  // ── Proxy: POST /proxy/saveas ─────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/proxy/saveas') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      const bodyBuf = Buffer.from(body);
      const options = {
        hostname: SAVEAS_HOST,
        path:     SAVEAS_PATH,
        method:   'POST',
        headers:  {
          'Content-Type':   'application/json',
          'Content-Length': bodyBuf.length,
          'Origin':         'https://calculator.aws',
          'Referer':        'https://calculator.aws/',
          'User-Agent':     'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
          'Accept':         'application/json, text/plain, */*',
        },
      };

      console.log(`  [proxy] → AWS Calculator saveAs`);
      const proxy = https.request(options, (upstream) => {
        let data = '';
        upstream.on('data', c => { data += c; });
        upstream.on('end', () => {
          try {
            const parsed = JSON.parse(data);
            const inner  = typeof parsed.body === 'string' ? JSON.parse(parsed.body) : parsed;
            const key    = inner.savedKey || '(no savedKey)';
            console.log(`  [proxy] ← HTTP ${upstream.statusCode} | savedKey: ${key}`);
          } catch(e) {
            console.log(`  [proxy] ← HTTP ${upstream.statusCode} | raw: ${data.slice(0,120)}`);
          }
          res.writeHead(upstream.statusCode, { 'Content-Type': 'application/json' });
          res.end(data);
        });
      });

      proxy.on('error', (err) => {
        console.error('[proxy error]', err.message);
        res.writeHead(502, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: err.message }));
      });

      proxy.write(bodyBuf);
      proxy.end();
    });
    return;
  }

  // ── Serve the HTML tool ───────────────────────────────────────────────────
  if (req.method === 'GET' && (req.url === '/' || req.url === '/index.html')) {
    fs.readFile(HTML_FILE, (err, data) => {
      if (err) {
        res.writeHead(404);
        res.end('Concierto_PPT_Generator.html not found in this directory.');
        return;
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(data);
    });
    return;
  }

  res.writeHead(404);
  res.end('Not found');
});

function tryListen(port) {
  server.removeAllListeners('error');
  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.log('');
      console.log(`  Port ${port} is already in use.`);
      console.log(`  The Concierto PPT Generator may already be running at http://localhost:${port}`);
      console.log('  Stop the existing server with Ctrl + C before starting a new one.');
      console.log('');
      process.exit(0);
    } else {
      console.error('\n  Server error:', err.message);
      process.exit(1);
    }
  });
  server.listen(port, '127.0.0.1', () => {
    const url  = `http://localhost:${port}`;
    const pad  = ' '.repeat(Math.max(0, 28 - url.length));
    console.log('');
    console.log('  ╔══════════════════════════════════════════════════╗');
    console.log('  ║   Concierto PPT Generator — Local Server         ║');
    console.log('  ╠══════════════════════════════════════════════════╣');
    console.log(`  ║   Open in browser:  ${url}${pad}║`);
    console.log('  ║   Stop server:      Ctrl + C                     ║');
    console.log('  ╚══════════════════════════════════════════════════╝');
    console.log('');
    // Auto-open in default browser
    exec(`open "${url}"`, err => {
      if (err) console.log(`  (Could not auto-open browser — navigate to ${url} manually)`);
    });
  });
}

tryListen(PORT);
