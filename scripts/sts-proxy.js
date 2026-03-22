// Простой HTTP-прокси для STS API
// 1С подключается к localhost:8099 (HTTP), прокси перенаправляет на HTTPS
// Запуск: node sts-proxy.js

const http = require('http');
const https = require('https');

const PROXY_PORT = 8099;
const TARGET_HOST = 'pos.autooplata.ru';
const TARGET_BASE = '/tms';

const server = http.createServer((req, res) => {
    const targetPath = TARGET_BASE + req.url;

    const options = {
        hostname: TARGET_HOST,
        port: 443,
        path: targetPath,
        method: req.method,
        headers: { ...req.headers, host: TARGET_HOST }
    };

    console.log(`${req.method} ${req.url} → https://${TARGET_HOST}${targetPath}`);

    const proxyReq = https.request(options, (proxyRes) => {
        res.writeHead(proxyRes.statusCode, proxyRes.headers);
        proxyRes.pipe(res);
    });

    proxyReq.on('error', (e) => {
        console.error('Proxy error:', e.message);
        res.writeHead(502);
        res.end('Proxy error: ' + e.message);
    });

    req.pipe(proxyReq);
});

server.listen(PROXY_PORT, () => {
    console.log(`STS Proxy listening on http://localhost:${PROXY_PORT}`);
    console.log(`Proxying to https://${TARGET_HOST}${TARGET_BASE}`);
    console.log('В 1С укажите URL: http://localhost:8099');
});
