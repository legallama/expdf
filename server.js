const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3000;

const mimeTypes = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.css': 'text/css',
    '.xml': 'application/xml',
    '.json': 'application/json',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.gif': 'image/gif'
};

const server = http.createServer((req, res) => {
    console.log(`${req.method} ${req.url}`);

    let filePath = req.url === '/' ? '/taskpane.html' : req.url;
    filePath = path.join(__dirname, filePath);

    fs.readFile(filePath, (err, data) => {
        if (err) {
            if (err.code === 'ENOENT') {
                res.writeHead(404, { 'Content-Type': 'text/plain' });
                res.end('404 Not Found');
            } else {
                res.writeHead(500);
                res.end('Server error');
            }
            return;
        }

        const ext = path.extname(filePath);
        const contentType = mimeTypes[ext] || 'text/plain';
        res.writeHead(200, { 'Content-Type': contentType });
        res.end(data);
    });
});

server.listen(PORT, 'localhost', () => {
    console.log(`\n✓ Server running at http://localhost:${PORT}`);
    console.log(`\n📋 Files available:`);
    console.log(`  - http://localhost:${PORT}/taskpane.html`);
    console.log(`  - http://localhost:${PORT}/taskpane.js`);
    console.log(`  - http://localhost:${PORT}/taskpane.css`);
    console.log(`  - http://localhost:${PORT}/manifest.xml`);
    console.log(`\n⚠️  Press Ctrl+C to stop the server`);
});
