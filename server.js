const http = require('http');
const fs = require('fs');
const path = require('path');

// Create a server instance
const server = http.createServer((req, res) => {
    // Extract the requested URL
    const requestedUrl = req.url === '/' ? '/index.html' : req.url;
    console.log(requested.url);

    // Construct the file path based on the requested URL
    const filePath = path.join(__dirname, 'public', requestedUrl);

    // Read the file content
    fs.readFile(filePath, (err, data) => {
        if (err) {
            // If the file is not found, return a 404 response
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('File not found');
        } else {
            // Determine the content type based on the file extension
            const extname = path.extname(filePath);
            const contentType = getContentType(extname);

            // Set the response header
            res.writeHead(200, { 'Content-Type': contentType });

            // Send the file content
            res.end(data);
        }
    });
});

const port = 8080;
server.listen(port, () => {
    console.log(`Server running at http://localhost:${port}/`);
});

function getContentType(extname) {
    const mimeTypes = {
        '.html': 'text/html',
        '.js': 'text/javascript',
        '.css': 'text/css',
        '.json': 'application/json',
        '.xml': 'application/xml',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.gif': 'image/gif',
        '.bmp': 'image/bmp',
        '.ico': 'image/x-icon',
        '.svg': 'image/svg+xml',
        '.mp3': 'audio/mpeg',
        '.ogg': 'audio/ogg',
        '.wav': 'audio/wav',
        '.mp4': 'video/mp4',
        '.avi': 'video/x-msvideo',
        '.mpeg': 'video/mpeg',
        '.pdf': 'application/pdf',
        '.doc': 'application/msword',
        '.xls': 'application/vnd.ms-excel',
        '.ppt': 'application/vnd.ms-powerpoint',
        '.zip': 'application/zip',
        '.tar': 'application/x-tar',
        '.gz': 'application/gzip',
        '.exe': 'application/octet-stream',
        '.dll': 'application/octet-stream',
        // Add more mime types as needed
      };

    return mimeTypes[extname] || 'application/octet-stream';
}

