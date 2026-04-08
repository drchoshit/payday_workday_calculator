const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = 5500;
const ROOT = process.cwd();

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};

http
  .createServer((req, res) => {
    const urlPath = (req.url || "/").split("?")[0];
    const relPath = urlPath === "/" ? "/index.html" : urlPath;
    const filePath = path.join(ROOT, relPath);

    fs.readFile(filePath, (err, data) => {
      if (err) {
        if (urlPath !== "/" && !path.extname(urlPath)) {
          const fallback = path.join(ROOT, "index.html");
          return fs.readFile(fallback, (fallbackErr, fallbackData) => {
            if (fallbackErr) {
              res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
              res.end("Not Found");
              return;
            }
            res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
            res.end(fallbackData);
          });
        }
        res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
        res.end("Not Found");
        return;
      }

      const ext = path.extname(filePath).toLowerCase();
      res.writeHead(200, { "Content-Type": MIME[ext] || "application/octet-stream" });
      res.end(data);
    });
  })
  .listen(PORT, () => {
    console.log(`Local server running: http://localhost:${PORT}`);
  });
