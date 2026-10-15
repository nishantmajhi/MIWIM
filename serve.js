#!/usr/bin/env node
"use strict";

/**
 * A minimal static file server for local development.
 *
 * MIWIM is a plain HTML page, so opening index.html directly usually
 * works — but browsers block file:// pages from reading local files in some
 * configurations, which breaks the drop zone. Serving the folder over http
 * sidesteps that.
 *
 * This exists as a file rather than an inline `node -e` string in package.json
 * so that `npm run serve` prints one readable line instead of echoing the
 * entire program back at you.
 */

const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = Number(process.env.PORT) || 8080;

// Loopback only. This folder is where you drop the .docx files you are testing
// with, and the whole claim of the tool is that those documents never leave
// your machine — binding every interface hands them to anyone on the same
// network, which on a campus or a cafe wifi is a lot of people. Set HOST
// explicitly if you actually want to reach it from another device.
const HOST = process.env.HOST || "127.0.0.1";

const ROOT = __dirname;

const CONTENT_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".svg": "image/svg+xml",
};

function badRequest(res, why) {
  res.writeHead(400, { "Content-Type": "text/plain" });
  res.end("Bad request: " + why);
}

const server = http.createServer((req, res) => {
  // decodeURIComponent throws URIError on any half-written escape — "/%",
  // "/%zz", a truncated UTF-8 sequence. Thrown from inside a request handler
  // that is an uncaught exception, so before this try/catch a single request
  // to /% stopped the server for everyone.
  let requested;
  try {
    requested = decodeURIComponent(req.url.split("?")[0]);
  } catch (err) {
    return badRequest(res, "that URL is not valid percent-encoding");
  }

  // A NUL byte makes every fs call throw rather than call back, which escaped
  // the callback and killed the process the same way. No real filename
  // contains one, so there is nothing to serve here either way.
  if (requested.indexOf("\0") !== -1) {
    return badRequest(res, "that path contains a NUL byte");
  }

  const relative = requested === "/" ? "index.html" : requested.replace(/^\/+/, "");
  const filePath = path.join(ROOT, relative);

  // Refuse to serve anything outside the project folder, so a crafted path
  // like /../../.ssh/id_rsa can't walk out of it.
  if (!filePath.startsWith(ROOT + path.sep) && filePath !== ROOT) {
    res.writeHead(403, { "Content-Type": "text/plain" });
    return res.end("Forbidden");
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" });
      return res.end("Not found: " + relative);
    }
    const type = CONTENT_TYPES[path.extname(filePath).toLowerCase()] || "application/octet-stream";
    res.writeHead(200, { "Content-Type": type, "Cache-Control": "no-cache" });
    res.end(data);
  });
});

server.listen(PORT, HOST, () => {
  console.log("MIWIM is running at http://localhost:" + PORT);
  console.log("Press Ctrl+C to stop.");
});

server.on("error", (err) => {
  if (err.code === "EADDRINUSE") {
    console.error(
      "Port " + PORT + " is already in use. Try: PORT=8081 npm run serve",
    );
    process.exit(1);
  }
  throw err;
});
