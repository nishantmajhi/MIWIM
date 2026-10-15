"use strict";

/**
 * Tests for the development server (serve.js).
 *
 * It had none, and it had two defects a single request could reach: it bound
 * every network interface, so the folder you drop documents into was served to
 * the whole LAN, and a malformed percent-escape threw out of the request
 * handler and killed the process.
 *
 * These run the real file as a child process and talk to it over a socket,
 * because the bugs were in how the process behaves rather than in a function
 * that could be called directly. Every check here was confirmed failing
 * against the code at commit 25a70b9.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const http = require("node:http");
const net = require("node:net");
const os = require("node:os");
const path = require("node:path");
const { spawn } = require("node:child_process");

const ROOT = path.join(__dirname, "..");

/** A port nothing else in this suite uses. */
function pickPort() {
  return 19000 + Math.floor(Math.random() * 900);
}

/** Wait until something is accepting connections on host:port, or give up. */
function waitForListening(host, port, deadlineMs) {
  const started = Date.now();
  return new Promise(function (resolve, reject) {
    (function attempt() {
      const socket = net.connect(port, host);
      socket.once("connect", function () {
        socket.destroy();
        resolve();
      });
      socket.once("error", function () {
        socket.destroy();
        if (Date.now() - started > deadlineMs) return reject(new Error("never listened"));
        setTimeout(attempt, 40);
      });
    })();
  });
}

function startServer(env) {
  const port = pickPort();
  const child = spawn(process.execPath, ["serve.js"], {
    cwd: ROOT,
    env: Object.assign({}, process.env, { PORT: String(port) }, env || {}),
  });
  let stderr = "";
  let exited = null;
  child.stdout.on("data", function () {});
  child.stderr.on("data", function (d) {
    stderr += d;
  });
  child.on("exit", function (code, signal) {
    exited = { code: code, signal: signal };
  });
  return {
    port: port,
    child: child,
    stderr: function () {
      return stderr;
    },
    exited: function () {
      return exited;
    },
    ready: function () {
      return waitForListening("127.0.0.1", port, 8000);
    },
    stop: function () {
      child.kill();
    },
  };
}

/**
 * A GET that sends the path exactly as given, without a URL parser in the way.
 * `http.request` will not carry "/%" through, and "/%" is the request that
 * killed the server.
 */
function rawGet(host, port, rawPath) {
  return new Promise(function (resolve) {
    const socket = net.connect(port, host);
    let data = "";
    socket.on("connect", function () {
      socket.write("GET " + rawPath + " HTTP/1.1\r\nHost: " + host + "\r\nConnection: close\r\n\r\n");
    });
    socket.on("data", function (c) {
      data += c;
    });
    socket.on("close", function () {
      const m = data.match(/^HTTP\/1\.1 (\d{3})/);
      resolve({ status: m ? Number(m[1]) : null, raw: data });
    });
    socket.on("error", function (e) {
      resolve({ error: e.code || e.message });
    });
  });
}

function get(host, port, requestPath) {
  return new Promise(function (resolve) {
    const req = http.request({ host: host, port: port, path: requestPath }, function (res) {
      let body = "";
      res.on("data", function (c) {
        body += c;
      });
      res.on("end", function () {
        resolve({ status: res.statusCode, type: res.headers["content-type"], body: body });
      });
    });
    req.on("error", function (e) {
      resolve({ error: e.code || e.message });
    });
    req.end();
  });
}

/** Every non-loopback IPv4 address this machine answers on. */
function externalAddresses() {
  const out = [];
  const ifaces = os.networkInterfaces();
  Object.keys(ifaces).forEach(function (name) {
    (ifaces[name] || []).forEach(function (a) {
      if (a.family === "IPv4" && !a.internal) out.push(a.address);
    });
  });
  return out;
}

// ===========================================================================
test("the development server", async (t) => {
  const server = startServer();
  await server.ready();

  try {
    await t.test("serves the page", async () => {
      const res = await get("127.0.0.1", server.port, "/");
      assert.equal(res.status, 200);
      assert.match(res.type, /text\/html/);
      assert.match(res.body, /^<!doctype html>/i);
    });

    await t.test("serves index.html for the root and by name alike", async () => {
      const a = await get("127.0.0.1", server.port, "/");
      const b = await get("127.0.0.1", server.port, "/index.html");
      assert.equal(a.body, b.body);
    });

    await t.test("labels the engine files as JavaScript", async () => {
      const res = await get("127.0.0.1", server.port, "/miwim.core.js");
      assert.equal(res.status, 200);
      assert.match(res.type, /text\/javascript/);
    });

    await t.test("404s a file that is not there", async () => {
      const res = await get("127.0.0.1", server.port, "/nope.txt");
      assert.equal(res.status, 404);
    });

    await t.test("refuses to walk out of the project folder", async () => {
      for (const p of [
        "/../README.md",
        "/%2e%2e%2fREADME.md",
        "/%2e%2e%5cREADME.md",
        "/....//README.md",
      ]) {
        const res = await rawGet("127.0.0.1", server.port, p);
        assert.ok(
          res.status === 403 || res.status === 404,
          p + " answered " + res.status,
        );
      }
    });

    // ---- the two defects -------------------------------------------------

    await t.test("a malformed percent-escape is a 400, not a fatal error", async () => {
      // decodeURIComponent("%") throws URIError. Thrown from inside the
      // request handler that is an uncaught exception, and the process exits.
      const res = await rawGet("127.0.0.1", server.port, "/%");
      assert.equal(res.status, 400);
    });

    await t.test("and the server is still running afterwards", async () => {
      const res = await get("127.0.0.1", server.port, "/");
      assert.equal(res.status, 200, "server died on the previous request");
      assert.equal(server.exited(), null, "process exited: " + server.stderr().slice(0, 400));
    });

    await t.test("every malformed escape shape is survivable", async () => {
      for (const p of ["/%", "/%z", "/%zz", "/%e0%a4%a", "/a%", "/%c3%28"]) {
        const res = await rawGet("127.0.0.1", server.port, p);
        assert.ok(res.status === 400 || res.status === 404, p + " answered " + res.status);
      }
      assert.equal((await get("127.0.0.1", server.port, "/")).status, 200);
      assert.equal(server.exited(), null);
    });

    await t.test("a NUL byte in the path is a 400, not a fatal error", async () => {
      // fs.readFile rejects a path containing a NUL by throwing, and the throw
      // is synchronous, so it escapes the callback and kills the process.
      const res = await rawGet("127.0.0.1", server.port, "/a%00b");
      assert.equal(res.status, 400);
      assert.equal((await get("127.0.0.1", server.port, "/")).status, 200);
      assert.equal(server.exited(), null);
    });

    await t.test("binds the loopback interface only", async () => {
      // Liveness first. A dead server refuses every connection, so without
      // this the check passes most convincingly when the process has already
      // crashed — which is exactly what happened on the first run of this
      // file, because the malformed-escape test above had killed it.
      assert.equal(server.exited(), null, "process already exited");
      assert.equal(
        (await get("127.0.0.1", server.port, "/")).status,
        200,
        "server is not answering on loopback, so proving it is unreachable elsewhere proves nothing",
      );

      const external = externalAddresses();
      if (!external.length) {
        // Nothing to prove on a machine with no network. Assert the intent is
        // in the source instead, so the check is not silently vacuous.
        const src = require("fs").readFileSync(path.join(ROOT, "serve.js"), "utf8");
        assert.match(src, /listen\(\s*PORT\s*,\s*HOST/);
        return;
      }
      for (const address of external) {
        const res = await get(address, server.port, "/index.html");
        assert.ok(
          res.error === "ECONNREFUSED" || res.error === "ETIMEDOUT",
          "reachable from " + address + ": " + JSON.stringify(res),
        );
      }
    });
  } finally {
    server.stop();
  }
});

// ===========================================================================
test("HOST opens it up deliberately", async (t) => {
  const external = externalAddresses();
  const server = startServer({ HOST: "0.0.0.0" });
  await server.ready();
  try {
    await t.test("still serves on loopback", async () => {
      assert.equal((await get("127.0.0.1", server.port, "/")).status, 200);
    });
    await t.test("and on the network when asked", async () => {
      if (!external.length) return;
      const res = await get(external[0], server.port, "/index.html");
      assert.equal(res.status, 200);
    });
  } finally {
    server.stop();
  }
});
