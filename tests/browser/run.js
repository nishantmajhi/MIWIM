#!/usr/bin/env node
"use strict";

/**
 * Run the real engine files in a real browser.
 *
 * Everything in tests/*.test.js loads miwim.core.js and miwim.docx.js
 * through `require`, which takes the Node branch of their UMD wrapper and hands
 * the docx layer an xmldom parser and a Node JSZip. None of that is what runs
 * when somebody actually uses the tool. This driver serves the project over
 * loopback, loads the two files into headless Chrome exactly as index.html does,
 * and converts a document with the browser's own DOMParser, XMLSerializer and
 * JSZip.
 *
 * It needs no dependency of any kind: a ~60-line static server here, Chrome
 * invoked with --headless=new --dump-dom, and the results read back out of the
 * DOM as JSON. JSZip is served from node_modules rather than the CDN that
 * index.html uses, so the check does not depend on the network.
 *
 *   node tests/browser/run.js            run the checks
 *   node tests/browser/run.js --keep     leave the harness page for inspection
 */

const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { spawn } = require("child_process");

const ROOT = path.join(__dirname, "..", "..");
const KEEP = process.argv.indexOf("--keep") !== -1;

const CHROME_CANDIDATES = [
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  process.env.LOCALAPPDATA
    ? path.join(process.env.LOCALAPPDATA, "Google/Chrome/Application/chrome.exe")
    : null,
  "/usr/bin/google-chrome",
  "/usr/bin/chromium",
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
].filter(Boolean);

function findChrome() {
  if (process.env.CHROME && fs.existsSync(process.env.CHROME)) return process.env.CHROME;
  for (const c of CHROME_CANDIDATES) {
    try {
      if (fs.existsSync(c)) return c;
    } catch (e) {
      /* keep looking */
    }
  }
  return null;
}

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
};

/**
 * The harness page. It builds a .docx in the browser with JSZip, runs the real
 * engine over it, and writes one JSON object into #out. Every assertion lives
 * here rather than in Node, so what is checked is what the browser produced.
 */
const HARNESS = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>MIWIM browser checks</title></head>
<body><pre id="out">running</pre>
<script src="/node_modules/jszip/dist/jszip.min.js"></script>
<script src="/miwim.core.js"></script>
<script src="/miwim.docx.js"></script>
<script>
(async function () {
  var results = [];
  function check(name, fn) {
    try {
      var detail = fn();
      results.push({ name: name, ok: true, detail: detail === undefined ? null : detail });
    } catch (err) {
      results.push({ name: name, ok: false, error: String(err && err.message || err) });
    }
  }
  function eq(a, b, what) {
    var sa = JSON.stringify(a), sb = JSON.stringify(b);
    if (sa !== sb) throw new Error((what || "") + " expected " + sb + " got " + sa);
  }

  try {
    // --- the UMD browser branch -------------------------------------------
    check("core attaches itself to window.MIWIM", function () {
      if (!window.MIWIM || !window.MIWIM.core) throw new Error("no MIWIM.core");
      eq(typeof window.MIWIM.core.findCitationGroups, "function", "findCitationGroups");
    });
    check("the docx factory attaches itself too", function () {
      eq(typeof window.MIWIM.createDocxEngine, "function", "createDocxEngine");
    });
    check("the functions added in the 2026-09-26 pass are exported", function () {
      var core = window.MIWIM.core;
      eq(typeof core.splitLocator, "function", "splitLocator");
      eq(typeof core.tokenSuffixes, "function", "tokenSuffixes");
    });

    var core = window.MIWIM.core;

    // --- the logic, in the browser's own JS engine -------------------------
    check("a page locator is recognised here too", function () {
      eq(core.splitLocator(", p. 44"), { locator: "p. 44", residue: "" });
      eq(core.splitLocator(", Jones 2021"), { locator: "", residue: "Jones 2021" });
    });

    // --- a whole conversion, with the browser's DOMParser and JSZip --------
    var engine = window.MIWIM.createDocxEngine({
      DOMParser: window.DOMParser,
      XMLSerializer: window.XMLSerializer,
      JSZip: window.JSZip,
      core: core,
    });

    var W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    var RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";
    function para(text) {
      return "<w:p><w:r><w:t xml:space='preserve'>" +
        text.replace(/&/g, "&amp;").replace(/</g, "&lt;") + "</w:t></w:r></w:p>";
    }
    function heading(text) {
      return "<w:p><w:pPr><w:pStyle w:val='Heading1'/></w:pPr><w:r><w:t>" + text + "</w:t></w:r></w:p>";
    }

    var bodyXml = [
      para("Cyber Security in Australia"),
      para("As Smith (2020, p. 44) argues, incidents rose (ASD 2025)."),
      para("Complaints also rose (OAIC 2024, emphasis added)."),
      heading("References"),
      para("Smith, J 2020, A Book, Press, Sydney, https://example.org/book."),
      para("ASD (Australian Signals Directorate) 2025, Threat Report, ASD, https://asd.gov.au/threat."),
      para("OAIC (Office of the Australian Information Commissioner) 2024, Breaches, OAIC, https://oaic.gov.au/ndb."),
    ].join("");

    var zip = new JSZip();
    zip.file("[Content_Types].xml",
      "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" +
      "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
      "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>" +
      "<Default Extension='xml' ContentType='application/xml'/>" +
      "<Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/>" +
      "</Types>");
    zip.file("_rels/.rels",
      "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><Relationships xmlns='" + RELNS + "'>" +
      "<Relationship Id='rId1' Type='" + R + "/officeDocument' Target='word/document.xml'/></Relationships>");
    zip.file("word/document.xml",
      "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><w:document xmlns:w='" + W + "' xmlns:r='" + R + "'><w:body>" +
      bodyXml + "<w:sectPr><w:pgSz w:w='11906' w:h='16838'/></w:sectPr></w:body></w:document>");
    zip.file("word/_rels/document.xml.rels",
      "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><Relationships xmlns='" + RELNS + "'>" +
      "<Relationship Id='rId1' Type='" + R + "/styles' Target='styles.xml'/></Relationships>");
    zip.file("word/styles.xml",
      "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><w:styles xmlns:w='" + W + "'>" +
      "<w:style w:type='paragraph' w:styleId='Normal'><w:name w:val='Normal'/></w:style></w:styles>");

    var buffer = await zip.generateAsync({ type: "arraybuffer" });

    var analysis = await engine.analyzeDocx(buffer);
    check("the browser reads every reference", function () {
      eq(analysis.refs.length, 3, "reference count");
    });
    check("a narrative citation after 'As' is matched in the browser", function () {
      var ids = analysis.notes.map(function (n) { return n.refId; });
      eq(ids, ["c1", "c2"], "matched ids");
    });
    check("the bracket with prose in it is reported, not converted", function () {
      eq(analysis.unmatched.length, 1, "unmatched count");
      eq(analysis.unmatched[0].text, "(OAIC 2024, emphasis added)", "unmatched text");
    });
    check("the reference nothing cited is reported as uncited", function () {
      eq(analysis.uncitedRefs.map(function (r) { return r.id; }), ["c3"], "uncited");
    });

    var built = await engine.buildConvertedDocx(buffer, analysis.refs);
    check("the preview's note count is the document's note count", function () {
      eq(built.summary.footnotesCreated, analysis.notes.length, "footnote count");
    });

    var outBuf = await built.zip.generateAsync({ type: "arraybuffer" });
    var outZip = await JSZip.loadAsync(outBuf);
    var docXml = await outZip.file("word/document.xml").async("string");
    var fnXml = await outZip.file("word/footnotes.xml").async("string");
    var fnRels = await outZip.file("word/_rels/footnotes.xml.rels").async("string");

    check("the package reopens as a zip with the parts Word needs", function () {
      if (!docXml || !fnXml) throw new Error("missing a part");
      eq(/<w:footnote\\b/.test(fnXml), true, "footnotes present");
    });
    check("the locator reached the footnote", function () {
      eq(/p\\. 44/.test(fnXml), true, "p. 44 in footnotes.xml");
    });
    check("the prose bracket is still in the body, word for word", function () {
      eq(docXml.indexOf("(OAIC 2024, emphasis added)") !== -1, true, "bracket kept");
    });
    check("the converted citations left no bracket behind", function () {
      eq(docXml.indexOf("(ASD 2025)"), -1, "ASD bracket gone");
      eq(docXml.indexOf("(2020, p. 44)"), -1, "narrative bracket gone");
    });
    check("each footnote URL is a real relationship", function () {
      var ids = [];
      var re = /r:id="([^"]+)"/g;
      var m;
      while ((m = re.exec(fnXml))) ids.push(m[1]);
      if (!ids.length) throw new Error("no hyperlink in any footnote");
      ids.forEach(function (id) {
        if (fnRels.indexOf('Id="' + id + '"') === -1) throw new Error("dangling " + id);
      });
      return ids.length + " relationship(s)";
    });
    // Idempotence has to be awaited, so it sits outside check().
    try {
      var again = await engine.buildConvertedDocx(outBuf, (await engine.analyzeDocx(outBuf)).refs);
      var againXml = await (await JSZip.loadAsync(await again.zip.generateAsync({ type: "arraybuffer" })))
        .file("word/document.xml").async("string");
      results.push({
        name: "a second conversion is a no-op",
        ok: againXml === docXml && again.summary.footnotesCreated === 0,
        detail: "created " + again.summary.footnotesCreated,
      });
    } catch (err) {
      results.push({ name: "a second conversion is a no-op", ok: false, error: String(err.message || err) });
    }
  } catch (err) {
    results.push({ name: "harness itself", ok: false, error: String(err && err.stack || err) });
  }

  document.getElementById("out").textContent =
    "MRG_BEGIN" + JSON.stringify(results) + "MRG_END";
})();
</script></body></html>
`;

function serve() {
  const server = http.createServer((req, res) => {
    let p;
    try {
      p = decodeURIComponent(req.url.split("?")[0]);
    } catch (e) {
      res.writeHead(400).end("bad");
      return;
    }
    if (p === "/" || p === "/harness.html") {
      res.writeHead(200, { "Content-Type": TYPES[".html"] });
      return res.end(HARNESS);
    }
    const file = path.join(ROOT, p.replace(/^\/+/, ""));
    if (!file.startsWith(ROOT + path.sep)) {
      res.writeHead(403).end("no");
      return;
    }
    fs.readFile(file, (err, data) => {
      if (err) {
        res.writeHead(404).end("missing " + p);
        return;
      }
      res.writeHead(200, {
        "Content-Type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream",
      });
      res.end(data);
    });
  });
  return new Promise((resolve) => {
    server.listen(0, "127.0.0.1", () => resolve({ server, port: server.address().port }));
  });
}

function runChrome(chrome, url) {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), "miwim-chrome-"));
  const args = [
    "--headless=new",
    "--disable-gpu",
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-extensions",
    "--disable-component-update",
    "--user-data-dir=" + profile,
    "--virtual-time-budget=20000",
    "--dump-dom",
    url,
  ];
  return new Promise((resolve, reject) => {
    const child = spawn(chrome, args, { stdio: ["ignore", "pipe", "pipe"] });
    let out = "";
    let err = "";
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error("Chrome did not finish in 90s"));
    }, 90000);
    child.stdout.on("data", (d) => (out += d));
    child.stderr.on("data", (d) => (err += d));
    child.on("error", reject);
    child.on("exit", () => {
      clearTimeout(timer);
      try {
        fs.rmSync(profile, { recursive: true, force: true });
      } catch (e) {
        /* a locked profile directory is not a test failure */
      }
      resolve({ out, err });
    });
  });
}

(async function main() {
  const chrome = findChrome();
  if (!chrome) {
    console.error("No Chrome found. Set CHROME=<path to chrome.exe> and run again.");
    process.exit(2);
  }
  if (!fs.existsSync(path.join(ROOT, "node_modules", "jszip", "dist", "jszip.min.js"))) {
    console.error("node_modules/jszip is missing. Run `npm install` first.");
    process.exit(2);
  }

  const { server, port } = await serve();
  const url = "http://127.0.0.1:" + port + "/harness.html";
  console.log("MIWIM browser checks");
  console.log("  chrome: " + chrome);
  console.log("  page:   " + url);

  let dom;
  try {
    dom = await runChrome(chrome, url);
  } finally {
    if (!KEEP) server.close();
  }

  // Delimited at both ends, because --dump-dom prints the whole page and the
  // markup after the report has brackets of its own — reading to the last "]"
  // in the output parses most of index.html as JSON.
  const begin = dom.out.indexOf("MRG_BEGIN");
  const finish = dom.out.indexOf("MRG_END", begin);
  if (begin === -1 || finish === -1) {
    console.error("\nThe page never reported. Chrome printed:\n");
    console.error(dom.out.slice(0, 4000));
    if (dom.err) console.error("\nstderr:\n" + dom.err.slice(0, 2000));
    process.exit(1);
  }
  const json = dom.out
    .slice(begin + "MRG_BEGIN".length, finish)
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
  let results;
  try {
    results = JSON.parse(json);
  } catch (err) {
    console.error("Could not parse the page's report: " + err.message);
    console.error(json.slice(0, 1000));
    process.exit(1);
  }

  let failed = 0;
  console.log("");
  results.forEach((r) => {
    if (r.ok) {
      console.log("  ok    " + r.name + (r.detail ? "  (" + r.detail + ")" : ""));
    } else {
      failed++;
      console.log("  FAIL  " + r.name + "\n          " + (r.error || "no reason given"));
    }
  });
  console.log("\n  " + (results.length - failed) + " passed, " + failed + " failed");
  if (KEEP) console.log("  server still up at " + url + " — Ctrl+C to stop");
  process.exit(failed ? 1 : 0);
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
