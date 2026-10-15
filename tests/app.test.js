"use strict";

/**
 * Tests for the browser wiring.
 *
 * The engine tests all run under Node with injected dependencies, which proves
 * the logic but not that the page actually loads it. These checks cover the
 * seam: that index.html pulls in the two modules and JSZip, that the modules
 * export themselves correctly when there is no `module` object (the browser
 * branch of the UMD wrapper), and that the element ids the UI script reaches
 * for genuinely exist in the markup.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");

/** The favicon's SVG source, decoded back out of its data URI. */
function faviconSvg() {
  const m = html.match(/href="data:image\/svg\+xml,([^"]+)"/);
  assert.ok(m, "expected an inline SVG favicon");
  return decodeURIComponent(m[1]);
}

// ===========================================================================
test("index.html", async (t) => {
  await t.test("is a complete document", () => {
    assert.match(html, /^<!doctype html>/i);
    assert.ok(html.trim().endsWith("</html>"));
  });

  await t.test("loads JSZip and both engine modules", () => {
    assert.match(html, /jszip.*\.min\.js/);
    assert.match(html, /<script src="miwim\.core\.js"><\/script>/);
    assert.match(html, /<script src="miwim\.docx\.js"><\/script>/);
    // Order matters: the docx layer is handed core at construction time.
    assert.ok(
      html.indexOf("miwim.core.js") < html.indexOf("miwim.docx.js"),
      "core must load before docx",
    );
  });

  await t.test("contains every element id the script uses", () => {
    // getElementById returns null silently, so a renamed id fails at runtime
    // rather than at load. This catches it here instead.
    const used = new Set();
    const re = /getElementById\("([^"]+)"\)/g;
    let m;
    while ((m = re.exec(html))) used.add(m[1]);

    assert.ok(used.size > 0, "expected the script to look up some elements");
    used.forEach((id) => {
      if (id === "mrg-file-input") return; // created by the script itself
      assert.match(html, new RegExp('id="' + id + '"'), "missing element: " + id);
    });
  });

  await t.test("defines every CSS class the script applies", () => {
    const used = new Set();

    // Classes written into markup or into a template literal.
    const inMarkup = /class="((?:mrg-[\w-]+\s*)+)"/g;
    let m;
    while ((m = inMarkup.exec(html))) {
      m[1]
        .trim()
        .split(/\s+/)
        .forEach((c) => used.add(c));
    }

    // Classes assigned from JavaScript — `el.className = "mrg-fn-note"` never
    // appears in the markup, so a typo there would slip past a markup-only
    // check and simply render unstyled.
    const inScript = /className = "((?:mrg-[\w-]+\s*)+)"/g;
    while ((m = inScript.exec(html))) {
      m[1]
        .trim()
        .split(/\s+/)
        .forEach((c) => used.add(c));
    }

    // Classes toggled at runtime rather than assigned outright.
    ["mrg-active", "mrg-drag", "mrg-nolink"].forEach((c) => used.add(c));

    assert.ok(used.size > 10, "expected to find a good number of classes");
    used.forEach((cls) => {
      assert.match(html, new RegExp("\\." + cls + "[\\s,:{.]"), "undefined class: " + cls);
    });
  });

  await t.test("carries a favicon inlined as a data URI", () => {
    // Inline, so the icon costs no network request and the page stays a single
    // self-contained file you can open from anywhere.
    const link = html.match(/<link[^>]*rel="icon"[^>]*>/);
    assert.ok(link, "expected a rel=icon link");
    assert.match(link[0], /type="image\/svg\+xml"/);
    assert.match(link[0], /href="data:image\/svg\+xml,/);
  });

  await t.test("has a favicon that is valid, self-contained SVG", () => {
    // A favicon renders outside the document, so anything it needs from the
    // page — a CSS variable, a webfont — silently fails to apply. This checks
    // it depends on neither, and that URI encoding didn't mangle the markup.
    const svg = faviconSvg();
    assert.match(svg, /^<svg\b/);
    assert.match(svg, /xmlns='http:\/\/www\.w3\.org\/2000\/svg'/);
    assert.ok(svg.trim().endsWith("</svg>"));
    assert.doesNotMatch(svg, /var\(--/, "favicon must not rely on CSS variables");
    assert.doesNotMatch(svg, /<text/, "favicon must not rely on a webfont");
  });

  await t.test("draws the favicon in the same colours as the wordmark logo", () => {
    // The favicon is a separate drawing of the same mark, so the two can drift
    // apart. Pinning the palette to the page's own variables catches it.
    const svg = faviconSvg();
    const paletteOf = (name) => {
      const m = html.match(new RegExp("--" + name + ": (#[0-9a-f]{6});"));
      assert.ok(m, "missing --" + name);
      return m[1];
    };
    assert.ok(svg.includes(paletteOf("paper-raised")), "circle fill");
    assert.ok(svg.includes(paletteOf("ink")), "ink stroke and numeral");
    assert.ok(svg.includes(paletteOf("gold")), "gold underline");
  });

  await t.test("uses no browser storage APIs", () => {
    // Nothing about this tool needs persistence, and storing someone's
    // document text would be a surprise they never agreed to.
    assert.doesNotMatch(html, /localStorage|sessionStorage|indexedDB/);
  });

  await t.test("sends the document nowhere", () => {
    // The privacy claim in the README has to be true: no fetch, no XHR, no
    // form post. The only network requests are the fonts and JSZip in <head>.
    const scriptBodies = html.split("<script").slice(3).join("<script");
    assert.doesNotMatch(scriptBodies, /\bfetch\s*\(|XMLHttpRequest|navigator\.sendBeacon/);
    assert.doesNotMatch(html, /<form\b/);
  });

  await t.test("opens footnote links safely", () => {
    // target="_blank" without noopener hands the opened page a handle back to
    // this one.
    assert.match(html, /rel = "noreferrer noopener"/);
  });
});

// ===========================================================================
test("the UMD wrapper's browser branch", async (t) => {
  // Load both modules the way a <script> tag would: no `module`, no `exports`,
  // no `require`. If either file assumed CommonJS, this throws.
  function loadInBrowserSandbox() {
    const sandbox = { console: console };
    sandbox.self = sandbox;
    sandbox.window = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(
      fs.readFileSync(path.join(ROOT, "miwim.core.js"), "utf8"),
      sandbox,
      { filename: "miwim.core.js" },
    );
    vm.runInContext(
      fs.readFileSync(path.join(ROOT, "miwim.docx.js"), "utf8"),
      sandbox,
      { filename: "miwim.docx.js" },
    );
    return sandbox;
  }

  await t.test("attaches both modules to a single global", () => {
    const sandbox = loadInBrowserSandbox();
    assert.equal(typeof sandbox.MIWIM, "object");
    assert.equal(typeof sandbox.MIWIM.core, "object");
    assert.equal(typeof sandbox.MIWIM.createDocxEngine, "function");
  });

  await t.test("exposes the functions the page calls", () => {
    const sandbox = loadInBrowserSandbox();
    assert.equal(typeof sandbox.MIWIM.core.splitParagraphForPreview, "function");
    const engine = sandbox.MIWIM.createDocxEngine({
      DOMParser: require("@xmldom/xmldom").DOMParser,
      XMLSerializer: require("@xmldom/xmldom").XMLSerializer,
      JSZip: require("jszip"),
      core: sandbox.MIWIM.core,
    });
    assert.equal(typeof engine.analyzeDocx, "function");
    assert.equal(typeof engine.buildConvertedDocx, "function");
  });

  await t.test("leaves no other globals behind", () => {
    const sandbox = loadInBrowserSandbox();
    const own = Object.keys(sandbox).filter(
      (k) => !["console", "self", "window", "MIWIM"].includes(k),
    );
    assert.deepEqual(own, []);
  });
});
