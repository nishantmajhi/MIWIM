"use strict";

/**
 * A footnote marker that replaces a citation inside a hyperlink.
 *
 * The marker is deliberately lifted out of the link, because a footnote number
 * rendered as part of a hyperlink is both wrong to click and wrong to look at.
 * What was missing was the other half: anything that followed the citation
 * inside the link stayed inside it, so the marker landed after the whole link.
 * "see the guide (ASD 2025) online" became "see the guide online¹" instead of
 * "see the guide¹ online".
 *
 * Confirmed failing against commit 25a70b9.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("path");

const H = require(path.join(__dirname, "helpers", "docx-fixture.js"));
const { p, buildDocx, readDocx, engine } = H;

// ===========================================================================
test("4. a marker lands where its citation was inside a hyperlink", async (t) => {
  await t.test("text after the citation stays after the marker", async () => {
    const body = [
      p("Links"),
      p([
        { text: "see the ", rid: "rId9" },
        { text: "guide (ASD 2025) online", rid: "rId9" },
        " and more.",
      ]),
      p("References", { style: "Heading1" }),
      p("ASD (Australian Signals Directorate) 2025, Report, ASD, https://asd.gov.au/t."),
    ];
    const buf = await buildDocx({
      body,
      rels: [{ id: "rId9", target: "https://example.org/guide" }],
    });
    const analysis = await engine.analyzeDocx(buf);
    const built = await engine.buildConvertedDocx(buf, analysis.refs);
    const out = await readDocx(await built.zip.generateAsync({ type: "nodebuffer" }));
    assert.equal(out.paragraphs[1], "see the guide{1} online and more.");
  });

  await t.test("the link either side of the marker is still a link", async () => {
    const body = [
      p("Links"),
      p([{ text: "see the guide (ASD 2025) online", rid: "rId9" }]),
      p("References", { style: "Heading1" }),
      p("ASD (Australian Signals Directorate) 2025, Report, ASD, https://asd.gov.au/t."),
    ];
    const buf = await buildDocx({
      body,
      rels: [{ id: "rId9", target: "https://example.org/guide" }],
    });
    const analysis = await engine.analyzeDocx(buf);
    const built = await engine.buildConvertedDocx(buf, analysis.refs);
    const zipped = await built.zip.generateAsync({ type: "nodebuffer" });
    const xml = await (await H.JSZip.loadAsync(zipped))
      .file("word/document.xml")
      .async("string");
    // Two hyperlinks now, both pointing at the same relationship, with the
    // footnote marker between them rather than tacked on the end.
    const links = xml.match(/<w:hyperlink[^>]*r:id="rId9"/g) || [];
    assert.equal(links.length, 2);
    const out = await readDocx(zipped);
    assert.equal(out.paragraphs[1], "see the guide{1} online");
  });
});
