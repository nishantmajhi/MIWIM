"use strict";

/**
 * The preview and the download have to be the same document.
 *
 * analyzeDocx renders the first non-empty paragraph as the document's title and
 * skips it; buildConvertedDocx has no notion of a title and converts every
 * paragraph. In a document that opens straight into body text — a chapter, an
 * extract, a pasted draft — those two disagreed: the first paragraph's citation
 * was missing from the preview, every footnote number after it was one too low,
 * and the reference it cited was reported as never cited immediately before the
 * download deleted it from the bibliography.
 *
 * Confirmed failing against commit 25a70b9.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("path");

const H = require(path.join(__dirname, "helpers", "docx-fixture.js"));
const { p, buildDocx, readDocx, engine } = H;

// ===========================================================================
test("3. the preview and the download are the same document", async (t) => {
  // A document that opens straight into body text, with no title line — which
  // is what a chapter, a section extract or a pasted draft looks like.
  const titleless = [
    p("Cyber incidents rose sharply (ASD 2025) last year."),
    p("Privacy complaints also rose (OAIC 2024)."),
    p("References", { style: "Heading1" }),
    p("ASD (Australian Signals Directorate) 2025, Annual Cyber Threat Report, ASD, https://asd.gov.au/threat."),
    p("OAIC (Office of the Australian Information Commissioner) 2024, Notifiable data breaches, OAIC, https://oaic.gov.au/ndb."),
  ];

  await t.test("the first paragraph's citation is previewed", async () => {
    const buf = await buildDocx({ body: titleless });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.citationCount, 2);
    assert.deepEqual(
      analysis.notes.map(function (n) {
        return n.refId;
      }),
      ["c1", "c2"],
    );
  });

  await t.test("neither reference is reported as never cited", async () => {
    // Before the fix the interface said the ASD reference was uncited, and the
    // download then deleted it from the bibliography anyway.
    const buf = await buildDocx({ body: titleless });
    const analysis = await engine.analyzeDocx(buf);
    assert.deepEqual(analysis.uncitedRefs, []);
  });

  await t.test("the preview's numbering is the document's numbering", async () => {
    const buf = await buildDocx({ body: titleless });
    const analysis = await engine.analyzeDocx(buf);
    const built = await engine.buildConvertedDocx(buf, analysis.refs);
    const out = await readDocx(await built.zip.generateAsync({ type: "nodebuffer" }));

    assert.equal(analysis.notes.length, built.summary.footnotesCreated);
    assert.equal(out.paragraphs[0], "Cyber incidents rose sharply{1} last year.");
    assert.equal(out.paragraphs[1], "Privacy complaints also rose{2}.");
  });

  await t.test("a real title is still treated as a title", async () => {
    const buf = await buildDocx({
      body: [
        p("Cyber Security in Australia (2025)"),
        p("Incidents rose (ASD 2025)."),
        p("References", { style: "Heading1" }),
        p("ASD (Australian Signals Directorate) 2025, Report, ASD, https://asd.gov.au/t."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.docTitle, "Cyber Security in Australia (2025)");
    assert.equal(analysis.citationCount, 1);
  });

  await t.test("a first paragraph carrying an existing footnote is body text", async () => {
    const buf = await buildDocx({
      body: [
        p(["Incidents rose", { footnoteRef: 2 }, " last year."]),
        p("More text here."),
      ],
      footnotesXml: H.footnotesPart([{ id: 2, text: "An earlier note." }]),
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.notes.length, 1);
    assert.equal(analysis.notes[0].kind, "existing");
  });
});

