"use strict";

/**
 * Tests for documents that already contain footnotes.
 *
 * Three situations, each behaving differently:
 *
 *   footnotes, no references   → nothing to convert; shown as it is
 *   references, no footnotes   → the ordinary conversion
 *   both                       → references converted, existing notes kept,
 *                                everything numbered together
 *
 * Plus the spacing that keeps a page of long URLs readable.
 */

const test = require("node:test");
const assert = require("node:assert/strict");

const h = require("./helpers/docx-fixture.js");
const { p, buildDocx, convert, engine } = h;

const NOTE_A = "https://example.com/existing-a";
const NOTE_B = "https://example.com/existing-b";

/** A document whose citations were already converted to linked footnotes. */
function alreadyConvertedDoc() {
  return buildDocx({
    footnotesXml: h.footnotesPart([
      { id: 1, rid: "rId1", url: NOTE_A },
      { id: 2, rid: "rId2", url: NOTE_B },
    ]),
    footnotesRelsXml: h.footnotesRels([
      { id: "rId1", target: NOTE_A },
      { id: "rId2", target: NOTE_B },
    ]),
    body: [
      p("A Report", { style: "Title" }),
      p(["Growth was strong", { footnoteRef: 1 }, " last year."]),
      p(["It held into the spring", { footnoteRef: 2 }, "."]),
      p("Bibliography", { style: "Heading1" }),
      p("[Insert citations and references using the footnotes]"),
    ],
  });
}

/** A document part-way through: one hand-written note, one live citation. */
function mixedDoc() {
  return buildDocx({
    footnotesXml: h.footnotesPart([{ id: 4, text: "A hand-written note." }]),
    body: [
      p("A Report", { style: "Title" }),
      p(["An aside", { footnoteRef: 4 }, " opens the argument."]),
      p("Growth was strong (Smith 2020) last year."),
      p("Later still (Jones 2021) it held."),
      p("References", { style: "Heading1" }),
      p("Smith, J 2020, A Book, https://example.com/smith."),
      p("Jones, K 2021, Another Book, https://example.com/jones."),
    ],
  });
}

// ===========================================================================
test("a document that only has footnotes", async (t) => {
  const buffer = await alreadyConvertedDoc();
  const analysis = await engine.analyzeDocx(buffer);

  await t.test("finds the footnotes it already contains", () => {
    assert.equal(analysis.existingFootnoteCount, 2);
    assert.deepEqual(
      analysis.existingFootnotes.map((f) => f.url),
      [NOTE_A, NOTE_B],
    );
  });

  await t.test("recognises there is nothing left to convert", () => {
    assert.equal(analysis.citationCount, 0);
    assert.equal(analysis.refs.length, 0);
    assert.equal(analysis.alreadyConverted, true);
  });

  await t.test("shows the document as it is, footnotes included", () => {
    // The old behaviour was to show nothing at all here, which made a
    // perfectly good document look empty.
    assert.equal(analysis.notes.length, 2);
    assert.deepEqual(
      analysis.notes.map((n) => n.kind),
      ["existing", "existing"],
    );
    assert.deepEqual(
      analysis.notes.map((n) => n.n),
      [1, 2],
    );
    assert.deepEqual(
      analysis.notes.map((n) => n.url),
      [NOTE_A, NOTE_B],
    );
  });

  await t.test("places each marker where it sits in the text", () => {
    const para = analysis.bodyParas[0];
    assert.deepEqual(
      para.parts.map((x) => x.type),
      ["text", "note", "text"],
    );
    assert.equal(para.parts[0].value, "Growth was strong");
    assert.equal(para.parts[1].n, 1);
    assert.equal(para.parts[2].value, " last year.");
  });

  await t.test("reports a note that nothing points to", async () => {
    const buf = await buildDocx({
      footnotesXml: h.footnotesPart([
        { id: 1, text: "Referenced." },
        { id: 2, text: "Orphaned by an edit." },
      ]),
      body: [
        p("A Report", { style: "Title" }),
        p(["Text", { footnoteRef: 1 }, "."]),
      ],
    });
    const a = await engine.analyzeDocx(buf);
    assert.equal(a.notes.length, 1);
    assert.equal(a.orphanFootnotes.length, 1);
    assert.match(a.orphanFootnotes[0].text, /Orphaned/);
  });

  await t.test("reads a plain-text note as well as a linked one", async () => {
    const buf = await buildDocx({
      footnotesXml: h.footnotesPart([{ id: 1, text: "Interview, 4 March 2026." }]),
      body: [p("A Report", { style: "Title" }), p(["Text", { footnoteRef: 1 }, "."])],
    });
    const a = await engine.analyzeDocx(buf);
    assert.equal(a.notes[0].text, "Interview, 4 March 2026.");
    assert.equal(a.notes[0].url, null);
  });

  await t.test("leaves the file untouched when converted anyway", async () => {
    // Nothing to do, so nothing should change — the interface disables the
    // download here, but the engine must be safe if it is called regardless.
    const r = await convert(buffer);
    assert.equal(r.summary.footnotesCreated, 0);
    assert.equal(r.summary.existingFootnotesRespaced, 0);
    assert.equal(r.out.footnotes.length, 2);
    assert.equal(r.out.paragraphs[1], "Growth was strong{1} last year.");
  });
});

// ===========================================================================
test("a document with both footnotes and references", async (t) => {
  const buffer = await mixedDoc();
  const analysis = await engine.analyzeDocx(buffer);
  const converted = await convert(buffer);

  await t.test("is not treated as already converted", () => {
    assert.equal(analysis.alreadyConverted, false);
    assert.equal(analysis.citationCount, 2);
    assert.equal(analysis.existingFootnoteCount, 1);
  });

  await t.test("numbers existing and new notes together, in reading order", () => {
    // Word numbers by position, so the first citation here is note 2, not
    // note 1 — the preview has to say the same thing the output will.
    assert.deepEqual(
      analysis.notes.map((n) => [n.n, n.kind]),
      [
        [1, "existing"],
        [2, "new"],
        [3, "new"],
      ],
    );
  });

  await t.test("keeps the existing note exactly as it was", () => {
    assert.equal(analysis.notes[0].text, "A hand-written note.");
    const kept = converted.out.footnotes.find((f) => f.id === "4");
    assert.ok(kept, "the original footnote should survive");
    assert.match(kept.text, /A hand-written note\./);
  });

  await t.test("converts the remaining references to footnotes", () => {
    assert.equal(converted.summary.footnotesCreated, 2);
    assert.equal(converted.summary.referencesConverted, 2);
    assert.equal(converted.out.paragraphs[2], "Growth was strong{5} last year.");
    assert.equal(converted.out.paragraphs[3], "Later still{6} it held.");
  });

  await t.test("gives new footnotes ids above the existing one", () => {
    // Reusing id 4 would silently overwrite the hand-written note.
    const ids = converted.out.footnotes.map((f) => Number(f.id)).sort((a, b) => a - b);
    assert.deepEqual(ids, [4, 5, 6]);
  });

  await t.test("collapses the reference list, as usual", () => {
    assert.equal(converted.summary.referenceSectionCollapsed, true);
    assert.ok(converted.out.paragraphs.includes(engine.PLACEHOLDER_TEXT));
  });

  await t.test("keeps the existing marker where it was", () => {
    assert.equal(converted.out.paragraphs[1], "An aside{4} opens the argument.");
  });
});

// ===========================================================================
test("spacing between footnotes", async (t) => {
  await t.test("ends each new footnote with a blank line", async () => {
    // Word packs footnotes flush together. With a long URL each, they run into
    // one unreadable block, so every note gets a line of air after it.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, https://example.com/smith."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.out.footnotes[0].paragraphCount, 2);
    assert.equal(r.out.footnotes[0].trailingBlanks, 1);
  });

  await t.test("uses the footnote style for the blank line", async () => {
    // A blank paragraph in the body style is a full-size gap, not a small one.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, https://example.com/smith."),
      ],
    });
    const r = await convert(buf);
    const xml = await r.out.zip.file("word/footnotes.xml").async("string");
    assert.match(xml, /<w:p><w:pPr><w:pStyle w:val="FootnoteText"\/><\/w:pPr><\/w:p>/);
  });

  await t.test("gives the document's own footnotes the same spacing", async () => {
    // Otherwise a document that had one note and gained two shows two
    // different spacings on the same page.
    const r = await convert(await mixedDoc());
    assert.equal(r.summary.existingFootnotesRespaced, 1);
    r.out.footnotes.forEach((fn) => {
      assert.equal(fn.trailingBlanks, 1, "footnote " + fn.id);
    });
  });

  await t.test("does not add a second blank line to a note that has one", async () => {
    // Converting a document twice must not keep growing the gaps.
    const buf = await buildDocx({
      footnotesXml: h.footnotesPart([
        { id: 4, text: "Already spaced.", trailingBlank: true },
      ]),
      body: [
        p("A Report", { style: "Title" }),
        p(["An aside", { footnoteRef: 4 }, "."]),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, https://example.com/smith."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.summary.existingFootnotesRespaced, 0);
    r.out.footnotes.forEach((fn) => assert.equal(fn.trailingBlanks, 1, "footnote " + fn.id));
  });

  await t.test("leaves spacing alone when nothing is converted", async () => {
    // A document nobody asked to change isn't reformatted.
    const buf = await buildDocx({
      footnotesXml: h.footnotesPart([{ id: 1, text: "Untouched." }]),
      body: [p("A Report", { style: "Title" }), p(["Text", { footnoteRef: 1 }, "."])],
    });
    const r = await convert(buf);
    assert.equal(r.summary.existingFootnotesRespaced, 0);
    assert.equal(r.out.footnotes[0].trailingBlanks, 0);
  });

  await t.test("stays stable across repeated conversions", async () => {
    const first = await convert(await mixedDoc());
    const second = await convert(first.buffer);
    // The second pass finds nothing to convert, so the spacing it inherited
    // is exactly the spacing it keeps.
    assert.equal(second.summary.footnotesCreated, 0);
    second.out.footnotes.forEach((fn) => assert.equal(fn.trailingBlanks, 1));
  });
});

// ===========================================================================
test("markers interleaved with citations in one paragraph", async (t) => {
  await t.test("keeps both, in the order they appear", async () => {
    const buf = await buildDocx({
      footnotesXml: h.footnotesPart([{ id: 4, text: "An aside." }]),
      body: [
        p("A Report", { style: "Title" }),
        p([
          "Growth was strong",
          { footnoteRef: 4 },
          " and confirmed later (Smith 2020) in review.",
        ]),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, https://example.com/smith."),
      ],
    });
    const a = await engine.analyzeDocx(buf);
    assert.deepEqual(
      a.bodyParas[0].parts.map((x) => x.type),
      ["text", "note", "text", "cite", "text"],
    );
    assert.equal(a.bodyParas[0].parts[1].n, 1);
    assert.equal(a.bodyParas[0].parts[3].n, 2);

    const r = await convert(buf);
    assert.equal(
      r.out.paragraphs[1],
      "Growth was strong{4} and confirmed later{5} in review.",
    );
  });

  await t.test("keeps an existing marker that follows a converted citation", async () => {
    const buf = await buildDocx({
      footnotesXml: h.footnotesPart([{ id: 4, text: "An aside." }]),
      body: [
        p("A Report", { style: "Title" }),
        p(["Growth was strong (Smith 2020) and later", { footnoteRef: 4 }, " confirmed."]),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, https://example.com/smith."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.out.paragraphs[1], "Growth was strong{5} and later{4} confirmed.");
    // Both survive — the citation replacement must not disturb the marker
    // sitting further along the same paragraph.
    assert.equal(r.out.footnotes.length, 2);
  });
});
