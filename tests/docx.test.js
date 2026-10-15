"use strict";

/**
 * Tests for the OOXML layer (miwim.docx.js).
 *
 * These build real .docx packages in memory, run them through the converter,
 * and read the result back — so they cover the parts that only go wrong once
 * XML is involved: runs split mid-citation, hyperlink wrappers, paragraphs
 * nested in tables, tables of contents, styles and relationships.
 */

const test = require("node:test");
const assert = require("node:assert/strict");

const h = require("./helpers/docx-fixture.js");
const { p, table, sdt, buildDocx, convert, engine, core } = h;

// A minimal document: title, one citing paragraph, a reference list.
function simpleDoc(extra) {
  return buildDocx(
    Object.assign(
      {
        body: [
          p("A Report", { style: "Title" }),
          p("Growth was strong (Smith 2020) last year."),
          p("References", { style: "Heading1" }),
          p("Smith, J 2020, A Book, Press, Sydney, https://example.com/book."),
        ],
      },
      extra || {},
    ),
  );
}

// ===========================================================================
test("reading paragraphs out of a document", async (t) => {
  await t.test("finds paragraphs nested inside a table", async () => {
    // A body laid out in a single-cell table is common in report templates.
    // A converter that only walks direct children of w:body finds nothing.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        table([p("Growth was strong (Smith 2020).")]),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.summary.footnotesCreated, 1);
    assert.match(r.out.paragraphs.join("\n"), /Growth was strong\{1\}\./);
  });

  await t.test("stitches text back together across split runs", async () => {
    // Word splits runs at spell-check and revision boundaries, so a citation
    // routinely spans three or four of them. Matching has to see one string.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p(["Growth was strong (Smi", "th ", "2020) last year."]),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.summary.footnotesCreated, 1);
    assert.equal(r.out.paragraphs[1], "Growth was strong{1} last year.");
  });

  await t.test("reads text inside a hyperlink wrapper", async () => {
    // A reference's URL is usually a w:hyperlink, one level below the runs a
    // naive walk inspects. Missing it corrupts both the text and the URL.
    const buf = await buildDocx({
      rels: [{ id: "rId9", target: "https://example.com/report" }],
      body: [
        p("A Report", { style: "Title" }),
        p("As reported (ASD 2025)."),
        p("References", { style: "Heading1" }),
        p([
          { text: "ASD 2025, Annual Report, Canberra, viewed 1 Jan 2026, " },
          { text: "https://example.com/report", rid: "rId9" },
          { text: "." },
        ]),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
    assert.match(analysis.refs[0].text, /viewed 1 Jan 2026, https:\/\/example\.com\/report\./);
    assert.equal(analysis.refs[0].url, "https://example.com/report");
  });

  await t.test("prefers the hyperlink target over the visible link text", async () => {
    // When the clickable text is a title rather than an address, the
    // relationship target is the only place the real URL exists.
    const buf = await buildDocx({
      rels: [{ id: "rId9", target: "https://example.com/actual-target" }],
      body: [
        p("A Report", { style: "Title" }),
        p("As reported (ASD 2025)."),
        p("References", { style: "Heading1" }),
        p([
          { text: "ASD 2025, Annual Report, Canberra, " },
          { text: "View source", rid: "rId9" },
          { text: "." },
        ]),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs[0].url, "https://example.com/actual-target");
  });

  await t.test("ignores a generated table of contents", async () => {
    // A TOC duplicates every heading, "Bibliography" included. Treated as the
    // real thing, the reference list becomes a single page-number entry.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        sdt([
          p("Introduction\t1", { style: "TOC1" }),
          p("Bibliography\t9", { style: "TOC1" }),
        ]),
        p("Growth was strong (Smith 2020)."),
        p("Bibliography", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
    assert.match(analysis.refs[0].text, /Smith, J 2020/);
  });

  await t.test("ignores a TOC built from field codes rather than styles", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p([{ instrText: ' TOC \\o "1-3" \\h ', text: "Bibliography\t9" }]),
        p("Growth was strong (Smith 2020)."),
        p("Bibliography", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
  });
});

// ===========================================================================
test("locating the reference section", async (t) => {
  await t.test("stops the reference list at the next section heading", async () => {
    // Without a boundary, every remaining paragraph is read as a reference.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
        p("Appendix A", { style: "Heading1" }),
        p("Survey instrument used in 2019 fieldwork."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
  });

  await t.test("stops at an unstyled but recognisable trailing heading", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
        p("Declaration of AI Use"),
        p("Generative AI was used to check grammar in 2026."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
  });

  await t.test("prefers a styled heading over a plain-text mention", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("References"), // a stray line, not the real section
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
    assert.equal(analysis.citationCount, 1);
  });

  await t.test("skips blank paragraphs between reference entries", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
        p(""),
        p("Jones, K 2021, Another Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 2);
  });

  await t.test("reads a bibliography subdivided by its own subheadings", async () => {
    // ELL_917's shape. A Heading 2 group label one paragraph below the
    // Heading 1 "Bibliography" used to end the section immediately, so the
    // list read as empty and the interface reported "a references heading
    // was found, but nothing followed it" — with 26 entries sitting in plain
    // sight below it. A heading deeper than the reference heading is a
    // subdivision of the list, not the section after it.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("The rule applies (Jones 2021)."),
        p("Bibliography", { style: "Heading1" }),
        p("Legislation and Cases", { style: "Heading2" }),
        p("Smith, J 2020, A Book, Press."),
        p("Secondary Sources", { style: "Heading2" }),
        p("Jones, K 2021, Another Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 2);
    assert.equal(analysis.citationCount, 2);
  });

  await t.test("does not count a group label as a reference entry", async () => {
    // The subheadings are inside the section but they are not sources.
    // Counting them would report phantom never-cited references.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("Bibliography", { style: "Heading1" }),
        p("Secondary Sources", { style: "Heading2" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
    assert.equal(analysis.uncitedRefs.length, 0);
  });

  await t.test("still ends the list at a same-level heading", async () => {
    // The level comparison must not swallow the rest of the document.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("Bibliography", { style: "Heading1" }),
        p("Secondary Sources", { style: "Heading2" }),
        p("Smith, J 2020, A Book, Press."),
        p("Conclusion", { style: "Heading1" }),
        p("The position in 2019 was different."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
  });

  await t.test("ends the list at a deeper heading that names a real section", async () => {
    // "Appendix A" styled Heading 2 under a Heading 1 "Bibliography" is the
    // appendix, not a group of references. Text wins over level here.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("Bibliography", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
        p("Appendix A", { style: "Heading2" }),
        p("Survey instrument used in 2019 fieldwork."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
  });

  await t.test("keeps old behaviour when the reference heading is unstyled", async () => {
    // With no level to compare against, guessing wrong would swallow the
    // whole rest of the document, so any styled heading still ends the list.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("Bibliography"),
        p("Smith, J 2020, A Book, Press."),
        p("Appendix", { style: "Heading2" }),
        p("Survey instrument used in 2019 fieldwork."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.refs.length, 1);
  });
});

// ===========================================================================
test("converting a document", async (t) => {
  await t.test("replaces the citation with a real footnote reference", async () => {
    const r = await convert(await simpleDoc());
    assert.equal(r.out.paragraphs[1], "Growth was strong{1} last year.");
    assert.equal(r.out.footnotes.length, 1);
    assert.equal(r.out.footnotes[0].url, "https://example.com/book");
    // w:footnoteRef is what makes Word number the footnote itself.
    assert.equal(r.out.footnotes[0].hasFootnoteRef, true);
    assert.equal(r.out.footnotes[0].styleId, "FootnoteText");
  });

  await t.test("numbers footnotes in reading order, not reference-list order", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("First (Wu 2022)."),
        p("Second (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, https://example.com/smith."),
        p("Wu, L 2022, Third Book, https://example.com/wu."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.out.paragraphs[1], "First{1}.");
    assert.equal(r.out.paragraphs[2], "Second{2}.");
    assert.equal(r.out.footnotes[0].url, "https://example.com/wu");
    assert.equal(r.out.footnotes[1].url, "https://example.com/smith");
  });

  await t.test("gives each occurrence of a repeated source its own footnote", async () => {
    // Word numbers footnotes by position, so a source cited three times needs
    // three marks — sharing one would put the same number in three places.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("One (ASD 2025)."),
        p("Two (ASD 2025)."),
        p("Three (ASD 2025)."),
        p("References", { style: "Heading1" }),
        p("ASD 2025, Annual Report, https://example.com/asd."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.summary.footnotesCreated, 3);
    assert.deepEqual(r.out.paragraphs.slice(1, 4), ["One{1}.", "Two{2}.", "Three{3}."]);
    r.out.footnotes.forEach((fn) => assert.equal(fn.url, "https://example.com/asd"));
  });

  await t.test("reuses one relationship for a repeated URL", async () => {
    // Duplicate relationship ids make Word refuse to open the file, and a
    // fresh id per occurrence bloats the package for no benefit.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("One (ASD 2025). Two (ASD 2025)."),
        p("References", { style: "Heading1" }),
        p("ASD 2025, Annual Report, https://example.com/asd."),
      ],
    });
    const r = await convert(buf);
    assert.equal(Object.keys(r.out.footnoteRels).length, 1);
  });

  await t.test("places several markers side by side for a shared bracket", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Both agree (Smith 2020; Jones 2021)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, https://example.com/smith."),
        p("Jones, K 2021, Another Book, https://example.com/jones."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.out.paragraphs[1], "Both agree{1},{2}.");
    assert.equal(r.out.footnotes[0].url, "https://example.com/smith");
    assert.equal(r.out.footnotes[1].url, "https://example.com/jones");
  });

  await t.test("keeps offsets correct with several citations in one paragraph", async () => {
    // Each replacement changes the paragraph's length, so edits are applied
    // right-to-left. Left-to-right, the later spans land in the wrong place.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("First (Smith 2020), then (Jones 2021), finally (Wu 2022) closes it."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, https://example.com/smith."),
        p("Jones, K 2021, Another Book, https://example.com/jones."),
        p("Wu, L 2022, Third Book, https://example.com/wu."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.out.paragraphs[1], "First{1}, then{2}, finally{3} closes it.");
  });

  await t.test("converts a citation split across runs without losing text", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p(["Before ", "(Smith ", "2020)", " after."]),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.out.paragraphs[1], "Before{1} after.");
  });

  await t.test("appends to an existing footnote instead of overwriting it", async () => {
    // A document may already contain hand-written footnotes; their ids must
    // survive and the new ones must start above them.
    const existing =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:footnotes xmlns:w="' + h.W_NS + '" xmlns:r="' + h.R_NS + '">' +
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="7"><w:p><w:r><w:t>An earlier note by hand.</w:t></w:r></w:p></w:footnote>' +
      "</w:footnotes>";
    const r = await convert(await simpleDoc({ footnotesXml: existing }));
    const texts = r.out.footnotes.map((f) => f.text);
    assert.ok(texts.some((t2) => /earlier note by hand/.test(t2)));
    // The new footnote must not collide with id 7.
    assert.equal(r.out.paragraphs[1], "Growth was strong{8} last year.");
  });

  await t.test("leaves a document with no citations completely untouched", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Nothing is cited here at all."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.summary.footnotesCreated, 0);
    assert.equal(r.out.footnotes.length, 0);
    // The uncited reference stays in the list — deleting someone's source
    // because no citation matched it would be a silent data loss.
    assert.match(r.out.paragraphs[3], /Smith, J 2020/);
  });
});

// ===========================================================================
test("the reference list after conversion", async (t) => {
  await t.test("collapses to a placeholder when every entry became a footnote", async () => {
    const r = await convert(await simpleDoc());
    assert.equal(r.summary.referenceSectionCollapsed, true);
    assert.equal(r.out.paragraphs[3], engine.PLACEHOLDER_TEXT);
    // The heading itself stays, so the document still has its section.
    assert.equal(r.out.paragraphs[2], "References");
  });

  await t.test("keeps uncited entries rather than deleting them", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
        p("Nguyen, T 2019, An Uncited Book, Press."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.summary.referenceSectionCollapsed, false);
    assert.equal(r.summary.referencesKept, 1);
    const joined = r.out.paragraphs.join("\n");
    assert.doesNotMatch(joined, /Smith, J 2020/); // became a footnote
    assert.match(joined, /Nguyen, T 2019/); // never cited, so kept
  });

  await t.test("strips list numbering from the placeholder", async () => {
    // A numbered reference list would otherwise leave the placeholder sitting
    // under a stray "1." that no longer means anything.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press.", { numId: 3 }),
      ],
    });
    const r = await convert(buf);
    const numPr = r.out.doc.getElementsByTagNameNS(h.W_NS, "numPr");
    assert.equal(numPr.length, 0);
  });

  await t.test("removes the old hyperlink along with the entry text", async () => {
    // Clearing only w:r elements strands the linked URL, which then sits next
    // to the placeholder text as an orphan.
    const buf = await buildDocx({
      rels: [{ id: "rId9", target: "https://example.com/report" }],
      body: [
        p("A Report", { style: "Title" }),
        p("As reported (ASD 2025)."),
        p("References", { style: "Heading1" }),
        p([
          { text: "ASD 2025, Annual Report, " },
          { text: "https://example.com/report", rid: "rId9" },
          { text: "." },
        ]),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.out.paragraphs[3], engine.PLACEHOLDER_TEXT);
  });

  await t.test("collapses a subdivided list, group labels included", async () => {
    // Once every entry is a footnote, leaving "Legislation and Cases" and
    // "Secondary Sources" above a single placeholder line is nonsense — the
    // groups they label no longer exist.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("The rule applies (Jones 2021)."),
        p("Bibliography", { style: "Heading1" }),
        p("Legislation and Cases", { style: "Heading2" }),
        p("Smith, J 2020, A Book, Press."),
        p("Secondary Sources", { style: "Heading2" }),
        p("Jones, K 2021, Another Book, Press."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.summary.referenceSectionCollapsed, true);
    const joined = r.out.paragraphs.join("\n");
    assert.match(joined, /Bibliography/); // the section heading stays
    assert.doesNotMatch(joined, /Legislation and Cases/);
    assert.doesNotMatch(joined, /Secondary Sources/);
    assert.match(joined, /\[Insert citations/);
  });

  await t.test("keeps a group label that still has entries under it", async () => {
    // Partial conversion: "Secondary Sources" keeps its uncited entry, so it
    // keeps its label. "Legislation and Cases" is emptied, so its label goes
    // rather than heading an empty group.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("Bibliography", { style: "Heading1" }),
        p("Legislation and Cases", { style: "Heading2" }),
        p("Smith, J 2020, A Book, Press."),
        p("Secondary Sources", { style: "Heading2" }),
        p("Nguyen, T 2019, An Uncited Book, Press."),
      ],
    });
    const r = await convert(buf);
    assert.equal(r.summary.referenceSectionCollapsed, false);
    const joined = r.out.paragraphs.join("\n");
    assert.doesNotMatch(joined, /Legislation and Cases/);
    assert.match(joined, /Secondary Sources/);
    assert.match(joined, /Nguyen, T 2019/);
  });
});

// ===========================================================================
test("package integrity", async (t) => {
  await t.test("adds the footnote style definitions when missing", async () => {
    // Word only writes these once a document actually has a footnote; without
    // them the footnote text renders as ordinary body copy.
    const r = await convert(await simpleDoc());
    ["FootnoteText", "FootnoteTextChar", "FootnoteReference", "Hyperlink"].forEach((id) => {
      assert.ok(r.out.styleIds.includes(id), id);
    });
  });

  await t.test("does not duplicate style definitions that already exist", async () => {
    const styles =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="' + h.W_NS + '">' +
      '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/></w:style>' +
      "</w:styles>";
    const r = await convert(await simpleDoc({ stylesXml: styles }));
    const count = r.out.styleIds.filter((id) => id === "FootnoteText").length;
    assert.equal(count, 1);
  });

  await t.test("declares the footnotes part in [Content_Types].xml", async () => {
    const r = await convert(await simpleDoc({ contentTypes: false }));
    assert.match(r.out.contentTypes, /PartName="\/word\/footnotes\.xml"/);
  });

  await t.test("adds the footnotes relationship when the document lacks one", async () => {
    // A .docx saved without any footnotes has no such relationship, and Word
    // will not display footnotes it cannot resolve.
    const r = await convert(await simpleDoc({ footnotesXml: null }));
    const types = Object.keys(r.out.documentRels).map((k) => r.out.documentRels[k].type);
    assert.ok(types.some((t2) => /\/footnotes$/.test(t2)));
    assert.equal(r.out.footnotes.length, 1);
  });

  await t.test("writes external hyperlink relationships with TargetMode", async () => {
    // Without TargetMode="External", Word reads the URL as a path inside the
    // package and the footnote link is dead.
    const r = await convert(await simpleDoc());
    const relsXml = await r.out.zip
      .file("word/_rels/footnotes.xml.rels")
      .async("string");
    assert.match(relsXml, /TargetMode="External"/);
  });

  await t.test("produces XML that parses cleanly on a second pass", async () => {
    // The strongest cheap check that nothing was corrupted: feed the output
    // back into the converter.
    const r = await convert(await simpleDoc());
    const again = await engine.analyzeDocx(r.buffer);
    assert.equal(again.docTitle, "A Report");
    assert.equal(again.citationCount, 0); // already converted, nothing left to do
  });

  await t.test("leaves unrelated parts of the package alone", async () => {
    const buf = await simpleDoc();
    const before = await h.JSZip.loadAsync(buf);
    const r = await convert(buf);
    const beforeNames = Object.keys(before.files).sort();
    const afterNames = Object.keys(r.out.zip.files).sort();
    beforeNames.forEach((name) => assert.ok(afterNames.includes(name), name));
  });
});

// ===========================================================================
test("documents that are missing things", async (t) => {
  await t.test("reports a document with no reference section", async () => {
    const buf = await buildDocx({
      body: [p("A Report", { style: "Title" }), p("Growth was strong (Smith 2020).")],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.hasRefsSection, false);
    assert.equal(analysis.refs.length, 0);
    // The body still renders, so the person sees their document rather than
    // an error page.
    assert.equal(analysis.bodyParas.length, 1);
  });

  await t.test("reports a reference heading with nothing under it", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("Bibliography", { style: "Heading1" }),
        p(""),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.hasRefsSection, true);
    assert.equal(analysis.refs.length, 0);
  });

  await t.test("writes the full reference into the footnote when there is no URL", async () => {
    // An empty footnote is worse than a long one.
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press, Sydney."),
      ],
    });
    const r = await convert(buf);
    assert.match(r.out.footnotes[0].text, /Smith, J 2020, A Book, Press, Sydney\./);
    assert.equal(r.out.footnotes[0].url, null);
  });

  await t.test("uses a URL the person supplied for a reference that had none", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press, Sydney."),
      ],
    });
    const r = await convert(buf, (refs) => {
      refs[0].url = "https://example.com/added-by-hand";
    });
    assert.equal(r.out.footnotes[0].url, "https://example.com/added-by-hand");
  });

  await t.test("rejects a file that is not a .docx", async () => {
    const zip = new h.JSZip();
    zip.file("hello.txt", "not a document");
    const buf = await zip.generateAsync({ type: "nodebuffer" });
    await assert.rejects(() => engine.analyzeDocx(buf), /valid \.docx/);
  });

  await t.test("rejects a .docx whose XML is malformed", async () => {
    const zip = new h.JSZip();
    zip.file("word/document.xml", "<w:document><w:body><w:p>");
    const buf = await zip.generateAsync({ type: "nodebuffer" });
    await assert.rejects(() => engine.analyzeDocx(buf));
  });

  await t.test("refuses to construct an engine without its dependencies", () => {
    const createDocxEngine = require("../miwim.docx.js");
    assert.throws(() => createDocxEngine({}), /needs/);
    assert.throws(() => createDocxEngine(), /needs/);
  });
});

// ===========================================================================
test("what the preview reports", async (t) => {
  await t.test("counts citations and links them to their reference", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("One (Smith 2020) and two (Smith 2020)."),
        p("Also (Jones 2021)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
        p("Jones, K 2021, Another Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.citationCount, 3);
    assert.equal(analysis.refs[0].citationCount, 2);
    assert.equal(analysis.refs[1].citationCount, 1);
  });

  await t.test("lists references that are never cited", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Smith 2020)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
        p("Nguyen, T 2019, An Uncited Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.uncitedRefs.length, 1);
    assert.match(analysis.uncitedRefs[0].text, /Nguyen/);
  });

  await t.test("lists citations that matched no reference", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Growth was strong (Nguyen 2019)."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.unmatched.length, 1);
    assert.equal(analysis.unmatched[0].text, "(Nguyen 2019)");
  });

  await t.test("keeps heading levels so the preview reads like the document", async () => {
    const buf = await buildDocx({
      body: [
        p("A Report", { style: "Title" }),
        p("Chapter 1", { style: "Heading1" }),
        p("Some prose."),
        p("References", { style: "Heading1" }),
        p("Smith, J 2020, A Book, Press."),
      ],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.bodyParas[0].heading, 1);
    assert.equal(analysis.bodyParas[1].heading, null);
  });

  await t.test("treats the first non-empty paragraph as the title", async () => {
    const buf = await buildDocx({
      body: [p(""), p("A Report", { style: "Title" }), p("Some prose.")],
    });
    const analysis = await engine.analyzeDocx(buf);
    assert.equal(analysis.docTitle, "A Report");
    assert.equal(analysis.bodyParas.length, 1);
  });
});

// ===========================================================================
test("OOXML helpers in isolation", async (t) => {
  await t.test("maps run offsets onto the paragraph's text", () => {
    const pEl = h.parseParagraph(
      "<w:p><w:r><w:t>abc</w:t></w:r><w:r><w:t>defg</w:t></w:r></w:p>",
    );
    const entries = engine.collectRunEntries(pEl);
    assert.deepEqual(
      entries.map((e) => [e.start, e.end, e.text]),
      [
        [0, 3, "abc"],
        [3, 7, "defg"],
      ],
    );
    assert.equal(engine.paragraphText(pEl), "abcdefg");
  });

  await t.test("ignores runs that only describe formatting", () => {
    // A w:pPr/w:rPr run holds paragraph-mark properties, not visible text.
    const pEl = h.parseParagraph(
      "<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>visible</w:t></w:r></w:p>",
    );
    assert.equal(engine.paragraphText(pEl), "visible");
  });

  await t.test("ignores text marked as a tracked deletion", () => {
    const pEl = h.parseParagraph(
      '<w:p><w:r><w:t>kept</w:t></w:r><w:del><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>',
    );
    assert.equal(engine.paragraphText(pEl), "kept");
  });

  await t.test("reads a paragraph's style id", () => {
    const pEl = h.parseParagraph(
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
    );
    assert.equal(engine.paragraphStyleId(pEl), "Heading2");
    assert.equal(engine.headingLevel(pEl), 2);
  });

  await t.test("finds the highest existing footnote id", () => {
    const doc = engine.parseXml(
      '<w:footnotes xmlns:w="' + h.W_NS + '">' +
        '<w:footnote w:id="-1"/><w:footnote w:id="0"/><w:footnote w:id="12"/>' +
        "</w:footnotes>",
    );
    assert.equal(engine.getMaxFootnoteId(doc), 12);
  });

  await t.test("finds the highest existing relationship id", () => {
    const doc = engine.parseXml(
      '<Relationships xmlns="' + h.REL_NS + '">' +
        '<Relationship Id="rId1"/><Relationship Id="rId17"/><Relationship Id="rId3"/>' +
        "</Relationships>",
    );
    assert.equal(engine.getMaxRelId(doc), 17);
  });
});
