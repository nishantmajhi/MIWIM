"use strict";

/**
 * Pinpoint locators, and the rule that a bracket must add up.
 *
 * A footnote marker replaces the whole bracket, so everything the bracket held
 * has to be accounted for — the author and the year become the note, a pinpoint
 * is carried into the note, and anything else means the bracket is left exactly
 * as written. Before this suite existed, a bracket that continued past the year
 * was truncated to the year and the rest was deleted from the document without
 * a word: "(Smith 2020, p. 44)" lost its page number, and
 * "(Smith 2020, Jones 2021)" lost a whole second source.
 *
 * Every check here was written before the fix and confirmed failing against the
 * code as it stood at commit 25a70b9.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("path");

const core = require(path.join(__dirname, "..", "miwim.core.js"));
const H = require(path.join(__dirname, "helpers", "docx-fixture.js"));
const { p, buildDocx, readDocx, engine } = H;

let refCounter = 0;
function refs(list) {
  refCounter = 0;
  return list.map(function (text) {
    refCounter++;
    return Object.assign(
      { n: refCounter, id: "c" + refCounter, text: text },
      core.parseReferenceEntry(text),
    );
  });
}

/**
 * Every character of the original that is not inside a replaced span. Used to
 * assert the headline property: a conversion may move text and it may replace
 * a citation, but it must never make words disappear.
 */
function survivingText(text, references) {
  const found = core.findCitationGroups(text, references);
  let out = "";
  let cursor = 0;
  found.groups.forEach(function (g) {
    out += text.slice(cursor, g.start);
    cursor = g.end;
  });
  return out + text.slice(cursor);
}

// ===========================================================================
test("1. a bracket is never converted with words left over", async (t) => {
  const list = refs([
    "Smith, J 2020, A Book, Press, Sydney.",
    "Smith, J 2021, A Later Book, Press, Sydney.",
    "Jones, K 2021, Another Book, Press, Melbourne.",
  ]);

  // The rule the README states, and the one this file exists to enforce: a
  // bracket is only converted when every part of it is accounted for. Before
  // the fix each of these converted and silently dropped the leftover words.
  const mustNotLoseText = [
    "(Smith 2020, p. 44)",
    "(Smith 2020, pp. 44-46)",
    "(Smith 2020, para. 7)",
    "(Smith 2020, ch. 3)",
    "(Smith 2020, emphasis added)",
    "(Smith 2020, my translation)",
    "(Smith 2020, Jones 2021)",
    "(Smith 2020, 2021)",
    "(Smith 2020, as cited in Jones 2021)",
    "(Smith 2020, but see Jones 2021)",
    "(Smith 2020, Table 3)",
    "(Smith 2020 and Jones 2021)",
  ];

  await t.test("nothing after the year is dropped from any of them", () => {
    // The author and the year are the two things a footnote marker is *meant*
    // to absorb — they reappear in the note. Everything else in the bracket is
    // the author's own writing, and this is the property that was broken: for
    // every one of these, the text following the year was deleted.
    mustNotLoseText.forEach(function (bracket) {
      const text = "Growth was strong " + bracket + " last year.";
      const found = core.findCitationGroups(text, list);
      const inner = bracket.slice(1, -1);
      const yearMatch = core.findYearTokens(inner)[0];
      const tail = yearMatch ? inner.slice(yearMatch.end) : "";
      const words = tail.split(/[\s,;:]+/).filter(function (w) {
        return w && !/^[.,;:]+$/.test(w);
      });
      if (!words.length) return;

      if (!found.groups.length) {
        // Not converted: the bracket stands exactly as written.
        assert.equal(survivingText(text, list), text, bracket);
        return;
      }
      // Converted: every word after the year is carried in a locator.
      const carried = found.groups
        .map(function (g) {
          return (g.locators || []).join(" ");
        })
        .join(" ");
      words.forEach(function (w) {
        assert.ok(
          carried.indexOf(w) !== -1,
          "lost " + JSON.stringify(w) + " from " + bracket,
        );
      });
    });
  });

  await t.test("a page locator is kept, and the citation still converts", () => {
    const found = core.findCitationGroups("Growth rose (Smith 2020, p. 44).", list);
    assert.equal(found.groups.length, 1, "the citation should still convert");
    assert.deepEqual(found.groups[0].refIds, ["c1"]);
    assert.deepEqual(found.groups[0].locators, ["p. 44"]);
  });

  await t.test("every locator abbreviation this archive's documents use", () => {
    const locators = [
      ["(Smith 2020, p. 44)", "p. 44"],
      ["(Smith 2020, pp. 44-46)", "pp. 44-46"],
      ["(Smith 2020, p 44)", "p 44"],
      ["(Smith 2020, page 44)", "page 44"],
      ["(Smith 2020, pages 44-46)", "pages 44-46"],
      ["(Smith 2020, ch. 3)", "ch. 3"],
      ["(Smith 2020, chapter 3)", "chapter 3"],
      ["(Smith 2020, s. 12)", "s. 12"],
      ["(Smith 2020, sec. 12)", "sec. 12"],
      ["(Smith 2020, para. 7)", "para. 7"],
      ["(Smith 2020, paras. 7-9)", "paras. 7-9"],
      ["(Smith 2020, vol. 2)", "vol. 2"],
      ["(Smith 2020, no. 4)", "no. 4"],
      ["(Smith 2020, fig. 2)", "fig. 2"],
      ["(Smith 2020, art. 5)", "art. 5"],
      ["(Smith 2020, table 3)", "table 3"],
      ["(Smith 2020, vol. 2, p. 9)", "vol. 2, p. 9"],
      ["(Smith 2020: 44)", "44"],
      ["(Smith 2020, 44)", "44"],
      ["(Smith 2020, 44-46)", "44-46"],
    ];
    locators.forEach(function (pair) {
      const found = core.findCitationGroups("Text " + pair[0] + " end.", list);
      assert.equal(found.groups.length, 1, pair[0] + " should convert");
      assert.deepEqual(found.groups[0].locators, [pair[1]], pair[0]);
    });
  });

  await t.test("residue that is not a locator leaves the bracket alone", () => {
    // Left as written and reported, which is what the README promises and what
    // the semicolon case already did. Converting would delete the words.
    [
      "(Smith 2020, emphasis added)",
      "(Smith 2020, my translation)",
      "(Smith 2020, Jones 2021)",
      "(Smith 2020, 2021)",
      "(Smith 2020, as cited in Jones 2021)",
      "(Smith 2020, but see Jones 2021)",
    ].forEach(function (bracket) {
      const found = core.findCitationGroups("Text " + bracket + " end.", list);
      assert.equal(found.groups.length, 0, bracket + " should not convert");
      assert.equal(found.unmatched.length, 1, bracket + " should be reported");
      assert.equal(found.unmatched[0].text, bracket);
    });
  });

  await t.test("a year-shaped bare number is not read as a page number", () => {
    // "(Smith 2020, 2021)" cites two works, not page 2021 of one. Reading it
    // as a locator would attach the wrong note and lose the second citation.
    const found = core.findCitationGroups("Both (Smith 2020, 2021) agree.", list);
    assert.equal(found.groups.length, 0);
    assert.equal(found.unmatched.length, 1);
  });

  await t.test("each citation in a shared bracket keeps its own locator", () => {
    const found = core.findCitationGroups(
      "Both (Smith 2020, p. 44; Jones 2021, p. 8) agree.",
      list,
    );
    assert.equal(found.groups.length, 1);
    assert.deepEqual(found.groups[0].refIds, ["c1", "c3"]);
    assert.deepEqual(found.groups[0].locators, ["p. 44", "p. 8"]);
  });

  await t.test("one reference cited twice at different pages is two notes", () => {
    const found = core.findCitationGroups(
      "See (Smith 2020, p. 4; Smith 2020, p. 90).",
      list,
    );
    assert.equal(found.groups.length, 1);
    assert.deepEqual(found.groups[0].refIds, ["c1", "c1"]);
    assert.deepEqual(found.groups[0].locators, ["p. 4", "p. 90"]);
  });

  await t.test("the same reference at the same page is one note", () => {
    const found = core.findCitationGroups("See (Smith 2020; Smith 2020).", list);
    assert.deepEqual(found.groups[0].refIds, ["c1"]);
  });

  await t.test("a narrative citation carries its locator too", () => {
    const found = core.findCitationGroups("Smith (2020, p. 44) argues this.", list);
    assert.equal(found.groups.length, 1);
    assert.deepEqual(found.groups[0].refIds, ["c1"]);
    assert.deepEqual(found.groups[0].locators, ["p. 44"]);
  });

  await t.test("the locator reaches the footnote in the built document", async () => {
    const body = [
      p("Locators"),
      p("Growth rose (Smith 2020, p. 44) that year."),
      p("References", { style: "Heading1" }),
      p("Smith, J 2020, A Book, Press, Sydney, https://example.org/book."),
    ];
    const buf = await buildDocx({ body });
    const analysis = await engine.analyzeDocx(buf);
    const built = await engine.buildConvertedDocx(buf, analysis.refs);
    const out = await readDocx(await built.zip.generateAsync({ type: "nodebuffer" }));

    assert.equal(out.paragraphs[1], "Growth rose{1} that year.");
    assert.equal(out.footnotes.length, 1);
    assert.equal(out.footnotes[0].url, "https://example.org/book");
    assert.match(out.footnotes[0].text, /p\. 44/);
  });

  await t.test("a locator survives a reference that has no URL", async () => {
    const body = [
      p("Locators"),
      p("Growth rose (Smith 2020, p. 44) that year."),
      p("References", { style: "Heading1" }),
      p("Smith, J 2020, A Book, Press, Sydney."),
    ];
    const buf = await buildDocx({ body });
    const analysis = await engine.analyzeDocx(buf);
    const built = await engine.buildConvertedDocx(buf, analysis.refs);
    const out = await readDocx(await built.zip.generateAsync({ type: "nodebuffer" }));
    assert.match(out.footnotes[0].text, /A Book/);
    assert.match(out.footnotes[0].text, /p\. 44/);
  });
});

// ===========================================================================
test("1a. splitLocator draws the line between a pinpoint and prose", async (t) => {
  await t.test("recognises the pinpoint forms", () => {
    [
      "p. 44",
      "pp. 44-46",
      "p 44",
      "page 44",
      "pages 44-46",
      "ch. 3",
      "chapter 3",
      "s. 12",
      "sec. 12",
      "para. 7",
      "paras. 7-9",
      "vol. 2",
      "vol. 2, p. 9",
      "no. 4",
      "fig. 2",
      "figs. 1 and 2",
      "art. 5",
      "table 3",
      "n. 5",
      "p. iv",
      "pp. xii-xv",
      "44",
      "44-46",
      "44a",
      "pp. 3, 7 & 11",
    ].forEach(function (s) {
      assert.deepEqual(core.splitLocator(", " + s), { locator: s, residue: "" }, s);
    });
  });

  await t.test("refuses prose, a second citation, and a bare year", () => {
    [
      "emphasis added",
      "my translation",
      "Jones 2021",
      "2021",
      "as cited in Jones 2021",
      "but see Jones 2021",
      "and Jones 2021",
      "cited by Jones",
      "ibid",
    ].forEach(function (s) {
      assert.deepEqual(core.splitLocator(", " + s), { locator: "", residue: s }, s);
    });
  });

  await t.test("a word spelled from roman-numeral letters is prose", () => {
    // i, v, x, l, c, d and m spell real words. Accepting a bare roman numeral
    // would read "(Smith 2020, civil)" as page civil.
    ["civil", "mild", "mix", "id", "dim", "vivid"].forEach(function (s) {
      assert.deepEqual(core.splitLocator(", " + s), { locator: "", residue: s }, s);
    });
  });

  await t.test("nothing after the year is nothing to carry", () => {
    ["", " ", ".", ",", " ).", "; "].forEach(function (s) {
      assert.deepEqual(core.splitLocator(s), { locator: "", residue: "" }, JSON.stringify(s));
    });
  });
});

// ===========================================================================
test("1b. the accounting invariant, over a generated corpus", async (t) => {
  // The examples above are the shapes that were found by hand. This states the
  // rule itself and checks it against every combination, because the defect was
  // not any one of those brackets — it was that nothing anywhere asserted the
  // bracket had to add up.
  const list = refs([
    "Smith, J 2020, A Book, Press, Sydney.",
    "Jones, K 2021, Another Book, Press, Melbourne.",
    "Australian Signals Directorate 2025, Threat Report, ASD, Canberra.",
    "Office of the Australian Information Commissioner 2024a, Breaches, OAIC.",
    "World Health Organization n.d., Fact sheet, WHO, Geneva.",
  ]);

  const authors = [
    "Smith",
    "Smith,",
    "Jones",
    "ASD",
    "Australian Signals Directorate",
    "OAIC",
    "Smith et al.",
    "Nguyen",
    "",
  ];
  const years = ["2020", "2021", "2025", "2024a", "n.d.", "1999"];
  const tails = [
    "",
    ", p. 44",
    ", pp. 44-46",
    ", p 7",
    ", ch. 3",
    ", para. 12",
    ", s. 4",
    ", vol. 2, p. 9",
    ": 44",
    ", 44",
    ", 44-46",
    ", emphasis added",
    ", my translation",
    ", Jones 2021",
    ", 2021",
    ", as cited in Jones 2021",
    ", Table 3",
    ", cited by Jones",
  ];
  const prefixes = ["", "see ", "cf. ", "e.g. ", "see also "];

  /** Letters and digits only, so punctuation and spacing differences vanish. */
  function bare(s) {
    return String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  const corpus = [];
  prefixes.forEach(function (prefix) {
    authors.forEach(function (author) {
      years.forEach(function (year) {
        tails.forEach(function (tail) {
          const inner = (prefix + (author ? author + " " : "") + year + tail).trim();
          if (inner) corpus.push("(" + inner + ")");
        });
      });
    });
  });

  await t.test("the corpus is large enough to be worth calling one", () => {
    assert.ok(corpus.length > 2000, "only " + corpus.length + " cases");
  });

  await t.test("every converted bracket adds up to exactly what it held", () => {
    let convertedCount = 0;
    corpus.forEach(function (bracket) {
      const text = "Evidence shows " + bracket + " and more.";
      const found = core.findCitationGroups(text, list);
      if (!found.groups.length) {
        // Unconverted brackets are left character for character as written.
        assert.equal(survivingText(text, list), text, bracket);
        return;
      }
      convertedCount++;
      assert.equal(found.groups.length, 1, bracket);
      const group = found.groups[0];
      const inner = group.sourceText.slice(1, -1);

      // Rebuild the bracket from the parts the conversion keeps: the author
      // (which becomes the note), the year, the pinpoint, and the signal words
      // a footnote makes redundant. Anything left over is text that would have
      // been deleted.
      let accounted = "";
      core.splitSegments(inner, ";").forEach(function (seg, idx) {
        const unit = core.parseCitationUnit(seg.text);
        if (!unit) return;
        accounted +=
          bare(seg.text.slice(0, seg.text.length).match(/^[\s]*(?:see\s+also|see|e\.?\s?g\.?|i\.?\s?e\.?|cf\.?|compare|also|as\s+cited\s+in|cited\s+in|adapted\s+from|based\s+on|from|after|per|following)\b/i) || "") +
          bare(unit.authorText) +
          bare(unit.year) +
          bare(unit.yearSuffix) +
          bare((group.locators || [])[idx] || unit.locator || "");
      });
      assert.equal(
        bare(inner),
        accounted,
        bracket + " -> accounted for " + JSON.stringify(accounted),
      );
    });
    // Without this the invariant would be satisfied most completely by a
    // matcher that converted nothing at all. 435 of the 2,700 convert; the rest
    // pair an author with a year no reference has.
    assert.ok(convertedCount > 400, "only " + convertedCount + " converted");
  });

  await t.test("a bracket is either converted or reported, never both", () => {
    corpus.forEach(function (bracket) {
      const found = core.findCitationGroups("Evidence shows " + bracket + ".", list);
      assert.ok(
        found.groups.length + found.unmatched.length <= 1,
        bracket + " produced " + found.groups.length + " groups and " +
          found.unmatched.length + " unmatched",
      );
    });
  });
});

