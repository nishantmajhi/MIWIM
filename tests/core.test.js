"use strict";

/**
 * Tests for the pure matching engine (miwim.core.js).
 *
 * These are the rules that decide whether a bracket in someone's document is a
 * citation or just a bracket, and which reference it belongs to. Everything
 * here works on plain strings — no zip, no XML.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("path");

const core = require(path.join(__dirname, "..", "miwim.core.js"));

// Build a reference the way parseReferences would, so tests can talk in terms
// of reference *text* rather than hand-assembled key objects.
let refCounter = 0;
function ref(text, extra) {
  refCounter++;
  const parsed = core.parseReferenceEntry(text);
  return Object.assign(
    {
      n: refCounter,
      id: "c" + refCounter,
      text: text,
      hasUrl: !!parsed.url,
    },
    parsed,
    extra || {},
  );
}

function refs(list) {
  refCounter = 0;
  return list.map(function (t) {
    return ref(t);
  });
}

/** The reference ids matched in a paragraph, flattened in reading order. */
function matchIds(text, references) {
  return core
    .findCitationGroups(text, references)
    .groups.reduce(function (acc, g) {
      return acc.concat(g.refIds);
    }, []);
}

// ===========================================================================
test("heading recognition", async (t) => {
  await t.test("accepts the standard reference-list headings", () => {
    [
      "References",
      "REFERENCES",
      "  references  ",
      "Bibliography",
      "Works Cited",
      "Reference List",
      "Literature Cited",
    ].forEach((h) => {
      assert.equal(core.isReferencesHeadingText(h), true, h);
    });
  });

  await t.test("tolerates numbering and punctuation around the heading", () => {
    // "7. REFERENCES:" and "References" are the same heading to a reader, and
    // numbered headings are the norm in report templates.
    ["7. REFERENCES:", "5 Bibliography", "References."].forEach((h) => {
      assert.equal(core.isReferencesHeadingText(h), true, h);
    });
  });

  await t.test("rejects prose that merely mentions references", () => {
    [
      "References to the Act are set out below",
      "Bibliography and further reading are available online",
      "Introduction",
      "",
    ].forEach((h) => {
      assert.equal(core.isReferencesHeadingText(h), false, h);
    });
  });

  await t.test("recognises the headings that end a reference list", () => {
    [
      "Appendix",
      "Appendix A",
      "Appendix 1",
      "Annex B",
      "Acknowledgements",
      "Declaration of AI Use",
      "About the Authors",
      "Index",
      "Word count",
    ].forEach((h) => {
      assert.equal(core.isSectionEndHeadingText(h), true, h);
    });
    assert.equal(core.isSectionEndHeadingText("Smith, J 2020, A Book"), false);
  });
});

// ===========================================================================
test("year detection", async (t) => {
  await t.test("finds a plain four-digit year", () => {
    const found = core.findYearTokens("Published 2024 by Press");
    assert.equal(found.length, 1);
    assert.equal(found[0].year, "2024");
    assert.equal(found[0].suffix, "");
  });

  await t.test("captures the disambiguating suffix", () => {
    const found = core.findYearTokens("(OAIC 2025b)");
    assert.equal(found[0].year, "2025");
    assert.equal(found[0].suffix, "b");
  });

  await t.test("ignores digits glued to other digits or letters", () => {
    // C2024A00098 is a legislation identifier; 20255 is not a year at all.
    assert.deepEqual(core.findYearTokens("C2024A00098"), []);
    assert.deepEqual(core.findYearTokens("20255"), []);
    assert.deepEqual(core.findYearTokens("id-2024x"), []);
  });

  await t.test("ignores both halves of a numeric range", () => {
    // "the 2024-25 financial year" and "the 2023-2030 Strategy" are the two
    // shapes that appear constantly in Australian government writing; reading
    // either half as a publication year invents citations that aren't there.
    assert.deepEqual(core.findYearTokens("the 2024-25 financial year"), []);
    assert.deepEqual(core.findYearTokens("2023-2030 Australian Strategy"), []);
  });

  await t.test("ignores four-digit numbers outside the plausible year range", () => {
    // "[2025] FCA 1224" — the year is real, the case number is not.
    const found = core.findYearTokens("[2025] FCA 1224 was decided");
    assert.equal(found.length, 1);
    assert.equal(found[0].year, "2025");
  });

  await t.test("recognises n.d. in its several spellings", () => {
    ["n.d.", "n. d.", "nd", "no date"].forEach((s) => {
      const hit = core.findNoDate("Tripadvisor (" + s + "). The Wee Kitchen.");
      assert.ok(hit, s);
      assert.equal(hit.year, "n.d.");
    });
  });

  await t.test("does not read 'nd' inside a word as a missing date", () => {
    assert.equal(core.findNoDate("and then, secondhand"), null);
  });
});

// ===========================================================================
test("locating the year within a reference entry", async (t) => {
  await t.test("prefers the parenthesised year of APA style", () => {
    const located = core.locateReferenceYear(
      "Smith, J. (2020). Networks in 1998 revisited. Press.",
    );
    assert.equal(located.year, "2020");
  });

  await t.test("takes the bare year of Harvard/AGPS style", () => {
    const located = core.locateReferenceYear(
      "ASD (Australian Signals Directorate) 2025, Annual Cyber Threat Report 2024-25, Canberra.",
    );
    assert.equal(located.year, "2025");
  });

  await t.test("keeps the suffix that distinguishes same-year entries", () => {
    assert.equal(core.locateReferenceYear("OAIC 2025b, Latest statistics").suffix, "b");
  });

  await t.test("is not fooled by four-digit slugs inside a URL", () => {
    // The old naive scan read ".../report-2024-2025" as the year, which broke
    // matching for every reference carrying a dated URL.
    const located = core.locateReferenceYear(
      "ASD 2025, Report, viewed 11 August 2026, <https://cyber.gov.au/annual-cyber-threat-report-2024-2025>.",
    );
    assert.equal(located.year, "2025");
  });

  await t.test("returns null when there is no date at all", () => {
    assert.equal(core.locateReferenceYear("An untitled note, publisher unknown"), null);
  });
});

// ===========================================================================
test("author keys", async (t) => {
  await t.test("splits an acronym from its bracketed expansion", () => {
    // Cited as "(ASD 2025)" but written out in full in the reference list —
    // both spellings have to be matchable.
    const keys = core.buildAuthorKeys("ASD (Australian Signals Directorate) ");
    assert.deepEqual(keys, [["asd"], ["australian", "signals", "directorate"]]);
  });

  await t.test("drops connecting words so wording can vary", () => {
    // The primary key is the name's own words, with "of" dropped so that
    // "Dept of Home Affairs" and "Department of Home Affairs" compare equal.
    // Asserted on the primary key specifically rather than the whole list,
    // because derived keys (see below) are appended after it.
    assert.deepEqual(core.buildAuthorKeys("Department of Home Affairs ")[0], [
      "department",
      "home",
      "affairs",
    ]);
  });

  await t.test("derives the acronym an organisation is cited by", () => {
    // AGPS style leaves the acronym out of the author segment entirely — it
    // appears in the publisher field, after the year:
    //   Office of the Australian Information Commissioner 2025a, …, OAIC, Sydney
    // so "(OAIC 2025a)" has no literal text to match against. The acronym is
    // derived from the name instead. The words dropped in forming it are the
    // same stopwords normalizeTokens already removes.
    const keys = core.buildAuthorKeys(
      "Office of the Australian Information Commissioner ",
    );
    assert.deepEqual(keys[0], [
      "office",
      "australian",
      "information",
      "commissioner",
    ]);
    assert.ok(
      keys.some((k) => k.length === 1 && k[0] === "oaic"),
      "expected a derived 'oaic' key",
    );
  });

  await t.test("doesn't repeat an acronym the entry already spells out", () => {
    // "ASD (Australian Signals Directorate)" already carries both spellings.
    // Deriving "asd" from the expansion must not add a duplicate key.
    const keys = core.buildAuthorKeys("ASD (Australian Signals Directorate) ");
    assert.deepEqual(keys, [["asd"], ["australian", "signals", "directorate"]]);
  });

  await t.test("gives a single-word name no acronym", () => {
    // A one-token name has no acronym; its own first letter is not one.
    // Without this, "Kennedys" would answer to "(K 2025)".
    assert.equal(core.acronymOf(["kennedys"]), null);
    assert.equal(core.acronymOf(["bird", "bird"]), "bb");
  });

  await t.test("drops initials, which citations never repeat", () => {
    assert.deepEqual(core.normalizeTokens("Smith, J. A."), ["smith"]);
  });

  await t.test("treats & and 'and' alike", () => {
    assert.deepEqual(
      core.normalizeTokens("Smith & Jones"),
      core.normalizeTokens("Smith and Jones"),
    );
  });

  await t.test("collects surnames as a fallback key", () => {
    const surnames = core.buildSurnames("Smith, A., Jones, B. & Wu, C. ");
    assert.deepEqual(surnames, ["smith", "jones", "wu"]);
  });
});

// ===========================================================================
test("URL extraction", async (t) => {
  await t.test("pulls an https URL out of running text", () => {
    assert.equal(
      core.extractUrl("Viewed 11 August 2026, https://example.gov.au/report"),
      "https://example.gov.au/report",
    );
  });

  await t.test("strips the angle brackets and full stop of AGPS style", () => {
    // "viewed 11 August 2026, <https://www.cyber.gov.au/reports>." — the
    // trailing "." and ">" are punctuation, not part of the address.
    assert.equal(
      core.extractUrl("viewed 11 August 2026, <https://www.cyber.gov.au/reports>."),
      "https://www.cyber.gov.au/reports",
    );
  });

  await t.test("keeps a trailing slash, which is part of the path", () => {
    assert.equal(
      core.extractUrl("<https://asic.gov.au/media/26-021mr-fiig/>."),
      "https://asic.gov.au/media/26-021mr-fiig/",
    );
  });

  await t.test("keeps balanced parentheses inside a URL", () => {
    assert.equal(
      core.extractUrl("See https://en.wikipedia.org/wiki/Mercury_(planet) for more"),
      "https://en.wikipedia.org/wiki/Mercury_(planet)",
    );
  });

  await t.test("drops an unbalanced closing parenthesis that wraps the URL", () => {
    assert.equal(
      core.extractUrl("(available at https://example.com/doc)"),
      "https://example.com/doc",
    );
  });

  await t.test("gives a bare www address a scheme", () => {
    // Without one the link is relative and Word opens it as a local file.
    assert.equal(core.extractUrl("See www.oaic.gov.au/privacy"), "https://www.oaic.gov.au/privacy");
  });

  await t.test("turns a DOI into a resolvable link", () => {
    assert.equal(
      core.extractUrl("Journal of Things, 4(2), 11-20. doi:10.1000/abc123."),
      "https://doi.org/10.1000/abc123",
    );
  });

  await t.test("returns null when there is no URL", () => {
    assert.equal(core.extractUrl("Smith, J 2020, A Book, Press, Sydney."), null);
  });
});

// ===========================================================================
test("citation units", async (t) => {
  await t.test("parses author and year out of a parenthetical", () => {
    const unit = core.parseCitationUnit("Smith 2020");
    assert.deepEqual(unit.tokens, ["smith"]);
    assert.equal(unit.year, "2020");
  });

  await t.test("strips a signal phrase before the author", () => {
    ["see Smith 2020", "e.g. Smith 2020", "cf. Smith 2020", "as cited in Smith 2020"].forEach(
      (s) => {
        assert.deepEqual(core.parseCitationUnit(s).tokens, ["smith"], s);
      },
    );
  });

  await t.test("flags et al. citations", () => {
    assert.equal(core.parseCitationUnit("Castro et al., 2024").etAl, true);
  });

  await t.test("rejects brackets that hold no date", () => {
    // These are the false positives that made the tool unusable on legal
    // writing: statutory tags, subsection letters and form placeholders.
    ["Cth", "No 2", "a", "d", "h", "[Student Number 1]", "Australian Signals Directorate"].forEach(
      (s) => {
        assert.equal(core.parseCitationUnit(s), null, s);
      },
    );
  });

  await t.test("rejects an empty segment", () => {
    assert.equal(core.parseCitationUnit("   "), null);
  });
});

// ===========================================================================
test("matching citations to references", async (t) => {
  await t.test("matches on author and year", () => {
    const list = refs(["Smith, J 2020, A Book, Press."]);
    assert.deepEqual(matchIds("Growth was strong (Smith 2020).", list), ["c1"]);
  });

  await t.test("matches an acronym citation to a spelled-out reference", () => {
    const list = refs([
      "ASD (Australian Signals Directorate) 2025, Annual Cyber Threat Report 2024-25, Canberra.",
    ]);
    assert.deepEqual(matchIds("Reports rose sharply (ASD 2025).", list), ["c1"]);
  });

  await t.test("matches a spelled-out citation to an acronym-first reference", () => {
    const list = refs([
      "OAIC (Office of the Australian Information Commissioner) 2025b, Statistics.",
    ]);
    assert.deepEqual(
      matchIds("As recorded (Office of the Australian Information Commissioner 2025b).", list),
      ["c1"],
    );
  });

  await t.test("matches an acronym citation to a plain spelled-out reference", () => {
    // The bracketed style above is not the only one. AGPS also writes the
    // name out plainly and puts the acronym in the publisher field, after
    // the year — where the author segment never sees it. This is the shape
    // that appeared in ELL_917: six "(OAIC …)" citations reported unmatched
    // while the six entries they pointed at were reported never cited.
    const list = refs([
      "Office of the Australian Information Commissioner 2026a, Data breach notifications increase, OAIC, Sydney.",
    ]);
    assert.deepEqual(matchIds("Breaches hit a high (OAIC 2026a).", list), ["c1"]);
  });

  await t.test("keeps plain-style entries apart by suffix through the acronym", () => {
    // The acronym must not become a wildcard that collapses a/b/c together.
    const list = refs([
      "Office of the Australian Information Commissioner 2025a, Statistics, OAIC, Sydney.",
      "Office of the Australian Information Commissioner 2025b, Statutory tort, OAIC, Sydney.",
      "Office of the Australian Information Commissioner 2025c, Optus action, OAIC, Sydney.",
    ]);
    assert.deepEqual(matchIds("See (OAIC 2025a).", list), ["c1"]);
    assert.deepEqual(matchIds("See (OAIC 2025b).", list), ["c2"]);
    assert.deepEqual(matchIds("See (OAIC 2025c).", list), ["c3"]);
  });

  await t.test("prefers a literal name match over another entry's acronym", () => {
    // "ASIC" is both a real entry and the initials of the other one. The
    // entry that actually carries the word must win.
    const list = refs([
      "Australian Securities and Investments Commission 2026, media release.",
      "ASIC 2026, a different note.",
    ]);
    assert.deepEqual(matchIds("The regulator acted (ASIC 2026).", list), ["c2"]);
  });

  await t.test("does not invent an acronym match for an unrelated entry", () => {
    // "Bird and Bird" initialises to "bb". A "(BBC 2024)" citation shares no
    // author with it and must not be converted on the strength of the year.
    const list = refs(["Bird and Bird 2024, Australia's first cyber security law."]);
    assert.deepEqual(matchIds("As reported (BBC 2024).", list), []);
  });

  await t.test("keeps same-year references apart by their letter suffix", () => {
    // Three OAIC 2025 entries: the whole point of a/b/c is that "(OAIC 2025b)"
    // must land on exactly one of them.
    const list = refs([
      "OAIC 2025a, Clinical Labs penalties, media release.",
      "OAIC 2025b, Latest Notifiable Data Breach statistics.",
      "OAIC 2025c, Statutory tort for serious invasions of privacy.",
    ]);
    assert.deepEqual(matchIds("See (OAIC 2025a).", list), ["c1"]);
    assert.deepEqual(matchIds("See (OAIC 2025b).", list), ["c2"]);
    assert.deepEqual(matchIds("See (OAIC 2025c).", list), ["c3"]);
  });

  await t.test("distinguishes two organisations sharing a year", () => {
    const list = refs([
      "ASIC (Australian Securities and Investments Commission) 2026, media release.",
      "Federal Court of Australia 2026, ASIC v FIIG Securities [2026] FCA 92.",
    ]);
    assert.deepEqual(matchIds("The regulator acted (ASIC 2026).", list), ["c1"]);
    assert.deepEqual(matchIds("The court held (Federal Court of Australia 2026).", list), ["c2"]);
  });

  await t.test("matches et al. against the first author only", () => {
    const list = refs(["Castro, A., Diaz, B., Evans, C. & Fox, D. (2024). A paper. Journal."]);
    assert.deepEqual(matchIds("Others agree (Castro et al., 2024).", list), ["c1"]);
  });

  await t.test("matches a citation naming a later author", () => {
    const list = refs(["Smith, A., Jones, B. & Wu, C. (2021). Findings. Journal."]);
    assert.deepEqual(matchIds("As found (Jones, 2021).", list), ["c1"]);
  });

  await t.test("matches an n.d. citation", () => {
    const list = refs(["Tripadvisor (n.d.). The Wee Kitchen. https://example.com/wee"]);
    assert.deepEqual(matchIds("Reviews are mixed (Tripadvisor, n.d.).", list), ["c1"]);
  });

  await t.test("refuses to match on year alone", () => {
    // Otherwise every "(2025)" in the document collapses onto whichever 2025
    // reference happens to come first.
    const list = refs(["Smith, J 2025, A Book, Press."]);
    assert.deepEqual(matchIds("Costs rose that year (2025).", list), []);
  });

  await t.test("refuses to match the wrong year", () => {
    const list = refs(["Smith, J 2020, A Book, Press."]);
    assert.deepEqual(matchIds("Growth was strong (Smith 2021).", list), []);
  });
});

// ===========================================================================
test("brackets that are not citations", async (t) => {
  const list = refs([
    "ASD (Australian Signals Directorate) 2025, Annual Cyber Threat Report 2024-25.",
    "Federal Court of Australia 2026, ASIC v FIIG Securities Limited [2026] FCA 92.",
  ]);

  await t.test("leaves statutory tags alone", () => {
    // The single biggest source of false positives in legal writing: a report
    // citing the Privacy Act 1988 (Cth) three times per sentence.
    const text =
      "including the Privacy Act 1988 (Cth), the Privacy and Other Legislation Amendment Act 2024 (Cth) and the Cyber Security Act 2024 (Cth).";
    assert.deepEqual(matchIds(text, list), []);
  });

  await t.test("leaves subsection letters alone", () => {
    const text = "sections 912A(1)(a), (d) and (h) of the Corporations Act 2001 (Cth)";
    assert.deepEqual(matchIds(text, list), []);
  });

  await t.test("leaves case-name qualifiers alone", () => {
    const text =
      "In Australian Clinical Labs Limited (No 2) [2025] FCA 1224 the Federal Court ordered";
    assert.deepEqual(matchIds(text, list), []);
  });

  await t.test("leaves template placeholders alone", () => {
    assert.deepEqual(matchIds("Prepared by: [Student Name 1] ([Student Number 1])", list), []);
  });

  await t.test("leaves a bracket holding a date range alone", () => {
    assert.deepEqual(matchIds("over the period (2024-25) costs rose", list), []);
  });

  await t.test("reports an unrecognised year-bearing bracket as unmatched", () => {
    // Shaped like a citation, matches nothing — the person needs to see it
    // rather than have it silently ignored.
    const found = core.findCitationGroups("As shown (Nguyen 2019).", list);
    assert.equal(found.groups.length, 0);
    assert.equal(found.unmatched.length, 1);
    assert.equal(found.unmatched[0].text, "(Nguyen 2019)");
  });
});

// ===========================================================================
test("multiple citations in one bracket", async (t) => {
  const list = refs([
    "Smith, J 2020, A Book, Press.",
    "Jones, K 2021, Another Book, Press.",
    "Wu, L 2022, A Third Book, Press.",
  ]);

  await t.test("produces one marker per citation, in reading order", () => {
    const found = core.findCitationGroups("Several agree (Smith 2020; Jones 2021).", list);
    assert.equal(found.groups.length, 1);
    assert.deepEqual(found.groups[0].refIds, ["c1", "c2"]);
  });

  await t.test("preserves the order the citations were written in", () => {
    const found = core.findCitationGroups("Several agree (Jones 2021; Smith 2020).", list);
    assert.deepEqual(found.groups[0].refIds, ["c2", "c1"]);
  });

  await t.test("handles three at once", () => {
    const found = core.findCitationGroups(
      "All three (Smith 2020; Jones 2021; Wu 2022) concur.",
      list,
    );
    assert.deepEqual(found.groups[0].refIds, ["c1", "c2", "c3"]);
  });

  await t.test("leaves the bracket alone when only part of it is a citation", () => {
    // Converting half of "(2024-25 figures; Smith 2020)" would delete the words
    // "2024-25 figures" from the document. Better to convert nothing and say so.
    const found = core.findCitationGroups("See (2024-25 figures; Smith 2020).", list);
    assert.equal(found.groups.length, 0);
    assert.equal(found.unmatched.length, 1);
  });

  await t.test("leaves the bracket alone when one citation matches nothing", () => {
    const found = core.findCitationGroups("See (Smith 2020; Nguyen 2019).", list);
    assert.equal(found.groups.length, 0);
    assert.equal(found.unmatched.length, 1);
  });
});

// ===========================================================================
test("narrative citations", async (t) => {
  await t.test("takes the author from the words before the bracket", () => {
    const list = refs(["Department of Home Affairs 2026a, Cyber Security Act."]);
    const found = core.findCitationGroups(
      "The Department of Home Affairs (2026a) explains the measures.",
      list,
    );
    assert.equal(found.groups.length, 1);
    assert.deepEqual(found.groups[0].refIds, ["c1"]);
    // Only the bracket is replaced; the author's name stays in the sentence,
    // where it is doing grammatical work.
    assert.equal(found.groups[0].sourceText, "(2026a)");
  });

  await t.test("handles a single-surname narrative citation", () => {
    const list = refs(["Smith, J 2020, A Book, Press."]);
    assert.deepEqual(matchIds("Smith (2020) argues otherwise.", list), ["c1"]);
  });

  await t.test("does not invent an author from a lowercase preceding word", () => {
    const list = refs(["Smith, J 2020, A Book, Press."]);
    assert.deepEqual(matchIds("this happened again (2020).", list), []);
  });
});

// ===========================================================================
test("splitting a paragraph for the preview", async (t) => {
  const list = refs(["Smith, J 2020, A Book, Press.", "Jones, K 2021, Another Book, Press."]);

  await t.test("alternates text and citation parts", () => {
    const split = core.splitParagraphForPreview("Growth was strong (Smith 2020). It held.", list);
    assert.deepEqual(
      split.parts.map((p) => p.type),
      ["text", "cite", "text"],
    );
    assert.equal(split.parts[0].value, "Growth was strong");
    assert.equal(split.parts[1].refId, "c1");
    assert.equal(split.parts[2].value, ". It held.");
  });

  await t.test("swallows the space before the citation", () => {
    // The marker replaces " (Smith 2020)", so the sentence reads "strong¹."
    // rather than "strong ¹." with an orphaned space before the full stop.
    const split = core.splitParagraphForPreview("Growth was strong (Smith 2020).", list);
    assert.equal(split.parts[0].value, "Growth was strong");
  });

  await t.test("shows a shared bracket only once", () => {
    const split = core.splitParagraphForPreview("Both (Smith 2020; Jones 2021) agree.", list);
    const cites = split.parts.filter((p) => p.type === "cite");
    assert.equal(cites.length, 2);
    assert.equal(cites[0].text, "(Smith 2020; Jones 2021)");
    assert.equal(cites[1].text, "");
  });

  await t.test("returns the paragraph unchanged when nothing matches", () => {
    const split = core.splitParagraphForPreview("Nothing to see here.", list);
    assert.deepEqual(split.parts, [{ type: "text", value: "Nothing to see here." }]);
  });

  await t.test("handles an empty reference list without throwing", () => {
    const split = core.splitParagraphForPreview("Growth was strong (Smith 2020).", []);
    assert.equal(split.parts.length, 1);
    assert.equal(split.groups.length, 0);
  });
});
