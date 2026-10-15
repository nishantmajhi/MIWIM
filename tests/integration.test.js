"use strict";

/**
 * End-to-end test against the document that motivated this work.
 *
 * The BUS702 cybersecurity report is an awkward document in almost every way a
 * real student document is awkward, and each of those difficulties is
 * reproduced here in a single fixture:
 *
 *   - Australian Harvard/AGPS references, where the year is bare rather than
 *     bracketed:   ASD (Australian Signals Directorate) 2025, Annual Report...
 *   - Agencies cited by acronym but written out in full in the list.
 *   - Same-author, same-year entries separated only by a letter: 2025a/b/c.
 *   - Legal prose dense with bracketed non-citations: (Cth), (No 2), (a), (d),
 *     (h), and template placeholders like ([Student Number 1]).
 *   - Numbers that look like years but aren't: "2024-25", "2023-2030",
 *     "[2025] FCA 1224", "C2024A00098".
 *   - A generated table of contents that duplicates the "Bibliography" heading.
 *   - URLs wrapped in angle brackets with a full stop after the closing one.
 *
 * The expected output is the one produced by hand in ELL_197: every in-text
 * citation replaced by a linked footnote carrying that source's URL, and the
 * bibliography collapsed to a single placeholder line.
 */

const test = require("node:test");
const assert = require("node:assert/strict");

const h = require("./helpers/docx-fixture.js");
const { p, sdt, buildDocx, convert, engine } = h;

const URLS = {
  asd: "https://www.cyber.gov.au/about-us/view-all-content/reports-and-statistics/annual-cyber-threat-report-2024-2025",
  asic: "https://www.asic.gov.au/about-asic/news-centre/find-a-media-release/2026-releases/26-021mr-asic-action-sees-fiig-securities-ordered-to-pay-2-5-million-over-cyber-security-failures/",
  homeAffairsAct:
    "https://www.homeaffairs.gov.au/cyber-security-subsite/Pages/cyber-security-act.aspx",
  homeAffairsStrategy:
    "https://www.homeaffairs.gov.au/about-us/our-portfolios/cyber-security/strategy/2023-2030-australian-cyber-security-strategy",
  fedCourt:
    "https://download.asic.gov.au/media/o02h30dd/26-021mr-asic-v-fiig-securities-limited-judgment-13-feb-2026.pdf",
  legislation: "https://www.legislation.gov.au/C2024A00098/asmade/text",
  oaicA:
    "https://www.oaic.gov.au/news/media-centre/australian-clinical-labs-ordered-to-pay-penalties-in-relation-to-medlab-pathology-data-breach-in-first-for-privacy-act",
  oaicB:
    "https://www.oaic.gov.au/news/blog/latest-notifiable-data-breach-statistics-for-january-to-june-2025",
  oaicC:
    "https://www.oaic.gov.au/privacy/your-privacy-rights/more-privacy-rights/statutory-tort-for-serious-invasions-of-privacy",
};

// A bibliography entry in AGPS style: text, then the URL as a real hyperlink
// wrapped in angle brackets, then the closing full stop.
function bibEntry(prefix, rid, url) {
  return p([{ text: prefix + " <" }, { text: url, rid: rid }, { text: ">." }]);
}

function buildReport() {
  return buildDocx({
    rels: [
      { id: "rIdA", target: URLS.asd },
      { id: "rIdB", target: URLS.asic },
      { id: "rIdC", target: URLS.homeAffairsAct },
      { id: "rIdD", target: URLS.homeAffairsStrategy },
      { id: "rIdE", target: URLS.fedCourt },
      { id: "rIdF", target: URLS.legislation },
      { id: "rIdG", target: URLS.oaicA },
      { id: "rIdH", target: URLS.oaicB },
      { id: "rIdI", target: URLS.oaicC },
    ],
    body: [
      p("Cybersecurity Challenges for Australian Businesses", { style: "Title" }),
      p("BUS702 Business Law - Research and Presentation Project", { style: "Subtitle" }),

      p("Index", { style: "Heading1" }),
      sdt([
        p("Introduction\t1", { style: "TOC1" }),
        p("Bibliography\t9", { style: "TOC1" }),
      ]),

      p("Introduction", { style: "Heading1" }),
      p(
        "The Australian Signals Directorate recorded more than 84,700 cybercrime reports in the 2024-25 financial year, and found that the average self-reported cost of cybercrime to a small business rose by 14 per cent to approximately $56,600 (ASD 2025). Losses of that size can threaten the survival of an ordinary Australian business.",
      ),
      p(
        "The report is organised into two chapters, including the Privacy Act 1988 (Cth), the Privacy and Other Legislation Amendment Act 2024 (Cth) and the Cyber Security Act 2024 (Cth).",
      ),

      p("Chapter 1: The Threat Environment", { style: "Heading1" }),
      p("Prepared and presented by: [Student Name 1] ([Student Number 1])"),

      p("1.1 The scale of the threat", { style: "Heading2" }),
      // Run split mid-citation, as Word produces after an edit.
      p([
        "In 2024-25 the Australian Signals Directorate responded to more than 1,200 cyber security incidents, an increase of 11 per cent, and notified organisations more than 1,700 times (ASD ",
        "2025). Phishing featured in 60 per cent of reported incidents.",
      ]),
      p(
        "Averaged across all business sizes, the self-reported cost of cybercrime per report rose 50 per cent to $80,850, with medium businesses reporting $97,200 and large businesses $202,700 (ASD 2025).",
      ),

      p("1.2 Existing duties under the Privacy Act 1988", { style: "Heading2" }),
      p(
        "The Office of the Australian Information Commissioner received 532 breach notifications between January and June 2025, with cyber incidents affecting an average of just over 10,000 individuals each (OAIC 2025b).",
      ),

      p("1.3 The Privacy and Other Legislation Amendment Act 2024", { style: "Heading2" }),
      p(
        "An individual may now sue directly where the public interest in privacy outweighs any competing public interest (OAIC 2025c). Remedies include damages, injunctions and an order requiring an apology.",
      ),

      p("1.4 The Cyber Security Act 2024", { style: "Heading2" }),
      p(
        "Australia's first standalone cybersecurity statute received Royal Assent on 29 November 2024 (Federal Register of Legislation 2024). It delivers four measures, including a Cyber Incident Review Board (Department of Home Affairs 2026a).",
      ),
      p(
        "These measures implement the 2023-2030 Australian Cyber Security Strategy, which aims to make Australia a world leader in cyber security by 2030 (Department of Home Affairs 2026b).",
      ),

      p("Chapter 2: Enforcement in the Courts", { style: "Heading1" }),
      p("Prepared and presented by: [Student Name 2] ([Student Number 2])"),

      p("2.1 The first civil penalty under the Privacy Act", { style: "Heading2" }),
      p(
        "In Australian Information Commissioner v Australian Clinical Labs Limited (No 2) [2025] FCA 1224 the Federal Court ordered the pathology provider to pay $5.8 million. The penalty was divided into $4.2 million for failing to take reasonable steps, and $800,000 for failing to notify the Commissioner (OAIC 2025a).",
      ),

      p("2.2 Cybersecurity as a condition of holding a licence", { style: "Heading2" }),
      p(
        "Australian Securities and Investments Commission v FIIG Securities Limited [2026] FCA 92 extended enforcement into corporations law, and a 2023 attack removed roughly 385 gigabytes of data belonging to about 18,000 clients (ASIC 2026).",
      ),
      p(
        "The Court declared contraventions of sections 912A(1)(a), (d) and (h) of the Corporations Act 2001 (Cth). Derrington J made an important qualification: a successful attack does not by itself prove a breach (Federal Court of Australia 2026).",
      ),

      p("Bibliography", { style: "Heading1" }),
      bibEntry(
        "ASD (Australian Signals Directorate) 2025, Annual Cyber Threat Report 2024-25, Canberra, viewed 11 August 2026,",
        "rIdA",
        URLS.asd,
      ),
      bibEntry(
        "ASIC (Australian Securities and Investments Commission) 2026, ASIC action sees FIIG Securities ordered to pay $2.5 million, media release 26-021MR, 9 February, viewed 11 August 2026,",
        "rIdB",
        URLS.asic,
      ),
      bibEntry(
        "Department of Home Affairs 2026a, Cyber Security Act, Australian Government, Canberra, viewed 11 August 2026,",
        "rIdC",
        URLS.homeAffairsAct,
      ),
      bibEntry(
        "Department of Home Affairs 2026b, 2023-2030 Australian Cyber Security Strategy, Australian Government, Canberra, viewed 11 August 2026,",
        "rIdD",
        URLS.homeAffairsStrategy,
      ),
      bibEntry(
        "Federal Court of Australia 2026, Australian Securities and Investments Commission v FIIG Securities Limited [2026] FCA 92, judgment of Derrington J, 13 February, viewed 11 August 2026,",
        "rIdE",
        URLS.fedCourt,
      ),
      bibEntry(
        "Federal Register of Legislation 2024, Cyber Security Act 2024 (Cth), Australian Government, Canberra, viewed 11 August 2026,",
        "rIdF",
        URLS.legislation,
      ),
      bibEntry(
        "OAIC (Office of the Australian Information Commissioner) 2025a, Australian Clinical Labs ordered to pay penalties, media release, 9 October, viewed 11 August 2026,",
        "rIdG",
        URLS.oaicA,
      ),
      bibEntry(
        "OAIC (Office of the Australian Information Commissioner) 2025b, Latest Notifiable Data Breach statistics for January to June 2025, blog post, 4 November, viewed 11 August 2026,",
        "rIdH",
        URLS.oaicB,
      ),
      bibEntry(
        "OAIC (Office of the Australian Information Commissioner) 2025c, Statutory tort for serious invasions of privacy, viewed 11 August 2026,",
        "rIdI",
        URLS.oaicC,
      ),
    ],
  });
}

// The citations in the order they appear in the document, with the URL each
// footnote must end up carrying. This is the expected result, spelled out.
const EXPECTED_FOOTNOTE_URLS = [
  URLS.asd, // (ASD 2025)  intro
  URLS.asd, // (ASD 2025)  1.1 first
  URLS.asd, // (ASD 2025)  1.1 second
  URLS.oaicB, // (OAIC 2025b)
  URLS.oaicC, // (OAIC 2025c)
  URLS.legislation, // (Federal Register of Legislation 2024)
  URLS.homeAffairsAct, // (Department of Home Affairs 2026a)
  URLS.homeAffairsStrategy, // (Department of Home Affairs 2026b)
  URLS.oaicA, // (OAIC 2025a)
  URLS.asic, // (ASIC 2026)
  URLS.fedCourt, // (Federal Court of Australia 2026)
];

// ===========================================================================
test("the BUS702 cybersecurity report", async (t) => {
  const buffer = await buildReport();
  const analysis = await engine.analyzeDocx(buffer);
  const converted = await convert(buffer);

  await t.test("reads every bibliography entry", () => {
    assert.equal(analysis.hasRefsSection, true);
    assert.equal(analysis.refs.length, 9);
  });

  await t.test("recovers each entry's URL from its hyperlink", () => {
    assert.deepEqual(
      analysis.refs.map((r) => r.url),
      [
        URLS.asd,
        URLS.asic,
        URLS.homeAffairsAct,
        URLS.homeAffairsStrategy,
        URLS.fedCourt,
        URLS.legislation,
        URLS.oaicA,
        URLS.oaicB,
        URLS.oaicC,
      ],
    );
  });

  await t.test("reads the bare Harvard year, not a digit from the URL", () => {
    // "annual-cyber-threat-report-2024-2025" sits in the same paragraph and is
    // the trap the old bare-\d{4} scan fell into.
    assert.equal(analysis.refs[0].year, "2025");
    assert.equal(analysis.refs[0].yearSuffix, "");
  });

  await t.test("keeps the a/b/c suffixes that separate same-year entries", () => {
    assert.equal(analysis.refs[2].yearSuffix, "a"); // Home Affairs 2026a
    assert.equal(analysis.refs[3].yearSuffix, "b"); // Home Affairs 2026b
    assert.equal(analysis.refs[6].yearSuffix, "a"); // OAIC 2025a
    assert.equal(analysis.refs[7].yearSuffix, "b"); // OAIC 2025b
    assert.equal(analysis.refs[8].yearSuffix, "c"); // OAIC 2025c
  });

  await t.test("matches every in-text citation", () => {
    assert.equal(analysis.citationCount, EXPECTED_FOOTNOTE_URLS.length);
    assert.deepEqual(analysis.unmatched, []);
    assert.deepEqual(analysis.uncitedRefs, []);
  });

  await t.test("converts nothing that isn't a citation", () => {
    // (Cth) x4, (No 2), (1), (a), (d), (h), ([Student Number 1]) x2 —
    // eleven brackets that must survive the conversion untouched.
    const text = converted.out.paragraphs.join("\n");
    assert.match(text, /Privacy Act 1988 \(Cth\)/);
    assert.match(text, /Cyber Security Act 2024 \(Cth\)/);
    assert.match(text, /Clinical Labs Limited \(No 2\) \[2025\] FCA 1224/);
    assert.match(text, /sections 912A\(1\)\(a\), \(d\) and \(h\)/);
    assert.match(text, /\[Student Name 1\] \(\[Student Number 1\]\)/);
    assert.match(text, /\[Student Name 2\] \(\[Student Number 2\]\)/);
  });

  await t.test("leaves the year-shaped numbers that aren't years", () => {
    const text = converted.out.paragraphs.join("\n");
    assert.match(text, /the 2024-25 financial year/);
    assert.match(text, /2023-2030 Australian Cyber Security Strategy/);
    assert.match(text, /world leader in cyber security by 2030/);
    assert.match(text, /\[2026\] FCA 92/);
  });

  await t.test("creates one footnote per citation, in reading order", () => {
    assert.equal(converted.summary.footnotesCreated, EXPECTED_FOOTNOTE_URLS.length);
    assert.deepEqual(
      converted.out.footnotes.map((f) => f.url),
      EXPECTED_FOOTNOTE_URLS,
    );
  });

  await t.test("puts each marker where its citation was", () => {
    const text = converted.out.paragraphs.join("\n");
    assert.match(text, /approximately \$56,600\{1\}\. Losses/);
    assert.match(text, /more than 1,700 times\{2\}\. Phishing/);
    assert.match(text, /large businesses \$202,700\{3\}\./);
    assert.match(text, /10,000 individuals each\{4\}\./);
    assert.match(text, /Royal Assent on 29 November 2024\{6\}\./);
    assert.match(text, /Cyber Incident Review Board\{7\}\./);
  });

  await t.test("removes the citation and the space in front of it", () => {
    // Not "…$56,600 ¹." with a stranded space before the full stop.
    const text = converted.out.paragraphs.join("\n");
    assert.doesNotMatch(text, /\s\{\d+\}/);
    assert.doesNotMatch(text, /\(ASD 2025\)/);
    assert.doesNotMatch(text, /\(OAIC 2025[abc]\)/);
    assert.doesNotMatch(text, /\(Department of Home Affairs 2026[ab]\)/);
  });

  await t.test("collapses the bibliography to the placeholder", () => {
    assert.equal(converted.summary.referenceSectionCollapsed, true);
    assert.equal(converted.summary.referencesKept, 0);
    const text = converted.out.paragraphs.join("\n");
    assert.match(text, /Bibliography/);
    assert.ok(converted.out.paragraphs.includes(engine.PLACEHOLDER_TEXT));
    assert.doesNotMatch(text, /Annual Cyber Threat Report 2024-25/);
  });

  await t.test("does not mistake the table of contents for the real section", () => {
    // The TOC also contains the word "Bibliography"; if it were picked up, the
    // reference list would be a single page-number line.
    assert.equal(analysis.refs.length, 9);
    assert.match(converted.out.paragraphs.join("\n"), /Introduction\t1/);
  });

  await t.test("writes each footnote as a live link", () => {
    converted.out.footnotes.forEach((fn) => {
      assert.equal(fn.hasFootnoteRef, true);
      assert.equal(fn.styleId, "FootnoteText");
      assert.equal(fn.text, fn.url);
    });
  });

  await t.test("shares one relationship per distinct URL", () => {
    // ASD is cited three times but is one source; nine sources, eight of which
    // appear once, means eight distinct external relationships.
    const targets = Object.keys(converted.out.footnoteRels).map(
      (k) => converted.out.footnoteRels[k],
    );
    assert.equal(targets.length, new Set(EXPECTED_FOOTNOTE_URLS).size);
  });

  await t.test("gives every footnote a blank line after it", () => {
    // A page of long government URLs packed flush together is unreadable.
    converted.out.footnotes.forEach((fn) => {
      assert.equal(fn.trailingBlanks, 1, "footnote " + fn.id);
    });
  });

  await t.test("produces a package that reopens as already converted", () => {
    return engine.analyzeDocx(converted.buffer).then((again) => {
      assert.equal(again.docTitle, "Cybersecurity Challenges for Australian Businesses");
      // Everything already converted; nothing left to find.
      assert.equal(again.citationCount, 0);
      // And the result is recognised for what it is, rather than reported as
      // a document with a missing reference list.
      assert.equal(again.alreadyConverted, true);
      assert.equal(again.notes.length, EXPECTED_FOOTNOTE_URLS.length);
      assert.deepEqual(
        again.notes.map((n) => n.url),
        EXPECTED_FOOTNOTE_URLS,
      );
      // The placeholder left in the bibliography is not read back as a source.
      assert.equal(again.refs.length, 0);
    });
  });

  await t.test("is stable when run twice", async () => {
    // Converting the same input twice must give the same output — no counter
    // or id that creeps between runs.
    const second = await convert(buffer);
    assert.deepEqual(second.out.paragraphs, converted.out.paragraphs);
    assert.deepEqual(
      second.out.footnotes.map((f) => f.url),
      converted.out.footnotes.map((f) => f.url),
    );
  });
});
