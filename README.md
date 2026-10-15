# MIWIM

Turns the in-text citations in a Word document into real, clickable footnotes.

You drop in a `.docx` that cites its sources the ordinary way — `(ASD 2025)`,
`(Smith & Jones, 2021)`, `Department of Home Affairs (2026a)` — and get back the
same document with each of those citations replaced by a numbered footnote
linking to that source. Not a superscript character typed in by hand: a genuine
Word footnote, the kind `Insert → Footnote` makes, which renumbers itself when
you add one, appears at the bottom of the page when you print, and survives
export to PDF.

Everything happens in your browser. The document never leaves your computer —
there is no server, no upload, and no account.

```
┌─ Original ─────────────────────┐  ┌─ Converted ────────────────────┐
│ …cost of cybercrime rose 14%   │  │ …cost of cybercrime rose 14%   │
│ to about $56,600 (ASD 2025).   │  │ to about $56,600¹.             │
│                                │  │                                │
│ References                     │  │ Footnotes                      │
│ ASD (Australian Signals        │  │ 1. https://cyber.gov.au/…      │
│ Directorate) 2025, Annual…     │  │                                │
└────────────────────────────────┘  └────────────────────────────────┘
```

---

## Contents

- [Getting started](#getting-started)
- [What it does](#what-it-does)
- [The rule that governs everything](#the-rule-that-governs-everything)
- [What counts as a citation](#what-counts-as-a-citation)
- [Page numbers and other pinpoints](#page-numbers-and-other-pinpoints)
- [How matching works](#how-matching-works)
- [What happens to your reference list](#what-happens-to-your-reference-list)
- [Documents that already have footnotes](#documents-that-already-have-footnotes)
- [Edge cases it handles](#edge-cases-it-handles)
- [Limitations](#limitations)
- [Architecture](#architecture)
- [Tests](#tests)
- [Extending it](#extending-it)
- [License](#license)

---

## Getting started

**To use it**, open `index.html` in a browser. That is the whole installation.
Some browsers stop `file://` pages from reading local files, so if the drop zone
does not respond, serve the folder instead:

```bash
npm run serve       # then open http://localhost:8080
```

It listens on `127.0.0.1` only. The folder being served is the one you are
dropping documents into, and a tool whose promise is that your document stays on
your machine should not hand the folder to everyone on the wifi. To reach it from
a phone on the same network, ask for that explicitly:

```bash
HOST=0.0.0.0 npm run serve
PORT=8081 npm run serve      # if 8080 is taken
```

(`npm` echoes the command before running it — that extra line is npm, not the
program.)

Drop a `.docx` on the left pane. The right pane shows what the conversion will
produce, with every citation and its reference highlighted together when you
hover either one. When it looks right, click the save button in the corner to
download `yourfile-footnotes.docx`.

**To work on it**, you need Node 18 or newer:

```bash
npm install         # @xmldom/xmldom and jszip, both dev-only
npm test            # 254 checks
npm run test:browser  # 15 more, in a real headless Chrome
```

The two dependencies exist so the test suite can run without a browser. The page
itself uses the browser's own XML parser and loads JSZip from a CDN, so the page
— unlike your document — does need the network the first time it is opened.

---

## What it does

Given a document, MIWIM:

1. Finds the reference list — a paragraph reading `References`, `Bibliography`,
   `Works Cited` or similar, and everything under it up to the next section.
2. Parses each entry into an author, a year, and a source URL.
3. Scans the body for brackets that are citations, and works out which entry
   each one points to.
4. Replaces each citation with a footnote reference, and writes a matching
   footnote containing that source's URL as a live link.
5. Collapses the reference list, since every entry now lives in a footnote.

Steps 3 and 4 are where the difficulty is, and most of this README is about
them.

Each footnote is written with a blank line after it. Word packs footnotes flush
against each other at the foot of the page, which is fine when each is a short
line and unreadable when each is a long URL. The blank line is an empty
`FootnoteText` paragraph — the same thing you get by pressing Enter at the end
of a footnote.

**The preview is the document.** Whatever the right pane numbers 4, Word will
number 4. That is not a nicety; it is the only way to check the output without
opening Word, and the two halves are computed from the same functions over the
same text for exactly that reason. `tests/agreement.test.js` exists because they
once disagreed.

---

## The rule that governs everything

> A bracket is converted only when every part of it is accounted for.

The author and the year become the footnote. A pinpoint like `p. 44` is carried
into the footnote. Anything else in the bracket — a second citation, a
qualification, a note to the reader — means the bracket is **left exactly as
written** and reported to you in the interface.

The reason is that a footnote marker replaces the whole bracket. There is no
half-conversion available: whatever is inside and unaccounted for is deleted from
the document, and it is deleted invisibly, because the sentence still reads
correctly afterwards. `(Smith 2020, as cited in Jones 2021)` becoming a footnote
to Smith is not a small error in an academic document.

So:

| Bracket | What happens |
|---|---|
| `(Smith 2020)` | converted |
| `(Smith 2020, p. 44)` | converted; `p. 44` goes into the footnote |
| `(Smith 2020; Jones 2021)` | converted; two markers, two footnotes |
| `(Smith 2020, emphasis added)` | **left alone**, reported |
| `(Smith 2020, Jones 2021)` | **left alone**, reported |
| `(Smith 2020, as cited in Jones 2021)` | **left alone**, reported |
| `(2024-25 data; Smith 2020)` | **left alone**, reported |
| `(Cth)`, `(No 2)`, `(a)` | skipped silently — not citation-shaped |

The last row is the difference between "left alone and reported" and "skipped".
A legal document contains dozens of `(Cth)` brackets and flagging each one would
bury the real problems, so a bracket with no plausible year in it is not
mentioned. A bracket that *does* look like a citation and was not converted
always is.

`tests/locators.test.js` states this as an invariant rather than a list of
examples: 2,700 generated brackets, and for every one that converts, the letters
and digits it contained must equal the author, the year, the pinpoint and the
signal words a footnote makes redundant.

---

## What counts as a citation

The hard part is not finding brackets. It is deciding which brackets are
citations, and this matters more than it sounds — a legal or policy document is
full of brackets that look citation-shaped and are not. From one real report:

```
…including the Privacy Act 1988 (Cth), the Privacy and Other Legislation
Amendment Act 2024 (Cth) and the Cyber Security Act 2024 (Cth).

…sections 912A(1)(a), (d) and (h) of the Corporations Act 2001 (Cth)

In Australian Clinical Labs Limited (No 2) [2025] FCA 1224 the Court…

Prepared and presented by: [Student Name 1] ([Student Number 1])
```

Eleven brackets, not one of them a citation. A converter that grabs anything
parenthesised turns that page into nonsense.

A bracket must contain a plausible publication year, and each
semicolon-separated part must resolve to an actual entry in the reference list.
Which four-digit numbers count as years is its own problem. These are all
rejected:

| Text | Why |
|---|---|
| `the 2024-25 financial year` | either half of a numeric range |
| `2023-2030 Australian Cyber Security Strategy` | same |
| `[2025] FCA 1224` | 1224 is outside the plausible year range |
| `C2024A00098` | digits glued to letters |
| `…/annual-cyber-threat-report-2024-2025` | hyphen-joined slug inside a URL |
| `$1,200` | digits inside a longer run |

That last one deserves a note. Reference entries routinely end in a URL, and
URLs are full of four-digit numbers. The obvious `/\d{4}/` search reads
`report-2024-2025` as the publication year, which quietly breaks matching for
every reference that carries a dated link.

`n.d.` is recognised as a year in its own right — it is the standard stand-in
for a missing date, and without it those references can never be matched.

Signal phrases are stripped: `(see Smith 2020)`, `(cf. ASD 2025)`,
`(e.g. Jones 2021)` all convert, and the signal word goes with the bracket,
because a footnote is the citation and "see" was only ever pointing at it.

---

## Page numbers and other pinpoints

A citation that quotes something carries the place it was quoted from, and that
pinpoint is how a reader checks the marker against the source. It is the last
thing that should go missing.

These are recognised, carried into the footnote after the URL, and left outside
the hyperlink, because a page number is not part of an address:

```
p. 44        pp. 44-46      page 44      pages 44-46
ch. 3        chapter 3      s. 12        sec. 12
para. 7      paras. 7-9     n. 5         art. 5
vol. 2       vol. 2, p. 9   no. 4        fig. 2
table 3      p. iv          pp. xii-xv   44        44-46
```

Two boundaries in that grammar are deliberate.

**A bare number is a page, unless it is year-shaped.** `(Smith 2020: 44)` and
`(Smith 2020, 44)` are Harvard page references. `(Smith 2020, 2021)` is two
works by Smith, and reading `2021` as a page would both attach the wrong note
and lose the second citation — so it falls under the rule above and the bracket
is left alone.

**A pinpoint word must be separated from its number.** `p. 44`, `p.44` and
`p 44` are all fine; `p44` is not. Without that requirement, `v` for verse
followed by the roman numeral `ivid` reads the word *vivid* as a pinpoint, and
the letters i, v, x, l, c, d and m spell enough English — *civil*, *mild*, *mix*
— that a bare roman numeral is refused outright and only accepted after a word.

A misread pinpoint is a mild failure, because it still ends up preserved in the
footnote. The failure this grammar is tuned against is the other one.

---

## How matching works

Once a bracket is known to be a citation, it has to find its reference. The
matcher compares normalised token lists — lowercased, with initials, `of`,
`the`, `and` and punctuation dropped — so that `Smith, J. A.` and `Smith`
compare equal, and `Smith & Jones` matches `Smith and Jones`.

Four things make this harder than a string comparison.

**Acronyms.** Australian government style writes the acronym first and expands
it in brackets:

```
ASD (Australian Signals Directorate) 2025, Annual Cyber Threat Report…
```

but cites it as `(ASD 2025)`. Treating the entry as one flat bag of words means
the citation matches nothing. Instead the name outside the brackets and each
bracketed expansion each become a key of their own, and a citation matching any
key wins. This works in both directions: `(ASD 2025)` and `(Australian Signals
Directorate 2025)` both find the same entry.

The bracketed form is the easy case, because both spellings are present. The
plain form does not cooperate:

```
Office of the Australian Information Commissioner 2025a, Latest statistics…, OAIC, Sydney
```

Here the acronym appears only as the *publisher*, after the year, outside the
author segment entirely — so `(OAIC 2025a)` has no literal text to match
against. The acronym is therefore derived from the name as well. It falls out of
the token list for free, because the words a writer drops when forming an
acronym are the same ones normalisation already removes:
`[office, australian, information, commissioner]` → `oaic`. A derived key that
duplicates one the entry already carries is discarded rather than added twice.

Derived keys score below literal ones, so an entry that genuinely carries the
cited words still beats an entry that merely shares its initials.

**Same author, same year.** When someone cites three OAIC documents from 2025,
the reference list distinguishes them with letters, and `(OAIC 2025b)` must land
on exactly one:

```
OAIC 2025a, Australian Clinical Labs ordered to pay penalties…
OAIC 2025b, Latest Notifiable Data Breach statistics…
OAIC 2025c, Statutory tort for serious invasions of privacy…
```

Suffixes are tracked separately from the year, and a mismatch is a hard
rejection rather than a lower score.

**Narrative citations.** `Department of Home Affairs (2026a) explains…` puts the
author outside the bracket, where it is doing grammatical work in the sentence.
When a bracket holds only a year, the author is recovered by reading backwards
through the preceding capitalised words and name connectors — and then only the
bracket is replaced, leaving the name in the sentence where it belongs.

The backward scan cannot tell where the name begins, and this is the interesting
part. The word in front of an author is capitalised about as often as not:

```
As Smith (2020) argues…          While Smith (2020) argues…
However, Smith (2020) found…     Recently Smith (2020) revisited…
…published by the OAIC. Smith (2020) notes…
```

Stopping at punctuation does not help: nothing distinguishes `Inc.` or `OAIC.`
from a full stop, and the commas in `Smith, Jones and Wu (2020)` are part of the
author list. So the matcher tries the **suffixes** of what the scan collected,
longest first — `["as","smith"]`, then `["smith"]` — and takes the first that
matches anything. Longest-first keeps the three-author entry winning over the
one-author entry for `Smith, Jones and Wu (2020)`.

Suffixes only, never prefixes, and that is what makes it safe: dropping words
from the front can only discard something that was never part of the name, while
dropping them from the back would let `the Smith Review (2020)` match a
`Smith 2020` entry, which is a different source. There is a test asserting that
specific non-match.

**Scoring rather than first-hit.** Candidates are scored so an exact author match
beats a partial one, which is what keeps `(ASIC 2026)` and `(Federal Court of
Australia 2026)` apart when both are 2026 entries.

---

## What happens to your reference list

Once every entry has become a footnote, leaving the list in place duplicates it.
So the section collapses to a single line:

```
Bibliography
[Insert citations and references using the footnotes]
```

The heading stays, so the document keeps its structure.

**But only if every entry was converted.** A reference that no citation matched
has no footnote to live in, so it stays exactly where it is, and the section is
not collapsed. Deleting a source from someone's bibliography because the matcher
could not find its citation would be the worst thing this tool could do, and it
is the kind of loss you might not notice until after submitting. The interface
tells you which references were never cited so you can check whether that is
deliberate.

Read that alongside the first bullet under [Limitations](#limitations). Between
the two of them, a converted document has the source's URL and not its author,
title or publisher — and that is a decision about what the output should be, not
an accident.

---

## Documents that already have footnotes

A document does not always arrive as a clean before-picture. It may already
contain footnotes — inserted by hand, or left by an earlier run of this tool.
Three cases, each handled differently.

**Only footnotes, no references.** Nothing to convert, so nothing is converted.
The document is shown as it is, with its existing footnotes listed and each
marker in place, and the download is disabled. This is what a document
MIWIM already processed looks like when you feed it back in, and telling
someone "no references found" there would suggest something is broken when the
document is finished.

**Only references, no footnotes.** The ordinary conversion.

**Both.** The references are converted; the existing footnotes are left exactly
as they are. New footnotes take ids above the highest one already in use, so a
hand-written note is never overwritten.

The numbering matters here and is easy to get wrong. Word numbers footnotes by
position in the document, not by the order they were created. A document whose
first three footnotes come before its first citation will number that citation
**4**. So the preview numbers existing and new notes together, in reading order,
and the footnote list shows them interleaved — what you see is what Word will
produce.

Existing footnotes also get the same blank line the new ones get, so the page
does not end up with two different spacings. That only happens when something
was actually converted — a document nobody asked to change is not reformatted —
and a note that already has its blank line does not get a second one, so running
a document through repeatedly does not keep growing the gaps.

Two smaller behaviours fall out of this:

- **Orphaned footnotes** — defined in the file but referenced nowhere in the
  text, usually left by an edit — are listed separately and left untouched.
- **The placeholder** MIWIM leaves in a collapsed reference list is not read
  back as a reference on a second pass. Neither is any other bracketed template
  line like `[Add your references here]`.

---

## Edge cases it handles

Most of these are invisible until they break something.

**Runs split mid-citation.** Word splits text into runs at every formatting and
revision boundary, so `(Smith 2020)` is routinely stored as three separate runs.
Text is read from a single traversal, and the edit map is built from that same
traversal, so the character offsets always land on a real run. Deriving them
separately — the usual approach — makes the two disagree the moment a paragraph
contains a hyperlink, and the replacement then cuts the sentence in the wrong
place.

**Citations inside hyperlinks.** The marker is lifted out of the link, because a
footnote number rendered as part of a hyperlink is wrong to click and wrong to
look at. Text on either side of the citation stays linked: the link is split in
two around the marker, both halves keeping the same relationship, so the number
lands where the citation was rather than at the end of the phrase.

**Hyperlinked URLs.** A reference's URL is almost always a `w:hyperlink` element,
one level below the runs a naive walk inspects. Missing it corrupts both the
reference text and the URL. When the visible link text is not the address — a
title, or "View source" — the relationship target is the only place the real URL
exists, so that is preferred over anything scraped from the text.

**Angle-bracketed URLs.** AGPS style writes `viewed 11 August 2026,
<https://example.gov.au/report>.` The `>` and the `.` are not part of the
address; the trailing `/` in a path is. Balanced parentheses inside a URL
(`/wiki/Mercury_(planet)`) are kept, while an unbalanced one that merely wraps it
is dropped.

**Subdivided reference lists.** A bibliography is often broken into groups:

```
Bibliography                 (Heading 1)
  Legislation and Cases      (Heading 2)
  Secondary Sources          (Heading 2)
```

Ending the list at the first styled heading stops it dead at the first group
label, one paragraph in — so the section reads as empty and the interface reports
that nothing followed the heading, while the entries sit in plain sight below. A
heading only ends the list when it is at the same level as the reference heading
or shallower. Deeper headings are group labels: they are not read as entries, and
when the list collapses they go with it rather than being left heading an empty
group. A heading naming a section that genuinely follows a reference list
(`Appendix A`) still ends it whatever its level, and when the reference heading is
hand-formatted rather than styled there is no level to compare against, so any
styled heading ends the list as before.

**Generated tables of contents.** A TOC duplicates every heading in the document,
"Bibliography" included. Treated as the real section, the reference list becomes
a single page-number entry. TOC paragraphs are detected by style and by field
code, and skipped.

**Paragraphs inside tables.** Plenty of report templates lay their body out in a
single-cell table. Walking only the direct children of the body finds nothing at
all in those documents.

**Documents with no title line.** The first non-empty paragraph is rendered as
the document's title rather than as body text — unless it carries a footnote or a
citation, in which case it is body text like any other paragraph. A chapter, an
extract or a draft pasted into a fresh file has no title, and treating its first
sentence as one used to drop that sentence's citation out of the preview while
the download converted it anyway.

**Repeated sources.** Word numbers footnotes by position, so a source cited three
times needs three marks. It gets three footnotes pointing at the same URL, but
only one relationship — duplicate relationship ids make Word refuse to open the
file. A source cited twice at *different* pages gets two notes, because the pages
are the only thing distinguishing them.

**Missing parts.** A `.docx` saved without a footnotes part gets one created,
along with the relationship and the `[Content_Types].xml` override. Word only
writes the `FootnoteText` / `FootnoteReference` style definitions once a document
actually contains a footnote, so those are added when missing — otherwise the
footnote text renders as ordinary body copy.

**Tracked changes.** Text inside `w:del` is a tracked deletion and is not part of
the visible paragraph, so it is excluded from matching.

---

## Limitations

Worth knowing before you rely on it.

- **Footnotes carry the URL, not the full reference — and the reference list is
  then collapsed.** Those two behaviours together mean a converted document holds
  the address of each source and not its author, title or publisher. The
  reasoning for URL-only footnotes is that those details belong in a reference
  list and the link is what you actually want to click; the collapse exists
  because leaving a converted entry in the list duplicates it. The combination is
  worth deciding on deliberately rather than discovering: if you want the full
  citation kept, `buildFootnoteContent` already has the branch that writes it,
  which is what happens for a reference with no URL.
- **Author-date styles only.** Harvard, APA and AGPS work. Numeric styles
  (Vancouver `[1]`, IEEE) and existing footnote styles (Chicago notes, OSCOLA)
  are not converted — there is nothing to convert, since they are already notes.
- **One reference list.** A document with per-chapter bibliographies will only
  have its last one recognised.
- **Citations after the reference list are not converted.** The body ends where
  the reference section begins, so a citation in an appendix is left alone, and
  it is not reported either.
- **Endnotes are not produced.** Footnotes only. Endnotes a document already has
  are left alone but are not shown in the preview.
- **Ibid., op. cit. and shortened forms** are not understood.
- **Comma-separated citations in one bracket** are not split. `(Smith 2020;
  Jones 2021)` converts; `(Smith 2020, Jones 2021)` does not, because a comma
  after the year is far more often a pinpoint than a second author. It is
  reported rather than guessed at.
- **Reference entries must be one paragraph each.** An entry broken across two
  paragraphs by a manual line break is read as two entries.
- **The page needs the network once.** JSZip comes from cdnjs and the three
  webfonts from Google Fonts. Your document goes nowhere — `tests/app.test.js`
  asserts there is no `fetch`, no `XMLHttpRequest`, no `sendBeacon` and no form
  in the page — but those two `<head>` requests mean it will not work fully
  offline, and the conversion itself does not run at all without JSZip.

If a citation is not matched, it is left untouched and listed in the interface.
Nothing is ever removed on a guess.

---

## Architecture

Four files, in dependency order:

```
miwim.core.js    pure functions — no DOM, no zip, no I/O
      ↑
miwim.docx.js    OOXML reading and writing
      ↑
index.html       interface

serve.js         optional static server for local development
```

**`miwim.core.js`** answers two questions: given a reference entry, who
wrote it and when; and given a paragraph, which stretches of it are citations
pointing at which reference. It works on plain strings and plain objects, which
is why it is testable without a browser and why the matching rules can be read
without wading through XML.

**`miwim.docx.js`** does everything involving the package: unzipping,
walking paragraphs, resolving relationships, splicing footnote references into
runs, writing the footnotes part. Its dependencies are injected rather than
reached for —

```js
const engine = createDocxEngine({ DOMParser, XMLSerializer, JSZip, core });
```

— so the same code runs in the browser with native `DOMParser` and under Node
with `@xmldom/xmldom`, with no branching and no shims. It exposes two calls:

```js
const analysis = await engine.analyzeDocx(arrayBuffer);
// → { docTitle, bodyParas, refs, hasRefsSection, unmatched, citationCount,
//     uncitedRefs, notes, existingFootnotes, orphanFootnotes,
//     existingFootnoteCount, alreadyConverted }

const { zip, summary } = await engine.buildConvertedDocx(arrayBuffer, analysis.refs);
// → summary: { footnotesCreated, referencesConverted, referencesKept,
//              referenceSectionCollapsed, existingFootnotesRespaced, unmatched }
```

`notes` is the useful one for display: every footnote a reader will meet, both
the document's own and the ones this conversion would add, numbered together in
document order. `analysis.notes.length` and `summary.footnotesCreated` are
required to agree, and `tests/agreement.test.js` holds them to it.

Passing `analysis.refs` back into the build is what lets you edit a reference —
in practice, paste in a URL for one that had none — and have the conversion use
it. Refs are matched by id, so a filtered or reordered list still lines up.

**`index.html`** holds the interface and nothing else: state, rendering, drag and
drop, the download. No parsing logic lives here. The favicon is the wordmark
redrawn as a standalone SVG and inlined as a data URI — it cannot reuse the
topbar's inline logo, which colours itself from the page's CSS variables and sets
the numeral in Fraunces, because a favicon renders with no stylesheet and no
webfont. A test pins its palette to those same variables so the two cannot
quietly drift apart.

**Three ordering rules** hold the conversion together, and all three are
load-bearing:

1. Footnotes are created in reading order, because that is what determines the
   numbering a reader sees.
2. Body edits are applied right-to-left, because each replacement changes the
   paragraph's length and would shift the offsets every later group was measured
   against.
3. Each paragraph is read once and every group in it measured against that single
   snapshot, for the same reason.

Doing any of them the other way produces output that looks fine on a document
with one citation per paragraph and falls apart on a real one.

---

## Tests

```bash
npm test              # 254 checks across ten files
npm run test:watch
npm run test:browser  # 15 checks in a real headless Chrome
```

Node's built-in runner. No framework, no config, no dependency beyond the two
that let the OOXML layer run headlessly.

| File | Covers |
|---|---|
| `tests/core.test.js` | year detection, author keys, URL extraction, citation matching, the brackets that are not citations |
| `tests/docx.test.js` | reading paragraphs from real packages, run splitting, hyperlinks, TOC handling, footnote generation, package integrity |
| `tests/footnotes.test.js` | documents that already have footnotes, joint numbering, spacing, repeat conversions |
| `tests/locators.test.js` | pinpoints, and the invariant that a converted bracket adds up |
| `tests/narrative.test.js` | authors recovered from outside the bracket |
| `tests/agreement.test.js` | the preview and the download being one document |
| `tests/hyperlinks.test.js` | markers placed inside a link |
| `tests/integration.test.js` | a full report end to end |
| `tests/app.test.js` | the browser wiring — module loading, element ids, the favicon, and the privacy claim: no storage, no `fetch`, no form |
| `tests/browser/run.js` | the real files in real Chrome, driven over loopback |
| `tests/serve.test.js` | the dev server, as a child process over a socket |
| `tests/helpers/docx-fixture.js` | builds and reads `.docx` packages in memory |

Fixtures are **built in code rather than checked in as binaries**. A test that
loads `sample.docx` and asserts a number tells you nothing about which quirk it
guards; a test that builds its own document states the quirk in the test:

```js
await t.test("stitches text back together across split runs", async () => {
  const buf = await buildDocx({
    body: [
      p("A Report", { style: "Title" }),
      p(["Growth was strong (Smi", "th ", "2020) last year."]),
      p("References", { style: "Heading1" }),
      p("Smith, J 2020, A Book, Press."),
    ],
  });
  const r = await convert(buf);
  assert.equal(r.out.paragraphs[1], "Growth was strong{1} last year.");
});
```

Each string in that array becomes its own `w:r`, reproducing exactly the run
split Word introduces. The helper also handles hyperlink wrappers, tables,
content controls, field codes and list numbering, so any structural quirk can be
reproduced in a line or two.

`tests/integration.test.js` builds a document carrying every difficulty at once —
Harvard bare years, acronym expansions, `2025a`/`b`/`c` suffixes, eleven
non-citation brackets, a table of contents, angle-bracketed URLs, a citation
split across runs — and asserts the exact footnote ordering and URLs, that the
`(Cth)` brackets survive, and that converting twice gives identical output.

Three of the suites are worth reading for how they are built rather than what
they assert. `tests/locators.test.js` generates 2,700 brackets and checks a
property instead of enumerating answers. `tests/serve.test.js` asserts the server
is alive before asserting it is unreachable from elsewhere, because a dead server
refuses every connection and would otherwise pass the check most convincingly.
`tests/browser/run.js` serves the project to a real Chrome, because everything
else loads these files through `require` and takes the branch a user never runs.

---

## Extending it

**A new reference-list heading**: add it to `HEADING_WORDS` in
`miwim.core.js`. A new section that ends a reference list:
`SECTION_END_WORDS`.

**A new pinpoint abbreviation**: `LOCATOR_WORDS`, same file. Read the two
paragraphs above it first — the separator requirement and the roman-numeral
restriction are both there because of a specific misreading.

**A new citation shape**: `parseCitationUnit` turns one bracket segment into
`{ tokens, year, yearSuffix, locator, residue }`, and
`scoreUnitAgainstReference` decides how well that fits a reference. Both are pure
and directly testable. If your new shape leaves anything in `residue`, the
bracket will be left alone and reported, which is the safe default — make it
parse rather than working around it.

**Full reference text in footnotes** rather than just the URL: that is
`buildFootnoteContent` in `miwim.docx.js`, which already has the branch —
it is what happens when a reference has no URL.

**Endnotes instead of footnotes**: the shape is the same. `w:endnoteReference`
for `w:footnoteReference`, `word/endnotes.xml` for `word/footnotes.xml`,
`EndnoteText` for `FootnoteText`.

Before changing matching behaviour, run the tests to see which rule you are about
to trade away. Several of them exist because a plausible-looking simplification
broke a real document, and several more because a plausible-looking fix broke a
different one.

---

## License

Free for personal, non-commercial use. You may use, study and modify this
software for yourself, but you may not use it commercially, and you may not
share or publish it, in its original form or modified. See [LICENSE](LICENSE)
for the full terms.
