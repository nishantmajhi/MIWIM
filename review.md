# Code Review — MIWIM

**Remediated 2026-09-26 on branch `cleanup/2026-09-26`. Not pushed.** This file
was a to-do list written 2026-09-04; it is now the record of the pass. The
original review is kept in full at the bottom, because two of its four findings
were right, one was already done, and the reasoning it used is worth keeping.

Before: 203 tests, all passing. After: 254 plus 15 in a real browser, all
passing, run three consecutive times. Four defects fixed, none of which the
first review found.

**Renamed from Marginalia to MIWIM on 2026-10-03**, because the name Marginalia
is already in use by other projects on the web. Everything below has been
updated to the new name, including this record. The directory, the page title
and wordmark, `marginalia.core.js` and `marginalia.docx.js` (now `miwim.*.js`),
the `window.Marginalia` global (now `window.MIWIM`) and the package name all
changed. The `mrg-` CSS class prefix was left as it is: it is internal and
renaming it would touch every rule in the page for no visible difference. The
remote is still `https://github.com/nishantmajhi/Marginalia.git`; renaming the
GitHub repository is the author's to do.

---

## Corrections to the first review

**Finding 3 was already implemented, and had been from the start.** The review
asked that unmatched or ambiguous citations be "reported to the user rather than
skipping them silently", and called silent omission "the dangerous outcome".
`findCitationGroups` has returned an `unmatched` array since the first commit;
`analyzeDocx` and `buildConvertedDocx` both collect it, `index.html:1069` renders
it, and `tests/core.test.js:502` asserts it. The review was reading a real risk
and did not check whether the code had already answered it. One `grep -n
unmatched` would have shown that it had.

That matters beyond the bookkeeping, because the review then stopped. Having
satisfied itself that unmatched citations were the danger, it never asked the
harder version of its own question: what happens to a citation that *does* match
but is not entirely a citation. That was the worst defect in the project.

**Finding 2 was right, and its list was where the bug was hiding.** It asked for
the edge cases to be "listed in the README and covered in tests, if not already"
— naming page suffixes, signal phrases, multiple works in one bracket and
non-citation brackets. Three of those five were already handled well. The first
one, `(Smith 2020, p. 44)`, was not handled at all: it deleted the page number.
The review wrote the reproduction down and did not run it.

**Finding 1 was right** and is fixed, along with a second defect in the same file
the review did not mention.

**Finding 4** (`miwim.docx.js` is 1,339 lines) is left alone deliberately.
It is now 1,377. The seam the review praises is between `core` and `docx`, and
that seam is real and load-bearing; splitting the OOXML file further by part
would separate code that has to agree about run offsets. Noted and not done.

**Finding 5**, no LICENSE, is done.

---

## What was fixed

### P1 — a bracket was converted with its contents deleted

`parseCitationUnit` took the first year-shaped token in a bracket segment and
discarded everything after it. Because a footnote marker replaces the whole
bracket, everything after the year left the document:

| Bracket | Became | Lost |
|---|---|---|
| `(Smith 2020, p. 44)` | a footnote | the page number |
| `(Smith 2020, pp. 44-46)` | a footnote | the page range |
| `(Smith 2020, ch. 3)` | a footnote | the chapter |
| `(Smith 2020, emphasis added)` | a footnote | the qualification |
| `(Smith 2020, my translation)` | a footnote | the attribution |
| `(Smith 2020, as cited in Jones 2021)` | a footnote to Smith | the secondary source |
| `(Smith 2020, Jones 2021)` | a footnote to Smith | a whole second citation |
| `(Smith 2020, 2021)` | a footnote to the 2020 work | a whole second citation |

Twenty shapes were tried and twenty lost text. Silently: the sentence still
reads correctly afterwards, so there is nothing for a proofreader to notice.

**The project's own README already forbade this.** "A bracket is only converted
when every part of it is a citation that matched a reference ... If half of
`(2024-25 data; Smith 2020)` matches and half doesn't, nothing is converted —
replacing the bracket would silently delete the words '2024-25 data' from
someone's document." That rule was enforced across semicolons and nowhere else,
and no test anywhere asserted that a bracket had to add up.

Fixed by classifying the tail after the year. A pinpoint is carried into the
footnote, after the URL and outside the hyperlink. Anything else is residue and
the bracket is left exactly as written and reported — the behaviour the
semicolon case already had.

Two boundaries in the new grammar were chosen rather than fallen into, and both
are in `tests/locators.test.js`:

- **A bare number is a page unless it is year-shaped.** `(Smith 2020: 44)`
  converts; `(Smith 2020, 2021)` is two works and is reported.
- **A pinpoint word must be separated from its number.** Otherwise `v` for verse
  plus the roman numeral `ivid` reads *vivid* as a pinpoint. A bare roman
  numeral is refused outright, because i, v, x, l, c, d and m spell *civil*,
  *mild*, *mix* and a dozen other words.

The test states the rule rather than the examples: 2,700 generated brackets, of
which 435 convert, and for each of those the bracket's letters and digits must
equal the author, the year, the pinpoint and the signal words a footnote makes
redundant. The non-vacuity floor is part of the check, because the invariant is
satisfied most completely by a matcher that converts nothing.

### P1 — narrative citations failed on ordinary English

`As Smith (2020) argues` produced no footnote. Nor did `However, Smith (2020)
found`, `While`, `Although`, `Recently`, `Indeed`, `Thus` or `See` — eleven of
eighteen ordinary sentences tried. The citation was reported unmatched and the
reference was reported as never cited, so a document written in normal academic
prose converted a fraction of what it should and then told the person its
sources were uncited.

`precedingAuthorTokens` reads backwards collecting capitalised words and name
connectors. Nothing in that rule stops it at the start of a sentence, because
the first word of a sentence is capitalised too, so it handed `["as","smith"]`
to a matcher that compares for equality, prefix or subset — and `["as","smith"]`
is none of those against `["smith"]`. The same thing happened across a sentence
boundary: `…published by the OAIC. Smith (2020) notes` gave `["oaic","smith"]`.

Stopping at punctuation is not the fix, and this is the part worth remembering.
Nothing distinguishes `Inc.` or `OAIC.` from a full stop, and the commas in
`Smith, Jones and Wu (2020)` are part of the author list — a comma rule would
have broken a case that worked. The matcher now tries the **suffixes** of what
the scan collected, longest first, and takes the first that matches anything.

Suffixes only, never prefixes, and that asymmetry is the safety argument:
dropping words from the front can only discard something that was never part of
the name, while dropping them from the back would let `the Smith Review (2020)`
match a `Smith 2020` entry, which is a different source. Attaching the wrong
source is worse than attaching none. `tests/narrative.test.js` asserts that
specific non-match, and that assertion passes against the unfixed code too — it
is there to catch the fix over-reaching, not to catch the bug.

### P2 — the preview and the download were different documents

`analyzeDocx` renders the first non-empty paragraph as the title and skips it.
`buildConvertedDocx` has no notion of a title and converts every paragraph. The
doc comment on `buildConvertedDocx` claims the two passes "agree on numbering
and matching". In a document that opens straight into body text they did not.

Given two paragraphs with one citation each, the preview reported one footnote
and the download produced two, so every number the person checked was one lower
than Word would show. Worse, the interface listed the first paragraph's
reference under **never cited** — the label the README explains at length as the
safeguard against losing a source — while the download converted that citation
and deleted the entry from the bibliography. The one thing the person was told
to check was the one thing that was wrong.

A document with no title line is not unusual: a chapter, a section extract, a
draft pasted into a fresh file.

Fixed narrowly. The first paragraph is the title unless it carries an existing
footnote or a citation this conversion would convert. Deliberately not "unless
it looks like body text": a real title often contains a bracketed year —
`Cyber Security in Australia (2025)` — and reclassifying those would change the
preview for documents that were never broken. An unmatched bracket alone does
not disqualify a title, for the same reason. `tests/agreement.test.js` pins the
ordinary title case so the narrowing cannot drift.

### P2 — the dev server served the LAN, and one request killed it

Both in `serve.js`, which had no tests at all.

**It bound every interface.** `server.listen(PORT, cb)` with no host argument
listens on `0.0.0.0`. Confirmed by fetching `/index.html` over this machine's
LAN address. The traversal guard confines it to the project folder, and the
project folder is exactly where you drop the `.docx` files you are testing with.
Now `127.0.0.1`, with `HOST` to ask for more. This was the first review's
finding 1.

**A single request stopped it.** `decodeURIComponent` throws `URIError` on any
half-written escape, from inside the request handler where nothing catches it:
`GET /%` exited the process with code 1. A path holding a NUL byte did the same
one line later, because `fs` rejects such a path by throwing rather than calling
back. Neither needs a crafted client — `curl http://host:8080/%` is enough, and
a link checker would find it by accident. Both answer 400 now.

Taken together these were one bug: anybody on the same wifi could stop the
server, and the first review found only the half that exposed the folder.

### P3 — a footnote marker landed after a whole hyperlink

`see the guide (ASD 2025) online` became `see the guide online¹` rather than
`see the guide¹ online`. Lifting the marker out of the link is correct — a
footnote number rendered as part of a hyperlink is wrong to click and wrong to
look at — but the text after the citation stayed inside the link while the
marker went after it. The link is now split in two around the marker, both
halves keeping the same relationship. Nothing was being lost, which is why this
is the least of the four; but a misplaced note is invisible to a reader.

### P3 — no LICENSE

`package.json` claimed MIT from the first commit with no licence text in the
repository. Added. The copyright year comes from the project's own first commit
(2026-08). The name form follows `Browser Extensions/ADBS`, the most recent
LICENSE added by one of these passes; `AutoHotKey/HorizontalScroll` uses the git
author name instead, so the archive is inconsistent and that is the author's to
settle.

---

## What was deliberately not done

**The URL-only footnote plus the collapsed reference list.** A converted
document keeps each source's address and not its author, title or publisher:
`buildFootnoteContent` writes only the URL when there is one, and
`collapseReferenceSection` then removes the entry it came from. Both halves are
deliberate and documented, and each is defensible alone. Together they mean the
full citation is gone from the document.

The old README gave "author, title and publisher belong in a reference list" as
the *reason* for URL-only footnotes, two screens after explaining that the
reference list is deleted. That is not a bug in the code; it is a decision about
what the output should be, and it is the author's. **Nothing was changed.** It is
now the first bullet under Limitations in the README, with a pointer at the
branch that would change it — `buildFootnoteContent` already writes the full
text for a reference with no URL, so switching is a one-line change if that is
wanted.

**Splitting `miwim.docx.js`.** See the corrections above.

**`unit.etAl`.** `parseCitationUnit` computes it, `tests/core.test.js:335`
asserts it, and nothing reads it. Left in place: it is one field, it records
something a future pinpoint-aware matcher would want, and deleting it would mean
deleting a passing test for no gain. Named here so the next reader knows it is
dead rather than load-bearing.

**Comma-separated citations in one bracket.** `(Smith 2020, Jones 2021)` is now
reported instead of silently truncated, but it is still not split into two
footnotes. Splitting on commas is not safe: a comma after the year is more often
a pinpoint than a second author, and the two are not distinguishable without
trying both against the reference list and preferring whichever matches. That
would be a real improvement and it is a bigger change than this pass. Reporting
it loses nothing in the meantime.

**Citations after the reference list.** `bodyEndIdx` stops the body at the
reference heading, so a citation in an appendix is neither converted nor
reported. Documented in Limitations, not changed — converting it would need the
reference-section boundary rethought, and appendix citations are rare.

**The `(Strategy 2030 Group 2021)` misparse.** An author name containing a
year-shaped number makes `findYearTokens` pick the wrong token. Before this pass
that produced a wrong footnote; now the leftover text is residue, so the bracket
is left alone and reported. The misparse is still there, but its consequence is
now safe, and fixing it properly means preferring the *last* year token in a
segment, which breaks `(Smith 2020, p. 44)`. Left as is.

**No UI or UX changes.** `index.html` is untouched. The new `locator` field on a
preview part and the new `locators` array on a group are additive and the page
ignores them.

---

## Notes for a future session

`npm test` is 254 checks across ten files; `npm run test:browser` is 15 more in
a real headless Chrome and takes about eight seconds. Both were run three
consecutive times, green every time. Chrome is at
`C:/Program Files/Google/Chrome/Application/chrome.exe`, version 153.

**Every regression check here was confirmed failing against commit `25a70b9`
first**, in a worktree with the new test files copied in: 36 of 51 node checks
failed, and 8 of 15 browser checks. The ones that passed against the unfixed
code are the guards against over-reaching, and they are labelled as such in the
test files.

Three things from this pass worth carrying elsewhere in the archive:

**A passing suite is evidence about coverage, not correctness — and the reverse
also holds for reviews.** 203 tests passed against code that deleted text out of
documents. They were good tests. Not one of them asked whether a conversion
preserved what it did not convert, and the README had stated that requirement in
prose for a year. **When a project documents a rule, write the test that checks
it before believing the rule holds.**

**A negative test passes when the thing under test is dead.** The first run of
`tests/serve.test.js` reported "binds the loopback interface only" as passing
against the unfixed code, which binds everything — because the malformed-escape
test two assertions earlier had already killed the server, and a dead server
refuses every connection. The suite now asserts liveness before asserting
unreachability. This is the `mountain_cipher` lesson again and it arrived
unprompted.

**A harness bug looks exactly like a finding.** The browser driver's first run
reported "dangling Id1" against correct code, because the arithmetic slicing
`r:id="rId2"` dropped a character. Before believing a new harness, check its
failures against code you already trust.

The seam between `miwim.core.js` and `miwim.docx.js` is still worth
preserving — keep citation logic out of the XML file and XML out of the core.
Before changing the footnote XML, run the suite for a green baseline; Word is
unforgiving about part ordering and relationship ids, and those tests encode what
it accepts.

---
---

# The original review, 2026-09-04

Reviewed against the working tree. Scope: `miwim.core.js` (804 lines),
`miwim.docx.js` (1,339 lines), `serve.js`, `index.html`, `tests/` (five
suites), `package.json`, `README.md`.

## What this project is

A browser tool that converts in-text citations in a `.docx` — `(ASD 2025)`,
`(Smith & Jones, 2021)`, `Department of Home Affairs (2026a)` — into genuine Word
footnotes: the real kind that renumber themselves, appear at the bottom of the
printed page, and survive PDF export. Everything runs locally.

## Verdict

Strong work. Three things stand out:

- **Real tests, no test framework.** Five suites run through Node's built-in
  `node --test`, with a `test:watch` variant and a `helpers/` directory. Zero
  devDependency weight for the runner.
- **The split between `miwim.core.js` and `miwim.docx.js`** separates
  citation detection from Word's XML plumbing, which is the right seam — the
  hard, testable logic is isolated from the format handling.
- **The traversal guard in `serve.js` is correct.** It decodes the URL, joins
  against `ROOT`, then checks the *resolved* path still starts with
  `ROOT + path.sep`. Because `path.join` normalises `..` first, this actually
  works — many implementations of this check do it in the wrong order and do
  not. The comment even names the attack it is stopping.

The README is the best in the archive: it explains what the tool produces and
why a real footnote differs from a typed superscript, with a before-and-after
diagram.

## Findings

**1. `serve.js` listens on all interfaces.** `server.listen(PORT, ...)` with no
host argument binds to every interface, so the whole project directory is served
to anyone on the network. The traversal guard confines it to the project folder —
but that folder is where you will be dropping `.docx` files to test with, and
those are exactly the documents the tool exists to keep local. Pass
`"127.0.0.1"` as the host.

**2. Citation detection is inherently approximate, and the boundaries should be
documented.** Regex-based citation matching will always have edge cases, and
users need to know which ones. Worth listing in the README and covering in
tests, if not already:

- Page or section suffixes: `(Smith 2020, p. 44)`, `(Smith 2020, ch. 3)`.
- Signal phrases inside the parentheses: `(see Smith 2020)`, `(cf. Smith 2020)`,
  `(e.g. Smith 2020)`.
- Multiple works in one parenthetical: `(Smith 2020; Jones 2021)`.
- Parentheses that are not citations at all: `(a common misconception)`, `(2020)`
  standing alone.
- Text inside quotations that happens to contain a citation-shaped string.

Every one of those either produces a wrong footnote or misses a real one, and a
user proofreading a converted document needs to know where to look.

**3. Confirm the reference-list matching behaviour.** A footnote is only useful
if it links to the right source. What happens when a citation has no matching
entry in the bibliography, or matches two (`Smith 2020a` versus `Smith 2020b`)?
Silent omission would be the dangerous outcome for an academic document. The
tool should report unmatched citations to the user rather than dropping them.

**4. `miwim.docx.js` is 1,339 lines** of Word XML manipulation. That is
inherent to the problem — footnotes touch `document.xml`, `footnotes.xml`, the
relationships part, and the content types — but it is the file most likely to
become unpleasant. If it grows further, split by part.

**5. No `LICENSE`.**

## Cleanup

Nothing to remove. No debug output, no secrets, no analytics, no committed
artifacts. `package.json` has only devDependencies, keeping the runtime
dependency-free, which matches the "runs in your browser, nothing uploaded"
claim.

## Action checklist

- [x] **P2** Bind `serve.js` to `127.0.0.1`.
- [x] **P2** Document the citation patterns that are and are not detected; add
      tests for the edge cases listed above.
- [x] **P2** Report unmatched or ambiguous citations to the user rather than
      skipping them silently. *(Already implemented before this review was
      written — see the corrections at the top.)*
- [x] **P3** Add a `LICENSE`.

## Licence, 2026-10-02

The author chose one licence for all of their projects: free for personal,
non-commercial use, modifiable for the user's own use, and not to be
distributed, original or modified. No standard licence says that, so `LICENSE`
is a short custom text, the same in every project apart from the year and what
it excludes. The README's licence section says the same, and `package.json`
declares `SEE LICENSE IN LICENSE`, npm's form for a licence with no SPDX
identifier. Anything above that calls the licence an open question, or calls
the project MIT, is superseded.

## CI, 2026-10-02

`.github/workflows/ci.yml` runs `npm test` and the browser suite as two jobs
on `ubuntu-latest`. `package-lock.json` is gitignored here, so both use `npm
install` rather than `npm ci`, and no npm cache. Nothing has been pushed, so
it has not run on GitHub yet. Every job was run before committing, from a
fresh clone of this branch, the way a runner would run it: Linux jobs in a
throwaway Ubuntu 24.04 (the image `ubuntu-latest` uses) with the same tool
versions. 254 node checks.

This browser runner fails rather than skipping when it finds no Chrome, so it
needs no `SKIP` guard.
