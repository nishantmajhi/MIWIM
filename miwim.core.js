/*!
 * MIWIM — core citation engine
 * ---------------------------------
 * Pure JavaScript. No DOM, no zip, no I/O, no globals touched. Everything in
 * here works on plain strings and plain objects, which is what makes it
 * testable under Node without a headless browser and reusable by the OOXML
 * layer in miwim.docx.js.
 *
 * The whole job of this file is to answer two questions:
 *   1. Given the text of a reference-list entry, who wrote it and when?
 *   2. Given a paragraph of body text, which stretches of it are in-text
 *      citations, and which reference does each one point at?
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.MIWIM = root.MIWIM || {};
    root.MIWIM.core = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ---------------------------------------------------------------------
  // Vocabulary
  // ---------------------------------------------------------------------

  // Headings that introduce a reference list. Compared after normHeading(),
  // which strips digits and punctuation, so "7. REFERENCES:" normalises to
  // "references" and still matches.
  var HEADING_WORDS = [
    "references",
    "reference",
    "reference list",
    "list of references",
    "bibliography",
    "select bibliography",
    "works cited",
    "work cited",
    "literature cited",
    "sources cited",
    "reference list and bibliography",
  ];

  // Headings that commonly *follow* a reference list. Used to find where the
  // list ends, so trailing sections aren't swallowed as reference entries.
  // Any paragraph carrying a real Heading/Title style also ends the list —
  // this list only exists for documents whose headings are styled by hand.
  var SECTION_END_WORDS = [
    "appendix",
    "appendices",
    "annex",
    "annexes",
    "acknowledgement",
    "acknowledgements",
    "acknowledgment",
    "acknowledgments",
    "author contributions",
    "supplementary material",
    "supplementary materials",
    "funding",
    "conflict of interest",
    "conflicts of interest",
    "ethics statement",
    "data availability statement",
    "generative ai acknowledgement",
    "generative ai acknowledgment",
    "ai acknowledgement",
    "ai acknowledgment",
    "declaration of ai use",
    "declaration",
    "statement of authorship",
    "biographies",
    "author biography",
    "author biographies",
    "about the author",
    "about the authors",
    "vitae",
    "index",
    "glossary",
    "abbreviations",
    "list of figures",
    "list of tables",
    "further reading",
    "notes",
    "endnotes",
    "footnotes",
    "word count",
  ];

  // Dropped when normalising an author name to comparable tokens. These are
  // the words that differ freely between an in-text citation and its
  // reference entry ("Dept of Home Affairs" vs "Department of Home Affairs").
  var STOPWORDS = {
    of: 1,
    the: 1,
    and: 1,
    for: 1,
    in: 1,
    on: 1,
    at: 1,
    a: 1,
    an: 1,
    to: 1,
    et: 1,
    al: 1,
  };

  // Lowercase words allowed *inside* a multi-word author name when scanning
  // backwards for a narrative citation ("Department of Home Affairs (2026)").
  var NAME_CONNECTORS = {
    of: 1,
    and: 1,
    the: 1,
    for: 1,
    de: 1,
    del: 1,
    della: 1,
    van: 1,
    von: 1,
    der: 1,
    den: 1,
    la: 1,
    le: 1,
    da: 1,
    di: 1,
    dos: 1,
    et: 1,
    al: 1,
  };

  // Introductory words that appear before the author inside a parenthetical:
  // "(see Smith 2020)", "(e.g. ASD 2025)".
  var SIGNAL_PREFIX_RE =
    /^(?:see\s+also|see|e\.?\s?g\.?|i\.?\s?e\.?|cf\.?|compare|also|as\s+cited\s+in|cited\s+in|adapted\s+from|based\s+on|from|after|per|following)\b[\s,:.]*/i;

  // Publication years live in a much narrower band than "any four digits".
  // Anchoring to it stops case numbers ("FCA 1224") and street numbers from
  // being read as years.
  var MIN_YEAR = 1450;
  var MAX_YEAR = 2199;

  // ---------------------------------------------------------------------
  // Locators
  // ---------------------------------------------------------------------

  // The words that introduce a pinpoint reference inside a citation:
  // "(Smith 2020, p. 44)", "(Smith 2020, ch. 3)", "(Smith 2020, paras. 7-9)".
  //
  // These matter because a bracket is replaced whole. Everything the bracket
  // held that is not the author and the year has to be accounted for, or the
  // replacement deletes it — and a page number is the one part of a citation a
  // reader needs in order to find the passage being quoted.
  var LOCATOR_WORDS =
    "pp?|pgs?|pages?|ch|chap|chapters?|ss?|secs?|sections?|paras?|paragraphs?|" +
    "vols?|volumes?|nos?|arts?|articles?|figs?|figures?|tbls?|tables?|" +
    "nn?|ll?|lines?|apps?|appendix|appendices|slides?|verses?|vv?|" +
    "clauses?|cll?|schs?|schedules?|regs?|regulations?|items?|" +
    "sub-?paras?|sub-?sections?";

  // A pinpoint's number: digits, or digits with a letter ("44a"). Roman
  // numerals are allowed too, but only after one of the words above — the
  // letters i, v, x, l, c, d and m spell ordinary English words ("civil",
  // "mild", "mix"), and a bare roman numeral as a page reference is rare enough
  // that accepting one would misread more prose than it would read citations.
  function locList(num) {
    var range = "(?:" + num + ")(?:\\s*[-–—]\\s*(?:" + num + "))?";
    return range + "(?:\\s*(?:,|&|and)\\s*" + range + ")*";
  }
  var LOC_DIGITS = "[0-9]+[a-z]?";
  var LOC_ANY = "[0-9]+[a-z]?|[ivxlcdm]+";
  // The word and its number are always separated by a full stop, a space, or
  // both. Without that requirement "v" + "ivid" reads "vivid" as verse ivid,
  // and the same trick turns a dozen other words into pinpoints.
  var LOC_ONE =
    "(?:(?:" + LOCATOR_WORDS + ")(?:\\.\\s*|\\s+)" + locList(LOC_ANY) +
    "|" + locList(LOC_DIGITS) + ")";
  var LOCATOR_RE = new RegExp(
    "^" + LOC_ONE + "(?:\\s*,\\s*" + LOC_ONE + ")*$",
    "i",
  );
  var LOCATOR_WORD_RE = new RegExp("(?:^|[\\s,])(?:" + LOCATOR_WORDS + ")\\b", "i");

  /**
   * Split whatever followed the year inside a citation into a pinpoint locator
   * and an unaccounted-for remainder.
   *
   * `locator` is text that will be carried into the footnote. `residue` is text
   * this file cannot explain, and its presence means the bracket must be left
   * exactly as the author wrote it — "(Smith 2020, as cited in Jones 2021)" and
   * "(Smith 2020, Jones 2021)" both land there, and converting either one would
   * throw away a second source.
   *
   * A bare number is accepted as a page ("(Smith 2020: 44)", "(Smith 2020, 44)")
   * but only when it is not year-shaped, because "(Smith 2020, 2021)" is two
   * works by Smith and not page 2021 of one.
   */
  function splitLocator(tail) {
    var t = String(tail || "")
      .replace(/^[\s.,:;()[\]]+/, "")
      .replace(/[\s.,:;()[\]]+$/, "");
    if (!t) return { locator: "", residue: "" };
    if (!LOCATOR_RE.test(t)) return { locator: "", residue: t };
    // No locator word at all means it is a bare number, which must not be a
    // year: that is a second citation, not a pinpoint.
    if (!LOCATOR_WORD_RE.test(t) && findYearTokens(t).length) {
      return { locator: "", residue: t };
    }
    return { locator: t, residue: "" };
  }

  // ---------------------------------------------------------------------
  // Small string helpers
  // ---------------------------------------------------------------------

  function normHeading(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z ]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function isReferencesHeadingText(s) {
    return HEADING_WORDS.indexOf(normHeading(s)) !== -1;
  }

  function isSectionEndHeadingText(s) {
    var h = normHeading(s);
    if (SECTION_END_WORDS.indexOf(h) !== -1) return true;
    // "Appendix A", "Appendix 1" both normalise with the letter still attached.
    return /^appendix\b/.test(h) || /^annex\b/.test(h);
  }

  /**
   * Reduce an author name to comparable lowercase tokens.
   * "OAIC (Office of the Australian Information Commissioner)" is handled by
   * the caller, which splits the parenthetical out first; this function only
   * sees one name at a time.
   *
   * Single characters are dropped because they are almost always initials
   * ("Smith, J., & Jones, K." -> ["smith", "jones"]).
   */
  function normalizeTokens(s) {
    return String(s || "")
      .replace(/&/g, " and ")
      .replace(/[’‘]/g, "'")
      .toLowerCase()
      .replace(/[^a-z0-9\s'-]/g, " ")
      .split(/\s+/)
      .map(function (t) {
        return t.replace(/^[-']+|[-']+$/g, "");
      })
      .filter(function (t) {
        return t.length > 1 && !STOPWORDS[t];
      });
  }

  function tokensEqual(a, b) {
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }

  function tokensArePrefix(short, long) {
    if (!short.length || short.length >= long.length) return false;
    for (var i = 0; i < short.length; i++) if (short[i] !== long[i]) return false;
    return true;
  }

  function tokensSubset(sub, sup) {
    if (!sub.length) return false;
    for (var i = 0; i < sub.length; i++)
      if (sup.indexOf(sub[i]) === -1) return false;
    return true;
  }

  // ---------------------------------------------------------------------
  // Years
  // ---------------------------------------------------------------------

  /**
   * Find every publication-year-shaped token in a string, with its optional
   * disambiguation suffix ("2025a").
   *
   * Deliberately rejects:
   *   - digits inside a longer run          20255, 1,200
   *   - digits glued to letters             C2024A00098, 2024AD
   *   - either half of a numeric range      2024-25, 2023-2030
   *   - anything outside MIN_YEAR..MAX_YEAR FCA 1224
   *
   * That last group matters more than it looks: reference entries routinely
   * carry URLs full of four-digit slugs, and a naive /\d{4}/ reads
   * ".../report-2024-2025" as the publication year.
   */
  function findYearTokens(text) {
    var out = [];
    var s = String(text || "");
    var re = /(\d{4})([a-z])?/g;
    var m;
    while ((m = re.exec(s))) {
      var start = m.index;
      var end = start + m[0].length;
      var before = s.charAt(start - 1);
      var before2 = s.charAt(start - 2);
      var after = s.charAt(end);
      var after2 = s.charAt(end + 1);

      if (/[0-9A-Za-z]/.test(before)) continue;
      if (/[0-9A-Za-z]/.test(after)) continue;
      // A hyphen joining the digits to a word or another number means this is
      // part of a compound, not a year: a URL slug ("...-report-2024"), an
      // identifier ("id-2024x"), or half of a range ("2024-25").
      if (before === "-" && /[0-9A-Za-z]/.test(before2)) continue;
      if (after === "-" && /[0-9]/.test(after2)) continue;

      var y = parseInt(m[1], 10);
      if (y < MIN_YEAR || y > MAX_YEAR) continue;

      out.push({
        index: start,
        end: end,
        year: m[1],
        suffix: m[2] || "",
      });
    }
    return out;
  }

  // "n.d." / "no date" — the standard stand-in for a missing year. Without
  // recognising it, such references have no year at all and can never match.
  var ND_RE = /\bn\.?\s*d\.?(?![a-z])|\bno\s+date\b/i;

  function findNoDate(text) {
    var m = String(text || "").match(ND_RE);
    return m ? { index: m.index, end: m.index + m[0].length, year: "n.d.", suffix: "" } : null;
  }

  // ---------------------------------------------------------------------
  // Reference entries
  // ---------------------------------------------------------------------

  /**
   * Split an author segment into comparable key token-lists.
   *
   * Australian government/agency style writes the acronym first and expands it
   * in brackets:
   *
   *   ASD (Australian Signals Directorate) 2025, Annual Cyber Threat Report...
   *
   * but cites it by the acronym alone: "(ASD 2025)". Treating the whole thing
   * as one flat bag of capitalised words means the citation matches none of
   * them. So the outside-the-brackets name and each bracketed expansion each
   * become a key in their own right, and a citation matching *any* key wins.
   */
  function buildAuthorKeys(authorSegment) {
    var seg = String(authorSegment || "");
    var inside = [];
    var outside = seg.replace(/\(([^)]*)\)/g, function (_all, g) {
      inside.push(g);
      return " ";
    });

    var keys = [];
    var primary = normalizeTokens(outside);
    if (primary.length) keys.push(primary);
    inside.forEach(function (s) {
      var t = normalizeTokens(s);
      if (t.length) keys.push(t);
    });

    // The acronym an organisation is cited by is often nowhere in the entry.
    // Where the bracketed style spells the expansion out —
    //   ASD (Australian Signals Directorate) 2025
    // — the plain style does not:
    //   Office of the Australian Information Commissioner 2025a, …, OAIC, Sydney
    // and the acronym that appears there sits in the *publisher* field, after
    // the year, outside the author segment entirely. Comparing the literal
    // words means "(OAIC 2025a)" matches nothing at all, and the entry is
    // reported as never cited while the citation is reported as unmatched —
    // one broken reference showing up twice as two separate complaints.
    //
    // Deriving the acronym from the name closes that gap. It falls out of the
    // token list for free, because the words a writer drops when forming an
    // acronym (of, the, and) are the same stopwords normalizeTokens already
    // removes: [office, australian, information, commissioner] -> "oaic".
    var acronyms = keys
      .map(acronymOf)
      .filter(Boolean)
      .map(function (a) {
        return [a];
      });
    acronyms.forEach(function (a) {
      var dup = keys.some(function (k) {
        return tokensEqual(k, a);
      });
      if (!dup) keys.push(a);
    });

    return keys;
  }

  /**
   * The acronym a multi-word name would be cited by, or null.
   *
   * Two words minimum, because a single-token name has no acronym, and its
   * own first letter is not one. Purely-numeric tokens are skipped so a name
   * carrying a number doesn't contribute a digit to the initials.
   */
  function acronymOf(tokens) {
    if (!tokens || tokens.length < 2) return null;
    var letters = "";
    for (var i = 0; i < tokens.length; i++) {
      var c = tokens[i].charAt(0);
      if (!/[a-z]/.test(c)) return null;
      letters += c;
    }
    return letters.length >= 2 ? letters : null;
  }

  /**
   * Capitalised words in the author segment, lowercased. Used as a last-resort
   * match for person-author styles where the citation names a subset of the
   * authors ("(Jones, 2021)" against "Smith, A., Jones, B. & Wu, C. (2021)").
   */
  function buildSurnames(authorSegment) {
    var words = String(authorSegment || "").match(/\b[A-Z][A-Za-z'’-]+\b/g) || [];
    var seen = {};
    var out = [];
    words.forEach(function (w) {
      var t = w.toLowerCase().replace(/[’]/g, "'");
      if (t.length < 2 || STOPWORDS[t] || seen[t]) return;
      seen[t] = 1;
      out.push(t);
    });
    return out;
  }

  /**
   * Locate the publication year in a reference entry and split the entry into
   * "who" (everything before the year) and "when".
   *
   * Order of preference:
   *   1. A parenthesised year, which is APA's unambiguous marker:
   *      Smith, J. (2020a). Title.
   *   2. A parenthesised "n.d.".
   *   3. The first plausible bare year token, which is Harvard/AGPS style:
   *      ASD (Australian Signals Directorate) 2025, Annual Cyber Threat Report
   *   4. A bare "n.d." anywhere.
   */
  function locateReferenceYear(text) {
    var s = String(text || "");

    var paren = s.match(/\(\s*(\d{4})([a-z])?\s*[.,)\]]/);
    if (paren) {
      var py = parseInt(paren[1], 10);
      if (py >= MIN_YEAR && py <= MAX_YEAR) {
        return { index: paren.index, year: paren[1], suffix: paren[2] || "" };
      }
    }

    var parenNd = s.match(/\(\s*n\.?\s*d\.?\s*[.,)\]]/i);
    if (parenNd) return { index: parenNd.index, year: "n.d.", suffix: "" };

    var bare = findYearTokens(s);
    if (bare.length) {
      return { index: bare[0].index, year: bare[0].year, suffix: bare[0].suffix };
    }

    var nd = findNoDate(s);
    if (nd) return { index: nd.index, year: "n.d.", suffix: "" };

    return null;
  }

  /**
   * Parse one reference-list entry into everything needed to match citations
   * against it. `url` is filled in here from the entry's visible text; the
   * OOXML layer overrides it when the entry carries a real hyperlink
   * relationship, which is more reliable.
   */
  function parseReferenceEntry(text) {
    var s = String(text || "");
    var located = locateReferenceYear(s);
    var year = located ? located.year : null;
    var suffix = located ? located.suffix : "";
    var authorSegment = located ? s.slice(0, located.index) : s;

    return {
      text: s,
      year: year,
      yearSuffix: suffix,
      authorSegment: authorSegment,
      authorKeys: buildAuthorKeys(authorSegment),
      surnames: buildSurnames(authorSegment),
      url: extractUrl(s),
    };
  }

  // ---------------------------------------------------------------------
  // URLs
  // ---------------------------------------------------------------------

  /**
   * Pull a source URL out of free text.
   *
   * The awkward cases this exists for:
   *   - AGPS/Harvard wraps URLs in angle brackets and puts a full stop after
   *     the closing one:  viewed 11 August 2026, <https://example.gov.au/a>.
   *   - Word autoformats bare "www." addresses without a scheme.
   *   - Trailing sentence punctuation is not part of the URL.
   *   - A URL may legitimately contain balanced parentheses
   *     (/wiki/Mercury_(planet)) while also being wrapped in one.
   */
  function extractUrl(text) {
    var s = String(text || "");
    var m = s.match(/\b(?:https?:\/\/|www\.)[^\s<>\[\]{}"'`]+/i);
    if (m) return cleanUrl(m[0]);

    var doi = s.match(/\bdoi:\s*(10\.\d{4,9}\/[^\s<>\[\]{}"'`]+)/i);
    if (doi) return "https://doi.org/" + trimTrailingPunctuation(doi[1]);

    return null;
  }

  function trimTrailingPunctuation(u) {
    return u.replace(/[.,;:!?'"»›]+$/, "");
  }

  function cleanUrl(raw) {
    var u = trimTrailingPunctuation(raw);
    while (u.charAt(u.length - 1) === ")" && countChar(u, ")") > countChar(u, "(")) {
      u = u.slice(0, -1);
      u = trimTrailingPunctuation(u);
    }
    if (/^www\./i.test(u)) u = "https://" + u;
    return u || null;
  }

  function countChar(s, c) {
    var n = 0;
    for (var i = 0; i < s.length; i++) if (s.charAt(i) === c) n++;
    return n;
  }

  // ---------------------------------------------------------------------
  // In-text citations
  // ---------------------------------------------------------------------

  /**
   * Parse one semicolon-delimited slice of a parenthetical into a citation
   * unit, or return null if it isn't citation-shaped.
   *
   * "(Cth)", "(No 2)", "(a)", "([Student Number 1])" all return null, which is
   * what keeps a legal document's statutory references from being converted
   * into footnotes.
   *
   * Whatever follows the year is split into `locator` and `residue`. A unit
   * carrying residue is one this file cannot fully explain, and the caller
   * refuses to convert the bracket it came from rather than deleting words
   * out of someone's document.
   */
  function parseCitationUnit(segment) {
    var seg = String(segment || "");
    if (!seg.trim()) return null;

    var years = findYearTokens(seg);
    var located = years.length ? years[0] : findNoDate(seg);
    if (!located) return null;

    var authorText = seg.slice(0, located.index).replace(SIGNAL_PREFIX_RE, "");
    var tail = splitLocator(seg.slice(located.end));
    return {
      raw: seg,
      year: located.year,
      yearSuffix: located.suffix || "",
      authorText: authorText,
      tokens: normalizeTokens(authorText),
      etAl: /\bet\s*al\b/i.test(seg),
      locator: tail.locator,
      residue: tail.residue,
    };
  }

  function splitSegments(inner, sep) {
    var out = [];
    var start = 0;
    for (var i = 0; i <= inner.length; i++) {
      if (i === inner.length || inner.charAt(i) === sep) {
        out.push({ text: inner.slice(start, i), start: start });
        start = i + 1;
      }
    }
    return out;
  }

  /**
   * How well does one citation unit match one reference? 0 means "not a
   * match at all"; higher is better. The scale exists so that a citation
   * naming an exact author key beats one that merely shares a surname, which
   * is what disambiguates same-year references from the same organisation.
   */
  function scoreUnitAgainstReference(unit, ref) {
    if (!unit || !ref || !ref.year || !unit.year) return 0;
    if (unit.year !== ref.year) return 0;

    // "(OAIC 2025b)" must not match the 2025a entry. When only one side
    // carries a suffix we still allow the match, but score it below an exact
    // suffix agreement so the better candidate wins.
    if (unit.yearSuffix && ref.yearSuffix && unit.yearSuffix !== ref.yearSuffix)
      return 0;
    var suffixPenalty = unit.yearSuffix === ref.yearSuffix ? 0 : 6;

    // The reverse of the acronym key built in buildAuthorKeys: the entry is
    // filed under the acronym ("OAIC 2025a, …") and the citation spells the
    // name out. Scored just below a literal key match so a reference that
    // actually carries those words still wins.
    var unitAcronym = acronymOf(unit.tokens);

    var best = 0;
    for (var i = 0; i < ref.authorKeys.length; i++) {
      var key = ref.authorKeys[i];
      var isPrimary = i === 0;
      if (tokensEqual(unit.tokens, key)) {
        best = Math.max(best, isPrimary ? 100 : 92);
      } else if (unitAcronym && key.length === 1 && key[0] === unitAcronym) {
        best = Math.max(best, 90);
      } else if (tokensArePrefix(unit.tokens, key)) {
        // "(Smith et al., 2022)" against "Smith, Jones & Wu", and truncated
        // organisation names, both land here.
        best = Math.max(best, isPrimary ? 74 : 66);
      } else if (tokensSubset(unit.tokens, key)) {
        best = Math.max(best, 48);
      }
    }

    if (!best && unit.tokens.length && tokensSubset(unit.tokens, ref.surnames)) {
      best = 34;
    }

    // A citation with a year but no author at all ("(2025)") only matches a
    // reference that also has no author. Narrative citations are handled
    // separately, by looking at the words before the bracket.
    if (!best && !unit.tokens.length && !ref.authorKeys.length) best = 20;

    return best > 0 ? Math.max(1, best - suffixPenalty) : 0;
  }

  /**
   * Every suffix of a token list, longest first: [a,b,c] -> [[a,b,c],[b,c],[c]].
   *
   * Used for the words before a narrative citation's bracket. The backward scan
   * cannot tell where the author's name starts, because the word before it is
   * capitalised just as often as not — "As Smith (2020) argues" and "However,
   * Smith (2020) found" are ordinary prose, and so is any sentence that ends in
   * an abbreviation before the next one begins.
   *
   * Suffixes only, never prefixes. Dropping words from the *front* of the
   * candidate can only ever discard something that was not part of the name;
   * dropping them from the back would let "the Smith Review (2020)" match a
   * "Smith 2020" entry, which is a different source.
   */
  function tokenSuffixes(tokens) {
    var out = [];
    for (var i = 0; i < tokens.length; i++) out.push(tokens.slice(i));
    return out;
  }

  function bestReferenceForUnit(unit, refs) {
    // Candidate author token lists, most specific first. Ordinary citations
    // have exactly one; a narrative citation has one per suffix of the words
    // in front of its bracket.
    var candidates =
      unit && unit.tokenCandidates && unit.tokenCandidates.length
        ? unit.tokenCandidates
        : [unit ? unit.tokens : []];

    for (var c = 0; c < candidates.length; c++) {
      var probe = candidates[c] === (unit && unit.tokens)
        ? unit
        : { year: unit.year, yearSuffix: unit.yearSuffix, tokens: candidates[c] };
      var bestRef = null;
      var bestScore = 0;
      for (var i = 0; i < refs.length; i++) {
        var score = scoreUnitAgainstReference(probe, refs[i]);
        if (score > bestScore) {
          bestScore = score;
          bestRef = refs[i];
        }
      }
      // The first candidate that matches anything wins. Because the list runs
      // longest to shortest, a citation naming three authors is tried in full
      // before any shorter reading of it.
      if (bestScore > 0) {
        return { ref: bestRef, score: bestScore, tokens: candidates[c] };
      }
    }
    return null;
  }

  /**
   * Walk backwards from an opening bracket collecting the words that make up a
   * narrative citation's author: "Department of Home Affairs (2026a) reports…".
   * Stops at the first word that is neither capitalised nor a name connector.
   */
  function precedingAuthorTokens(text, parenStart) {
    var before = String(text).slice(0, parenStart).replace(/\s+$/, "");
    if (!before) return [];
    var words = before.split(/\s+/);
    var picked = [];
    for (var i = words.length - 1; i >= 0 && picked.length < 7; i--) {
      var w = words[i].replace(/[^A-Za-z&'’.-]/g, "");
      if (!w) break;
      if (/^[A-Z]/.test(w) || NAME_CONNECTORS[w.toLowerCase().replace(/\.$/, "")]) {
        picked.unshift(w);
      } else break;
    }
    while (picked.length && !/^[A-Z]/.test(picked[0])) picked.shift();
    return normalizeTokens(picked.join(" "));
  }

  /**
   * Find every convertible citation in a paragraph.
   *
   * Returns groups rather than individual spans: one group is one stretch of
   * text that will be replaced by one or more footnote markers. A single
   * parenthetical holding several citations — "(Smith 2020; Jones 2021)" —
   * becomes one group with two reference ids, in the order they read.
   *
   * A parenthetical is only converted when *every* non-empty piece of it is a
   * citation that matched a reference. Half-converting
   * "(2024-25 data; ASD 2025)" would silently delete the words "2024-25 data"
   * from the document, so instead the whole bracket is left alone and reported
   * as unmatched, where the person can see it.
   */
  function findCitationGroups(text, refs) {
    var s = String(text || "");
    var groups = [];
    var unmatched = [];
    if (!refs || !refs.length) return { groups: groups, unmatched: unmatched };

    var re = /\(([^()]{0,240})\)/g;
    var m;
    while ((m = re.exec(s))) {
      var inner = m[1];
      var segments = splitSegments(inner, ";");
      var nonEmpty = segments.filter(function (seg) {
        return seg.text.trim().length > 0;
      });
      if (!nonEmpty.length) continue;

      var units = [];
      var unparsed = 0;
      for (var i = 0; i < nonEmpty.length; i++) {
        var unit = parseCitationUnit(nonEmpty[i].text);
        if (unit) units.push(unit);
        else unparsed++;
      }
      // Nothing in the bracket is citation-shaped, so it is just a bracket:
      // "(Cth)", "(No 2)", "(a)". Skip it without comment — a legal document
      // is full of these and flagging each one would bury the real problems.
      if (!units.length) continue;

      // Narrative form: the bracket holds only a year, and the author sits
      // immediately before it. Only the bracket is replaced, so the author's
      // name stays in the sentence where it belongs.
      if (units.length === 1 && !units[0].tokens.length) {
        var inherited = precedingAuthorTokens(s, m.index);
        if (inherited.length) {
          units[0].tokens = inherited;
          units[0].tokenCandidates = tokenSuffixes(inherited);
        }
      }

      var refIds = [];
      var locators = [];
      var allMatched = unparsed === 0;
      for (var j = 0; j < units.length; j++) {
        // Words inside the bracket that are neither author, year nor pinpoint.
        // Whatever they are, replacing the bracket would erase them.
        if (units[j].residue) {
          allMatched = false;
          continue;
        }
        var hit = bestReferenceForUnit(units[j], refs);
        if (!hit) {
          allMatched = false;
          continue;
        }
        var locator = units[j].locator || "";
        // Deduplicated on the pair, not on the reference: one source cited at
        // two different pages in one bracket is two notes, because the pages
        // are the only thing telling them apart.
        var seen = false;
        for (var k = 0; k < refIds.length; k++) {
          if (refIds[k] === hit.ref.id && locators[k] === locator) {
            seen = true;
            break;
          }
        }
        if (!seen) {
          refIds.push(hit.ref.id);
          locators.push(locator);
        }
      }

      // Convert only when the whole bracket is accounted for. Replacing part
      // of "(2024-25 figures; Smith 2020)" would delete the words "2024-25
      // figures" from the document, so the bracket is left as written and
      // reported instead — visible to the person, who can fix it or ignore it.
      if (!allMatched || !refIds.length) {
        unmatched.push({ text: m[0], start: m.index, end: m.index + m[0].length });
        continue;
      }

      groups.push({
        start: trimLeadingWhitespace(s, m.index),
        end: m.index + m[0].length,
        refIds: refIds,
        // Parallel to refIds: the pinpoint each note carries, "" for most.
        locators: locators,
        sourceText: m[0],
      });
    }

    return { groups: groups, unmatched: unmatched };
  }

  /**
   * A footnote marker replaces the citation *and* the space in front of it, so
   * "…rose by 14 per cent (ASD 2025)." becomes "…rose by 14 per cent¹." rather
   * than leaving an orphaned space before the full stop. The space after the
   * bracket, if any, is left in place, which keeps mid-sentence citations
   * reading correctly: "As noted¹ the figures…".
   */
  function trimLeadingWhitespace(text, index) {
    var i = index;
    while (i > 0 && /\s/.test(text.charAt(i - 1))) i--;
    return i;
  }

  /**
   * Split a paragraph into alternating parts for the preview pane.
   *
   * `existingMarks` are footnote markers the document already contained, as
   * `[{ offset, footnoteId }]`. They're woven in at their character offsets so
   * the preview shows the document's real reading order: a paragraph that
   * already had a footnote and also has a citation to convert must show both,
   * in the right order, or the numbering the person sees won't match the
   * numbering Word will produce.
   *
   * Mirrors findCitationGroups exactly, so what is highlighted is precisely
   * what the download will convert.
   */
  function splitParagraphForPreview(text, refs, existingMarks) {
    var s = String(text || "");
    var found = findCitationGroups(s, refs);
    var parts = [];

    var events = [];
    found.groups.forEach(function (g) {
      events.push({ start: g.start, end: g.end, group: g });
    });
    (existingMarks || []).forEach(function (mark) {
      // A marker sitting inside a range that's about to be replaced would be
      // destroyed by the replacement, so it isn't shown as surviving.
      var swallowed = found.groups.some(function (g) {
        return mark.offset > g.start && mark.offset < g.end;
      });
      if (swallowed) return;
      events.push({ start: mark.offset, end: mark.offset, mark: mark });
    });

    // Where an existing marker sits exactly at a group's start, the marker
    // came first in the document and keeps its place.
    events.sort(function (a, b) {
      return a.start - b.start || (a.end - a.start) - (b.end - b.start);
    });

    var cursor = 0;
    events.forEach(function (ev) {
      if (ev.start > cursor) {
        parts.push({ type: "text", value: s.slice(cursor, ev.start) });
      }
      if (ev.mark) {
        parts.push({ type: "note", footnoteId: ev.mark.footnoteId });
      } else {
        ev.group.refIds.forEach(function (refId, idx) {
          parts.push({
            type: "cite",
            refId: refId,
            locator: (ev.group.locators || [])[idx] || "",
            // The original bracket is shown once, on the first marker of the
            // group, so a shared parenthetical isn't printed twice. The space
            // the marker swallows is not shown, because the conversion removes
            // it — the preview should look like the result, not the input.
            text: idx === 0 ? ev.group.sourceText : "",
          });
        });
      }
      cursor = Math.max(cursor, ev.end);
    });

    if (cursor < s.length) parts.push({ type: "text", value: s.slice(cursor) });
    if (!parts.length) parts.push({ type: "text", value: s });

    return { parts: parts, groups: found.groups, unmatched: found.unmatched };
  }

  return {
    HEADING_WORDS: HEADING_WORDS,
    SECTION_END_WORDS: SECTION_END_WORDS,
    MIN_YEAR: MIN_YEAR,
    MAX_YEAR: MAX_YEAR,

    normHeading: normHeading,
    isReferencesHeadingText: isReferencesHeadingText,
    isSectionEndHeadingText: isSectionEndHeadingText,

    normalizeTokens: normalizeTokens,
    tokensEqual: tokensEqual,
    tokensArePrefix: tokensArePrefix,
    tokensSubset: tokensSubset,
    tokenSuffixes: tokenSuffixes,

    splitLocator: splitLocator,

    findYearTokens: findYearTokens,
    findNoDate: findNoDate,
    locateReferenceYear: locateReferenceYear,

    acronymOf: acronymOf,
    buildAuthorKeys: buildAuthorKeys,
    buildSurnames: buildSurnames,
    parseReferenceEntry: parseReferenceEntry,

    extractUrl: extractUrl,
    cleanUrl: cleanUrl,

    parseCitationUnit: parseCitationUnit,
    splitSegments: splitSegments,
    scoreUnitAgainstReference: scoreUnitAgainstReference,
    bestReferenceForUnit: bestReferenceForUnit,
    precedingAuthorTokens: precedingAuthorTokens,
    findCitationGroups: findCitationGroups,
    trimLeadingWhitespace: trimLeadingWhitespace,
    splitParagraphForPreview: splitParagraphForPreview,
  };
});
