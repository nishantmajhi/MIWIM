"use strict";

/**
 * Narrative citations — "As Smith (2020) argues", where the author sits outside
 * the bracket and is doing grammatical work in the sentence.
 *
 * Recovering the author means reading backwards from the bracket, and the
 * backward scan cannot tell where the name starts: the word in front of it is
 * capitalised about as often as not. Before the fix the scan swallowed that word
 * and then compared ["as","smith"] against ["smith"], so eleven of eighteen
 * ordinary sentences produced no footnote at all and reported the reference as
 * never cited. The matcher now tries the suffixes of what the scan collected,
 * longest first.
 *
 * Confirmed failing against commit 25a70b9.
 */

const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("path");

const core = require(path.join(__dirname, "..", "miwim.core.js"));

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

// ===========================================================================
test("2. a narrative citation matches through a preceding capital", async (t) => {
  const list = refs([
    "Smith, J 2020, A Book, Press, Sydney.",
    "Department of Home Affairs 2026a, Annual report, DHA, Canberra.",
  ]);

  await t.test("the ordinary sentence openings all match", () => {
    // Each of these failed before the fix, because the backward scan for the
    // author swallowed the capitalised first word and then compared
    // ["as","smith"] against ["smith"].
    const sentences = [
      ["As Smith (2020) argues, the data is poor.", "c1"],
      ["While Smith (2020) argues this, others disagree.", "c1"],
      ["Although Smith (2020) argues this, others disagree.", "c1"],
      ["However, Smith (2020) found the opposite.", "c1"],
      ["Recently Smith (2020) revisited the question.", "c1"],
      ["Indeed Smith (2020) says so.", "c1"],
      ["Thus Smith (2020) concludes the matter.", "c1"],
      ["See Smith (2020) for detail.", "c1"],
      ["In Australia Department of Home Affairs (2026a) reports a rise.", "c2"],
      ["However, the Department of Home Affairs (2026a) reports a rise.", "c2"],
    ];
    sentences.forEach(function (pair) {
      const found = core.findCitationGroups(pair[0], list);
      assert.equal(found.groups.length, 1, pair[0]);
      assert.deepEqual(found.groups[0].refIds, [pair[1]], pair[0]);
    });
  });

  await t.test("a sentence ending in a capitalised word does not bleed in", () => {
    // "…published by the OAIC. Smith (2020) notes…" fed ["oaic","smith"] to
    // the matcher. There is no punctuation rule that can tell an abbreviation
    // from a sentence end, so the matcher tries the shorter names too.
    const found = core.findCitationGroups(
      "The figures were published by the OAIC. Smith (2020) notes the gap.",
      list,
    );
    assert.equal(found.groups.length, 1);
    assert.deepEqual(found.groups[0].refIds, ["c1"]);
  });

  await t.test("the ones that already worked still work", () => {
    [
      "Smith (2020) argues that the data is poor.",
      "Growth slowed in 2019. Smith (2020) notes the gap.",
      "The data was poor. Smith (2020) notes the gap.",
      "According to Smith (2020) the data is poor.",
      "One study, Smith (2020), found otherwise.",
      "Department of Home Affairs (2026a) reports a rise.",
      "The Department of Home Affairs (2026a) reports a rise.",
    ].forEach(function (s) {
      assert.equal(core.findCitationGroups(s, list).groups.length, 1, s);
    });
  });

  await t.test("it never matches by dropping the last word of a name", () => {
    // Suffixes only. "Smith Review (2020)" must not match a "Smith 2020"
    // entry by throwing away "Review" — that would attach the wrong source.
    const found = core.findCitationGroups("The Smith Review (2020) recommended it.", list);
    assert.equal(found.groups.length, 0);
    assert.equal(found.unmatched.length, 1);
  });

  await t.test("a lowercase preceding word still invents no author", () => {
    assert.equal(core.findCitationGroups("this happened again (2020).", list).groups.length, 0);
  });

  await t.test("tokenSuffixes runs longest first", () => {
    assert.deepEqual(core.tokenSuffixes(["as", "smith"]), [["as", "smith"], ["smith"]]);
    assert.deepEqual(core.tokenSuffixes(["smith"]), [["smith"]]);
    assert.deepEqual(core.tokenSuffixes([]), []);
  });
});

