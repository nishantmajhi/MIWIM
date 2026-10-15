/*!
 * MIWIM — OOXML layer
 * ------------------------
 * Reads a .docx, finds its reference list, matches in-text citations against
 * it (using miwim.core.js), and writes a new .docx in which every matched
 * citation has become a real Word footnote — the same thing Word's own
 * "Insert Footnote" command produces, so it renumbers, prints and exports
 * normally.
 *
 * Every dependency is injected rather than reached for, so the same code runs
 * in the browser (native DOMParser + JSZip from a script tag) and under Node
 * (@xmldom/xmldom + the jszip package) without branching.
 *
 *   var engine = createDocxEngine({ DOMParser, XMLSerializer, JSZip, core });
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.MIWIM = root.MIWIM || {};
    root.MIWIM.createDocxEngine = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  var R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  var REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  var CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

  var FOOTNOTES_REL_TYPE = R_NS + "/footnotes";
  var HYPERLINK_REL_TYPE = R_NS + "/hyperlink";
  var FOOTNOTES_CONTENT_TYPE =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";

  var PLACEHOLDER_TEXT = "[Insert citations and references using the footnotes]";

  return function createDocxEngine(deps) {
    if (!deps || !deps.DOMParser || !deps.XMLSerializer || !deps.JSZip || !deps.core) {
      throw new Error(
        "createDocxEngine needs { DOMParser, XMLSerializer, JSZip, core }",
      );
    }
    var DOMParserImpl = deps.DOMParser;
    var XMLSerializerImpl = deps.XMLSerializer;
    var JSZip = deps.JSZip;
    var core = deps.core;

    // -----------------------------------------------------------------
    // Namespace-tolerant attribute access
    // -----------------------------------------------------------------
    // Some DOM implementations resolve w:id through getAttributeNS, others
    // only expose it under the literal prefixed name. Ask for both.

    function attr(el, ns, prefixed) {
      if (!el) return null;
      var local = prefixed.indexOf(":") >= 0 ? prefixed.split(":")[1] : prefixed;
      var v = null;
      if (el.getAttributeNS) {
        try {
          v = el.getAttributeNS(ns, local);
        } catch (e) {
          v = null;
        }
      }
      if (v === null || v === "") v = el.getAttribute(prefixed);
      return v === "" ? null : v;
    }

    // Browsers are stricter than @xmldom/xmldom about prefixed attribute names
    // on XML documents, so namespaced attributes go through setAttributeNS
    // where one exists, with a literal fallback for implementations that don't.
    function setNsAttr(el, ns, qname, value) {
      if (el.setAttributeNS) {
        try {
          el.setAttributeNS(ns, qname, value);
          return;
        } catch (e) {
          /* fall through */
        }
      }
      el.setAttribute(qname, value);
    }

    function elements(node, ns, local) {
      if (!node || !node.getElementsByTagNameNS) return [];
      return Array.prototype.slice.call(node.getElementsByTagNameNS(ns, local));
    }

    function parseXml(xml, what) {
      var doc = new DOMParserImpl().parseFromString(xml, "application/xml");
      if (doc.getElementsByTagName("parsererror").length) {
        throw new Error("Couldn't read " + (what || "this document") + "'s XML.");
      }
      return doc;
    }

    function serialize(doc) {
      return new XMLSerializerImpl().serializeToString(doc);
    }

    function hasAncestor(node, localName) {
      var p = node.parentNode;
      while (p && p.nodeType === 1) {
        if (p.localName === localName) return true;
        p = p.parentNode;
      }
      return false;
    }

    // -----------------------------------------------------------------
    // Reading paragraphs
    // -----------------------------------------------------------------

    function getBody(doc) {
      var b = elements(doc, W_NS, "body")[0];
      if (!b) throw new Error("This document has no body — it may be corrupt.");
      return b;
    }

    function paragraphStyleId(pEl) {
      var pPr = firstChildNamed(pEl, "pPr");
      if (!pPr) return null;
      var pStyle = firstChildNamed(pPr, "pStyle");
      return pStyle ? attr(pStyle, W_NS, "w:val") : null;
    }

    function firstChildNamed(el, localName) {
      if (!el) return null;
      for (var i = 0; i < el.childNodes.length; i++) {
        var n = el.childNodes[i];
        if (n.nodeType === 1 && n.localName === localName) return n;
      }
      return null;
    }

    function headingLevel(pEl) {
      var style = (paragraphStyleId(pEl) || "").toLowerCase();
      if (style === "title") return 0;
      var m = style.match(/^heading(\d)$/);
      return m ? parseInt(m[1], 10) : null;
    }

    /**
     * A generated table of contents duplicates every heading in the document,
     * including "Bibliography". Left in the paragraph list it produces a
     * phantom reference section pointing at a page number. Word marks these
     * with TOC1..TOC9 styles and/or a TOC/PAGEREF field, so both are checked.
     */
    function isTocParagraph(pEl) {
      var style = (paragraphStyleId(pEl) || "").toLowerCase();
      if (/^toc\d*$/.test(style) || style === "tocheading") return true;

      var instr = elements(pEl, W_NS, "instrText");
      for (var i = 0; i < instr.length; i++) {
        if (/\bTOC\b|\bPAGEREF\b/.test(instr[i].textContent || "")) return true;
      }
      var simple = elements(pEl, W_NS, "fldSimple");
      for (var j = 0; j < simple.length; j++) {
        if (/\bTOC\b|\bPAGEREF\b/.test(attr(simple[j], W_NS, "w:instr") || ""))
          return true;
      }
      return false;
    }

    /**
     * Every paragraph in the body, in reading order — including paragraphs
     * inside tables, content controls and text boxes, which the naive
     * "direct children of w:body" walk misses entirely. Documents that lay
     * their body text out in a single-cell table are common enough that
     * missing them means finding no citations at all.
     */
    function listParagraphs(body) {
      return elements(body, W_NS, "p").filter(function (p) {
        if (hasAncestor(p, "Fallback")) return false; // mc:AlternateContent
        return !isTocParagraph(p);
      });
    }

    /**
     * Ordered list of the runs that carry a paragraph's visible text, with the
     * character offsets each run occupies in that text.
     *
     * paragraphText() is defined as the concatenation of exactly these runs, so
     * offsets computed against the text always land on a real run. Deriving the
     * text from one traversal and the edit map from another (the usual mistake:
     * all w:t descendants for reading, direct w:r children for writing) makes
     * the two disagree the moment a paragraph contains a hyperlink, and the
     * replacement then chops the sentence in the wrong place.
     */
    function collectRunEntries(pEl) {
      var out = [];
      var offset = 0;
      elements(pEl, W_NS, "r").forEach(function (r) {
        // Runs inside w:pPr/w:rPr describe formatting, not content; runs
        // inside w:del are tracked deletions and carry w:delText, not w:t.
        if (hasAncestor(r, "pPr") || hasAncestor(r, "del")) return;
        var text = elements(r, W_NS, "t")
          .map(function (t) {
            return t.textContent || "";
          })
          .join("");
        out.push({ run: r, start: offset, end: offset + text.length, text: text });
        offset += text.length;
      });
      return out;
    }

    function paragraphText(pEl) {
      return collectRunEntries(pEl)
        .map(function (e) {
          return e.text;
        })
        .join("");
    }

    /**
     * Footnote markers the paragraph already contains, with the character
     * offset each sits at in paragraphText(). Walks runs in the same order and
     * with the same exclusions as collectRunEntries, so the offsets line up
     * with the text the matcher works on.
     */
    function collectFootnoteMarks(pEl) {
      var marks = [];
      var offset = 0;
      elements(pEl, W_NS, "r").forEach(function (r) {
        if (hasAncestor(r, "pPr") || hasAncestor(r, "del")) return;
        elements(r, W_NS, "footnoteReference").forEach(function (fr) {
          marks.push({ offset: offset, footnoteId: attr(fr, W_NS, "w:id") });
        });
        elements(r, W_NS, "t").forEach(function (t) {
          offset += (t.textContent || "").length;
        });
      });
      return marks;
    }

    function isSectionHeadingParagraph(pEl, text) {
      var style = (paragraphStyleId(pEl) || "").toLowerCase();
      if (/^heading\d?$/.test(style) || style === "title" || style === "subtitle")
        return true;
      return core.isSectionEndHeadingText(text);
    }

    /**
     * Is this paragraph a heading *nested inside* the reference section rather
     * than one that ends it?
     *
     * Reference lists are routinely subdivided:
     *
     *   Bibliography                 (Heading 1)
     *     Legislation and Cases      (Heading 2)   ← belongs to the section
     *       …entries…
     *     Secondary Sources          (Heading 2)   ← belongs to the section
     *       …entries…
     *
     * Treating any styled heading as the end of the list stops it dead at the
     * first subheading, so a document laid out this way reports "a references
     * heading was found, but nothing followed it" while its entries sit in
     * plain sight one paragraph below. A heading only ends the section when it
     * is at the same level as the reference heading or shallower — which is
     * what "the next section" actually means.
     *
     * Only applies when the reference heading is itself styled. If it was
     * hand-formatted there is no level to compare against, and the safe old
     * behaviour (any styled heading ends the list) is kept, since guessing
     * wrong there would swallow the whole rest of the document.
     */
    function isNestedSubheading(pEl, refLevel) {
      if (refLevel == null) return false;
      var level = headingLevel(pEl);
      return level != null && level > refLevel;
    }

    // -----------------------------------------------------------------
    // Locating the reference section
    // -----------------------------------------------------------------

    /**
     * Prefer a properly styled heading, and among equals prefer the last one.
     * A hand-typed table of contents that survived the TOC filter still leaves
     * a "Bibliography" line near the front of the document; the real one is
     * always further down.
     */
    function findReferencesSectionIndex(paragraphs) {
      var candidates = [];
      paragraphs.forEach(function (p, i) {
        var text = paragraphText(p);
        if (!core.isReferencesHeadingText(text)) return;
        var style = (paragraphStyleId(p) || "").toLowerCase();
        candidates.push({ index: i, styled: /^heading\d?$/.test(style) });
      });
      if (!candidates.length) return -1;
      var styled = candidates.filter(function (c) {
        return c.styled;
      });
      var pool = styled.length ? styled : candidates;
      return pool[pool.length - 1].index;
    }

    function findReferencesSectionEnd(paragraphs, refSectionIdx) {
      var refLevel = headingLevel(paragraphs[refSectionIdx]);
      for (var i = refSectionIdx + 1; i < paragraphs.length; i++) {
        var text = paragraphText(paragraphs[i]).trim();
        if (!text) continue;
        // A heading naming a section that follows reference lists ends it
        // whatever its level: "Appendix A" styled Heading 2 under a Heading 1
        // "Bibliography" is still the start of the appendix, not part of it.
        if (core.isSectionEndHeadingText(text)) return i;
        if (isNestedSubheading(paragraphs[i], refLevel)) continue;
        if (isSectionHeadingParagraph(paragraphs[i], text)) return i;
      }
      return paragraphs.length;
    }

    // -----------------------------------------------------------------
    // Relationships
    // -----------------------------------------------------------------

    function buildRelsMap(relsDoc) {
      var map = {};
      Array.prototype.slice
        .call(relsDoc.getElementsByTagName("Relationship"))
        .forEach(function (r) {
          map[r.getAttribute("Id")] = r.getAttribute("Target");
        });
      return map;
    }

    function getParagraphHyperlinkUrl(pEl, relsMap) {
      if (!relsMap) return null;
      var links = elements(pEl, W_NS, "hyperlink");
      for (var i = 0; i < links.length; i++) {
        var rid = attr(links[i], R_NS, "r:id");
        if (rid && relsMap[rid] && /^https?:/i.test(relsMap[rid])) return relsMap[rid];
      }
      return null;
    }

    function getMaxRelId(relsDoc) {
      var max = 0;
      Array.prototype.slice
        .call(relsDoc.getElementsByTagName("Relationship"))
        .forEach(function (r) {
          var m = (r.getAttribute("Id") || "").match(/^rId(\d+)$/);
          if (m) max = Math.max(max, parseInt(m[1], 10));
        });
      return max;
    }

    function addHyperlinkRelationship(relsDoc, url, nextIdRef) {
      // Word will not open a package containing two relationships with the
      // same id, so reuse an existing external relationship for the same URL
      // rather than adding a duplicate for every repeat citation.
      var existing = Array.prototype.slice.call(
        relsDoc.getElementsByTagName("Relationship"),
      );
      for (var i = 0; i < existing.length; i++) {
        if (
          existing[i].getAttribute("Type") === HYPERLINK_REL_TYPE &&
          existing[i].getAttribute("Target") === url
        ) {
          return existing[i].getAttribute("Id");
        }
      }
      var rel = relsDoc.createElementNS(REL_NS, "Relationship");
      var id = "rId" + nextIdRef.value++;
      rel.setAttribute("Id", id);
      rel.setAttribute("Type", HYPERLINK_REL_TYPE);
      rel.setAttribute("Target", url);
      rel.setAttribute("TargetMode", "External");
      relsDoc.documentElement.appendChild(rel);
      return id;
    }

    function ensureDocumentRelHasFootnotes(relsDoc, nextIdRef) {
      var has = Array.prototype.slice
        .call(relsDoc.getElementsByTagName("Relationship"))
        .some(function (r) {
          return /\/relationships\/footnotes$/.test(r.getAttribute("Type") || "");
        });
      if (has) return;
      var rel = relsDoc.createElementNS(REL_NS, "Relationship");
      rel.setAttribute("Id", "rId" + nextIdRef.value++);
      rel.setAttribute("Type", FOOTNOTES_REL_TYPE);
      rel.setAttribute("Target", "footnotes.xml");
      relsDoc.documentElement.appendChild(rel);
    }

    // -----------------------------------------------------------------
    // Reference parsing
    // -----------------------------------------------------------------

    /**
     * A reference-list paragraph that is really a template instruction rather
     * than a source: the placeholder MIWIM itself leaves behind, or the
     * "[Add your references here]" line a document template ships with. Both
     * are wrapped in square brackets and carry no publication details.
     *
     * Without this, converting a document twice reads the first pass's own
     * placeholder back as a reference entry.
     */
    function isPlaceholderEntry(text) {
      var t = String(text).trim();
      if (t === PLACEHOLDER_TEXT) return true;
      return /^\[[^\]]*\]$/.test(t) && !/https?:\/\//i.test(t);
    }

    function parseReferences(paragraphs, refSectionIdx, refSectionEndIdx, relsMap) {
      var refs = [];
      var endIdx = refSectionEndIdx == null ? paragraphs.length : refSectionEndIdx;
      var refLevel = headingLevel(paragraphs[refSectionIdx]);
      for (var i = refSectionIdx + 1; i < endIdx; i++) {
        var text = paragraphText(paragraphs[i]).trim();
        if (!text) continue;
        if (isPlaceholderEntry(text)) continue;
        // "Legislation and Cases", "Secondary Sources" — group labels that
        // organise the list. They are inside the section (that is why the
        // walk did not stop at them) but they are not sources, and counting
        // them as entries would report phantom uncited references and block
        // the list from collapsing.
        if (isNestedSubheading(paragraphs[i], refLevel)) continue;

        var parsed = core.parseReferenceEntry(text);
        // A real hyperlink relationship beats a URL scraped out of the visible
        // text: it is the only thing that works when the clickable text is a
        // title or "View source" rather than the address itself.
        var url = getParagraphHyperlinkUrl(paragraphs[i], relsMap) || parsed.url;

        refs.push({
          n: refs.length + 1,
          id: "c" + (refs.length + 1),
          pIndex: i,
          text: text,
          url: url,
          hasUrl: !!url,
          year: parsed.year,
          yearSuffix: parsed.yearSuffix,
          authorSegment: parsed.authorSegment,
          authorKeys: parsed.authorKeys,
          surnames: parsed.surnames,
          citationCount: 0,
        });
      }
      return refs;
    }

    // -----------------------------------------------------------------
    // Building runs and footnotes
    // -----------------------------------------------------------------

    function cloneRunProps(rEl) {
      var rPr = firstChildNamed(rEl, "rPr");
      return rPr ? rPr.cloneNode(true) : null;
    }

    function makeTextRun(doc, text, rPrTemplate) {
      var r = doc.createElementNS(W_NS, "w:r");
      if (rPrTemplate) r.appendChild(rPrTemplate.cloneNode(true));
      var t = doc.createElementNS(W_NS, "w:t");
      t.setAttribute("xml:space", "preserve");
      t.appendChild(doc.createTextNode(text));
      r.appendChild(t);
      return r;
    }

    /**
     * The run Word itself writes at a citation point: styled FootnoteReference
     * (superscript comes from the style, not a hand-set property) wrapping a
     * live w:footnoteReference. This is what makes the mark auto-numbered,
     * clickable and renumberable, rather than a superscript digit typed by
     * hand that breaks the moment anyone inserts a footnote above it.
     */
    function makeFootnoteReferenceRun(doc, footnoteId) {
      var r = doc.createElementNS(W_NS, "w:r");
      var rPr = doc.createElementNS(W_NS, "w:rPr");
      var rStyle = doc.createElementNS(W_NS, "w:rStyle");
      rStyle.setAttribute("w:val", "FootnoteReference");
      rPr.appendChild(rStyle);
      r.appendChild(rPr);
      var ref = doc.createElementNS(W_NS, "w:footnoteReference");
      ref.setAttribute("w:id", String(footnoteId));
      r.appendChild(ref);
      return r;
    }

    /**
     * Replace [start, end) of a paragraph's text with footnote markers.
     *
     * Runs are edited in place relative to their own parent, so a citation
     * that happens to sit inside a w:hyperlink or w:smartTag keeps the rest of
     * that wrapper intact. The marker itself is lifted out of a hyperlink
     * parent, because a footnote number rendered as part of a link is both
     * wrong to click and wrong to look at.
     */
    function replaceSpanWithFootnoteRefs(doc, pEl, start, end, footnoteIds) {
      var entries = collectRunEntries(pEl);
      var inserted = false;

      entries.forEach(function (e) {
        if (e.end <= start || e.start >= end) return;

        var parent = e.run.parentNode;
        if (!parent) return;

        var rPr = cloneRunProps(e.run);
        var before = e.text.slice(0, Math.max(0, start - e.start));
        var after = e.text.slice(
          Math.min(e.text.length, Math.max(0, end - e.start)),
        );

        var pending = [];
        if (before) pending.push(makeTextRun(doc, before, rPr));

        var markers = [];
        if (!inserted && e.start <= start && start <= e.end) {
          footnoteIds.forEach(function (id, idx) {
            // A plain comma keeps adjacent markers legible as two footnotes
            // instead of one run-together number.
            if (idx > 0) markers.push(makeTextRun(doc, ",", rPr));
            markers.push(makeFootnoteReferenceRun(doc, id));
          });
          inserted = true;
        }

        var markerTarget = parent;
        var markerAnchor = e.run;
        var tailLink = null;
        if (parent.localName === "hyperlink" && parent.parentNode) {
          // The marker is lifted out of the link, because a footnote number
          // rendered as part of a hyperlink is both wrong to click and wrong to
          // look at. Anything that followed the citation inside the link has to
          // be lifted out with it and re-wrapped, or it ends up in front of the
          // marker and the number lands at the end of the link instead of where
          // the citation was: "see the guide online¹" rather than "see the
          // guide¹ online". A shallow clone keeps r:id and w:history, so both
          // halves stay the same link.
          markerTarget = parent.parentNode;
          markerAnchor = parent.nextSibling;
          if (after && markers.length) {
            tailLink = parent.cloneNode(false);
            tailLink.appendChild(makeTextRun(doc, after, rPr));
            after = "";
          }
        } else {
          pending = pending.concat(markers);
          markers = [];
        }

        if (after) pending.push(makeTextRun(doc, after, rPr));

        pending.forEach(function (n) {
          parent.insertBefore(n, e.run);
        });
        parent.removeChild(e.run);

        markers.forEach(function (n) {
          markerTarget.insertBefore(n, markerAnchor);
        });
        if (tailLink) markerTarget.insertBefore(tailLink, markerAnchor);

        // A hyperlink emptied of all its runs is invalid; drop it.
        if (
          parent.localName === "hyperlink" &&
          !elements(parent, W_NS, "r").length &&
          parent.parentNode
        ) {
          parent.parentNode.removeChild(parent);
        }
      });

      if (!inserted) {
        footnoteIds.forEach(function (id, idx) {
          if (idx > 0) pEl.appendChild(makeTextRun(doc, ",", null));
          pEl.appendChild(makeFootnoteReferenceRun(doc, id));
        });
      }
    }

    // -----------------------------------------------------------------
    // The footnotes part
    // -----------------------------------------------------------------

    // A .docx saved by Word already ships an otherwise-empty word/footnotes.xml
    // holding just the separator entries, so usually this is an append. The
    // from-scratch branch only matters for files produced by other tools.
    function loadOrCreateFootnotesXml(zip) {
      var file = zip.file("word/footnotes.xml");
      if (file) {
        return file.async("string").then(function (xml) {
          return parseXml(xml, "the footnotes part");
        });
      }
      var xml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:footnotes xmlns:r="' +
        R_NS +
        '" xmlns:w="' +
        W_NS +
        '">' +
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
        "</w:footnotes>";
      return Promise.resolve(parseXml(xml, "the footnotes part"));
    }

    function loadOrCreateFootnotesRels(zip) {
      var file = zip.file("word/_rels/footnotes.xml.rels");
      if (file) {
        return file.async("string").then(function (xml) {
          return parseXml(xml, "the footnote relationships");
        });
      }
      var xml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' +
        REL_NS +
        '"></Relationships>';
      return Promise.resolve(parseXml(xml, "the footnote relationships"));
    }

    function getMaxFootnoteId(footnotesDoc) {
      var max = 0;
      elements(footnotesDoc, W_NS, "footnote").forEach(function (f) {
        var id = parseInt(attr(f, W_NS, "w:id"), 10);
        if (!isNaN(id)) max = Math.max(max, id);
      });
      return max;
    }

    /**
     * A footnote's body is the source's URL as a live link — the reference's
     * author, title and publication details stay where a reader expects them,
     * and the footnote is the thing you can actually click. When a reference
     * has no URL at all, the full reference text is written instead, because
     * an empty footnote is worse than a long one.
     *
     * `locator` is the pinpoint the citation carried — "p. 44", "ch. 3". It is
     * appended because the bracket it came from is about to be deleted, and it
     * is the one part of a citation a reader needs in order to find the passage.
     */
    function buildFootnoteContent(footnotesDoc, ref, footnotesRelsDoc, nextRelIdRef, locator) {
      var nodes = [];
      var pinpoint = String(locator || "").trim();

      if (!ref.url) {
        var body = ref.text || "";
        if (pinpoint) body = body ? body.replace(/\.?$/, ", " + pinpoint) : pinpoint;
        if (body) nodes.push(makeTextRun(footnotesDoc, body, null));
        return nodes;
      }

      var rid = addHyperlinkRelationship(footnotesRelsDoc, ref.url, nextRelIdRef);
      var hyperlink = footnotesDoc.createElementNS(W_NS, "w:hyperlink");
      setNsAttr(hyperlink, R_NS, "r:id", rid);
      hyperlink.setAttribute("w:history", "1");

      var r = footnotesDoc.createElementNS(W_NS, "w:r");
      var rPr = footnotesDoc.createElementNS(W_NS, "w:rPr");
      var rStyle = footnotesDoc.createElementNS(W_NS, "w:rStyle");
      rStyle.setAttribute("w:val", "Hyperlink");
      rPr.appendChild(rStyle);
      r.appendChild(rPr);
      var t = footnotesDoc.createElementNS(W_NS, "w:t");
      t.appendChild(footnotesDoc.createTextNode(ref.url));
      r.appendChild(t);
      hyperlink.appendChild(r);

      nodes.push(hyperlink);
      // Outside the hyperlink: the pinpoint is not part of the address, and a
      // page number rendered as part of a link is wrong to click.
      if (pinpoint) nodes.push(makeTextRun(footnotesDoc, ", " + pinpoint, null));
      return nodes;
    }

    function buildFootnoteElement(footnotesDoc, footnoteId, contentNodes) {
      var footnote = footnotesDoc.createElementNS(W_NS, "w:footnote");
      footnote.setAttribute("w:id", String(footnoteId));

      var p = footnotesDoc.createElementNS(W_NS, "w:p");
      var pPr = footnotesDoc.createElementNS(W_NS, "w:pPr");
      var pStyle = footnotesDoc.createElementNS(W_NS, "w:pStyle");
      pStyle.setAttribute("w:val", "FootnoteText");
      pPr.appendChild(pStyle);
      p.appendChild(pPr);

      var markRun = footnotesDoc.createElementNS(W_NS, "w:r");
      var markRPr = footnotesDoc.createElementNS(W_NS, "w:rPr");
      var markStyle = footnotesDoc.createElementNS(W_NS, "w:rStyle");
      markStyle.setAttribute("w:val", "FootnoteReference");
      markRPr.appendChild(markStyle);
      markRun.appendChild(markRPr);
      markRun.appendChild(footnotesDoc.createElementNS(W_NS, "w:footnoteRef"));
      p.appendChild(markRun);

      p.appendChild(makeTextRun(footnotesDoc, " ", null));
      contentNodes.forEach(function (n) {
        p.appendChild(n);
      });

      footnote.appendChild(p);
      addFootnoteSpacing(footnotesDoc, footnote);
      return footnote;
    }

    /**
     * Add a blank line after a footnote's content.
     *
     * Word packs footnotes flush against each other at the foot of the page.
     * With one short line each that is fine; with a long URL each, they run
     * together into an unreadable block. A trailing empty FootnoteText
     * paragraph gives each note its own line of air — the same thing a person
     * gets by pressing Enter at the end of a footnote, and what the
     * hand-prepared reference conversion did.
     *
     * Skipped when the footnote already ends in an empty paragraph, so running
     * a document through twice doesn't keep adding blank lines.
     */
    function addFootnoteSpacing(footnotesDoc, footnoteEl) {
      var paras = elements(footnoteEl, W_NS, "p");
      if (!paras.length) return false;
      var last = paras[paras.length - 1];
      if (!paragraphText(last).trim() && paras.length > 1) return false;

      var spacer = footnotesDoc.createElementNS(W_NS, "w:p");
      var pPr = footnotesDoc.createElementNS(W_NS, "w:pPr");
      var pStyle = footnotesDoc.createElementNS(W_NS, "w:pStyle");
      pStyle.setAttribute("w:val", "FootnoteText");
      pPr.appendChild(pStyle);
      spacer.appendChild(pPr);
      footnoteEl.appendChild(spacer);
      return true;
    }

    function isRealFootnote(footnoteEl) {
      // Every document carries two synthetic entries — the separator line and
      // its continuation — which are not notes anyone wrote.
      return !attr(footnoteEl, W_NS, "w:type");
    }

    /**
     * Read the footnotes a document already contains, resolving each one's
     * hyperlink through footnotes.xml.rels. A document converted previously,
     * or one where someone inserted notes by hand, arrives here with real
     * content that has to be shown rather than ignored.
     */
    function readExistingFootnotes(zip) {
      var fnFile = zip.file("word/footnotes.xml");
      if (!fnFile) return Promise.resolve([]);

      var relsFile = zip.file("word/_rels/footnotes.xml.rels");
      var relsPromise = relsFile
        ? relsFile.async("string").then(function (xml) {
            return buildRelsMap(parseXml(xml, "the footnote relationships"));
          })
        : Promise.resolve({});

      return Promise.all([fnFile.async("string"), relsPromise]).then(function (parts) {
        var doc = parseXml(parts[0], "the footnotes part");
        var relsMap = parts[1];

        return elements(doc, W_NS, "footnote")
          .filter(isRealFootnote)
          .map(function (fn) {
            var text = elements(fn, W_NS, "t")
              .map(function (t) {
                return t.textContent || "";
              })
              .join("")
              .trim();
            var links = elements(fn, W_NS, "hyperlink");
            var url = null;
            for (var i = 0; i < links.length && !url; i++) {
              var rid = attr(links[i], R_NS, "r:id");
              if (rid && relsMap[rid]) url = relsMap[rid];
            }
            return {
              footnoteId: attr(fn, W_NS, "w:id"),
              text: text,
              url: url || core.extractUrl(text),
            };
          });
      });
    }

    // -----------------------------------------------------------------
    // Styles and content types
    // -----------------------------------------------------------------

    /**
     * Word only writes the footnote style *definitions* into styles.xml once a
     * document actually contains a footnote. A document that has never had one
     * therefore references styles it doesn't define, and the footnote text
     * renders as ordinary body copy. These are the definitions Word writes,
     * added only when missing.
     */
    function ensureRequiredStyles(stylesDoc) {
      var existing = elements(stylesDoc, W_NS, "style").map(function (s) {
        return attr(s, W_NS, "w:styleId");
      });
      var root = stylesDoc.documentElement;

      function addStyle(id, xmlStr) {
        if (existing.indexOf(id) !== -1) return;
        var wrapped = parseXml(
          '<w:root xmlns:w="' + W_NS + '">' + xmlStr + "</w:root>",
          "a style definition",
        );
        root.appendChild(
          stylesDoc.importNode(wrapped.documentElement.firstChild, true),
        );
        existing.push(id);
      }

      addStyle(
        "FootnoteText",
        '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:link w:val="FootnoteTextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
      );
      addStyle(
        "FootnoteTextChar",
        '<w:style w:type="character" w:customStyle="1" w:styleId="FootnoteTextChar"><w:name w:val="Footnote Text Char"/><w:link w:val="FootnoteText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
      );
      addStyle(
        "FootnoteReference",
        '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>',
      );
      addStyle(
        "Hyperlink",
        '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>',
      );
    }

    function ensureContentTypeOverride(zip, partName, contentType) {
      var path = "[Content_Types].xml";
      var file = zip.file(path);
      if (!file) return Promise.resolve();
      return file.async("string").then(function (xml) {
        if (xml.indexOf('PartName="' + partName + '"') !== -1) return;
        var ctDoc = parseXml(xml, "the content types part");
        var ns = ctDoc.documentElement.namespaceURI || CT_NS;
        var override = ctDoc.createElementNS(ns, "Override");
        override.setAttribute("PartName", partName);
        override.setAttribute("ContentType", contentType);
        ctDoc.documentElement.appendChild(override);
        zip.file(path, serialize(ctDoc));
      });
    }

    // -----------------------------------------------------------------
    // Collapsing the reference list
    // -----------------------------------------------------------------

    function clearParagraphContent(pEl) {
      Array.prototype.slice
        .call(pEl.childNodes)
        .filter(function (n) {
          return (
            n.nodeType === 1 &&
            (n.localName === "r" ||
              n.localName === "hyperlink" ||
              n.localName === "smartTag" ||
              n.localName === "bookmarkStart" ||
              n.localName === "bookmarkEnd")
          );
        })
        .forEach(function (n) {
          pEl.removeChild(n);
        });

      // Strip list numbering so the placeholder doesn't inherit "1." from the
      // reference entry it replaced.
      var pPr = firstChildNamed(pEl, "pPr");
      var numPr = pPr && firstChildNamed(pPr, "numPr");
      if (numPr) pPr.removeChild(numPr);
    }

    /**
     * Once a reference has become a footnote, leaving it in the list as well
     * duplicates it. The section therefore collapses to a single placeholder
     * paragraph — but only when every entry was actually converted.
     *
     * References that were never cited are kept exactly where they are. They
     * have no footnote to live in, and silently deleting a source from a
     * student's bibliography because the matcher didn't find its citation is
     * the worst thing this tool could do.
     */
    function collapseReferenceSection(doc, paragraphs, refIdx, refEndIdx, convertedIndexes) {
      var refLevel = headingLevel(paragraphs[refIdx]);
      var sectionIdx = [];
      for (var i = refIdx + 1; i < refEndIdx; i++) sectionIdx.push(i);

      // Group labels inside the list ("Legislation and Cases") are structure,
      // not sources. They are counted separately from the entries so that a
      // list which is subdivided still collapses when all of its entries have
      // become footnotes.
      var subheadings = sectionIdx.filter(function (i) {
        return isNestedSubheading(paragraphs[i], refLevel);
      });
      var isSubheading = {};
      subheadings.forEach(function (i) {
        isSubheading[i] = true;
      });

      var nonEmpty = sectionIdx.filter(function (i) {
        if (isSubheading[i]) return false;
        return paragraphText(paragraphs[i]).trim().length > 0;
      });
      var converted = nonEmpty.filter(function (i) {
        return convertedIndexes.indexOf(i) !== -1;
      });
      if (!converted.length) return { collapsed: false, removed: 0, kept: nonEmpty.length };

      if (converted.length === nonEmpty.length) {
        var first = paragraphs[nonEmpty[0]];
        clearParagraphContent(first);
        first.appendChild(makeTextRun(doc, PLACEHOLDER_TEXT, null));
        sectionIdx.forEach(function (i) {
          var p = paragraphs[i];
          if (p === first || !p.parentNode) return;
          p.parentNode.removeChild(p);
        });
        return { collapsed: true, removed: nonEmpty.length, kept: 0 };
      }

      converted.forEach(function (i) {
        var p = paragraphs[i];
        if (p.parentNode) p.parentNode.removeChild(p);
      });

      // A group label whose every entry became a footnote would otherwise be
      // left heading an empty group.
      subheadings.forEach(function (h) {
        var survivors = 0;
        for (var i = h + 1; i < refEndIdx; i++) {
          if (isSubheading[i]) break;
          if (!paragraphText(paragraphs[i]).trim()) continue;
          if (convertedIndexes.indexOf(i) === -1) survivors++;
        }
        if (!survivors && paragraphs[h].parentNode) {
          paragraphs[h].parentNode.removeChild(paragraphs[h]);
        }
      });

      return {
        collapsed: false,
        removed: converted.length,
        kept: nonEmpty.length - converted.length,
      };
    }

    // -----------------------------------------------------------------
    // Shared analysis
    // -----------------------------------------------------------------

    function readDocument(zip) {
      var docFile = zip.file("word/document.xml");
      if (!docFile) throw new Error("This doesn't look like a valid .docx file.");

      return docFile.async("string").then(function (xml) {
        var doc = parseXml(xml, "this document");
        var body = getBody(doc);
        var paragraphs = listParagraphs(body);

        var refIdx = findReferencesSectionIndex(paragraphs);
        var hasRefsSection = refIdx !== -1;
        var refEndIdx = hasRefsSection
          ? findReferencesSectionEnd(paragraphs, refIdx)
          : paragraphs.length;

        var relsFile = zip.file("word/_rels/document.xml.rels");
        var relsPromise = relsFile
          ? relsFile.async("string").then(function (relsXml) {
              return buildRelsMap(parseXml(relsXml, "the document relationships"));
            })
          : Promise.resolve(null);

        return relsPromise.then(function (relsMap) {
          var refs = hasRefsSection
            ? parseReferences(paragraphs, refIdx, refEndIdx, relsMap)
            : [];
          return {
            doc: doc,
            body: body,
            paragraphs: paragraphs,
            refIdx: refIdx,
            refEndIdx: refEndIdx,
            hasRefsSection: hasRefsSection,
            relsMap: relsMap,
            refs: refs,
            bodyEndIdx: hasRefsSection ? refIdx : paragraphs.length,
          };
        });
      });
    }

    // -----------------------------------------------------------------
    // Public: analyse for the preview
    // -----------------------------------------------------------------

    /**
     * Read a .docx into everything the preview needs. Never throws for a
     * document that simply has no references — a missing reference list is a
     * result to report, not an error to fail on.
     *
     * The `notes` it returns are what a reader will actually see at the foot of
     * the page: every footnote the document already had, plus every one this
     * conversion would add, numbered together in document order. That joint
     * numbering is the point. Word numbers footnotes by position, so a
     * document that already has three footnotes before its first citation will
     * number that citation 4 — and a preview that numbered it 1 would be
     * showing something that never happens.
     */
    function analyzeDocx(arrayBuffer) {
      var zipRef;
      return JSZip.loadAsync(arrayBuffer)
        .then(function (zip) {
          zipRef = zip;
          return Promise.all([readDocument(zip), readExistingFootnotes(zip)]);
        })
        .then(function (parts) {
          var ctx = parts[0];
          var existingFootnotes = parts[1];

          var refById = {};
          ctx.refs.forEach(function (r) {
            refById[r.id] = r;
          });
          var existingById = {};
          existingFootnotes.forEach(function (fn) {
            existingById[fn.footnoteId] = fn;
          });

          var docTitle = "";
          var bodyParas = [];
          var unmatched = [];
          var notes = [];
          var citationCount = 0;

          for (var i = 0; i < ctx.bodyEndIdx; i++) {
            var text = paragraphText(ctx.paragraphs[i]);
            var marks = collectFootnoteMarks(ctx.paragraphs[i]).filter(function (m) {
              return existingById[m.footnoteId];
            });
            if (!text.trim() && !marks.length) continue;
            if (!docTitle && text.trim()) {
              // The first non-empty paragraph is the document's title and is
              // rendered separately rather than as a body paragraph.
              //
              // Unless it carries a footnote, or a citation this conversion
              // would convert. buildConvertedDocx has no notion of a title and
              // converts every paragraph, so treating body text as a title here
              // made the preview a different document from the download: its
              // first citation was missing, every footnote number after it was
              // one too low, and the reference it cited was reported as never
              // cited right before the download deleted it from the
              // bibliography. A document that opens straight into prose — a
              // chapter, an extract, a pasted draft — has no title line.
              if (!marks.length && !core.findCitationGroups(text, ctx.refs).groups.length) {
                docTitle = text.trim();
                continue;
              }
            }

            var split = core.splitParagraphForPreview(text, ctx.refs, marks);
            split.unmatched.forEach(function (u) {
              unmatched.push({ text: u.text, paragraphIndex: i });
            });

            // Number the markers as they are met, so the count follows reading
            // order across the whole document rather than restarting anywhere.
            split.parts.forEach(function (part) {
              if (part.type === "cite") {
                part.n = notes.length + 1;
                citationCount++;
                if (refById[part.refId]) refById[part.refId].citationCount++;
                notes.push({
                  n: part.n,
                  kind: "new",
                  refId: part.refId,
                  paragraphIndex: i,
                });
              } else if (part.type === "note") {
                part.n = notes.length + 1;
                var existing = existingById[part.footnoteId];
                notes.push({
                  n: part.n,
                  kind: "existing",
                  footnoteId: part.footnoteId,
                  text: existing ? existing.text : "",
                  url: existing ? existing.url : null,
                  paragraphIndex: i,
                });
              }
            });

            bodyParas.push({
              id: "p" + i,
              heading: headingLevel(ctx.paragraphs[i]),
              parts: split.parts,
            });
          }

          // Footnotes defined in the part but never referenced in the body —
          // orphans, usually left behind by an edit. Listed so they aren't
          // invisible, but they sit after the numbered ones.
          var referenced = {};
          notes.forEach(function (note) {
            if (note.kind === "existing") referenced[note.footnoteId] = true;
          });
          var orphanFootnotes = existingFootnotes.filter(function (fn) {
            return !referenced[fn.footnoteId];
          });

          return {
            docTitle: docTitle || "Untitled document",
            bodyParas: bodyParas,
            refs: ctx.refs,
            hasRefsSection: ctx.hasRefsSection,
            unmatched: unmatched,
            citationCount: citationCount,
            uncitedRefs: ctx.refs.filter(function (r) {
              return r.citationCount === 0;
            }),
            notes: notes,
            existingFootnotes: existingFootnotes,
            orphanFootnotes: orphanFootnotes,
            existingFootnoteCount: existingFootnotes.length,
            // Nothing to convert: the document is already in footnote form, so
            // it is shown as it is rather than presented as a conversion.
            alreadyConverted:
              citationCount === 0 && existingFootnotes.length > 0,
          };
        });
    }

    // -----------------------------------------------------------------
    // Public: build the converted .docx
    // -----------------------------------------------------------------

    /**
     * Rewrite the document with real footnotes.
     *
     * `refs` comes back from analyzeDocx, possibly with URLs the person filled
     * in by hand, so the two passes agree on numbering and matching. Returns
     * { zip, summary }; the caller decides what to generate from the zip.
     */
    function buildConvertedDocx(arrayBuffer, refs) {
      var zipRef;
      return JSZip.loadAsync(arrayBuffer)
        .then(function (zip) {
          zipRef = zip;
          return readDocument(zip);
        })
        .then(function (ctx) {
          var zip = zipRef;
          var doc = ctx.doc;
          var paragraphs = ctx.paragraphs;

          // The document is re-parsed here, so paragraph indexes and author
          // keys must come from *this* pass. Anything the person edited in the
          // preview — realistically just a URL they pasted in for a reference
          // that had none — is overlaid onto it, matched by id rather than by
          // array position so a filtered or reordered list still lines up.
          var supplied = {};
          (refs || []).forEach(function (r) {
            if (r && r.id) supplied[r.id] = r;
          });
          var effectiveRefs = ctx.refs.map(function (fresh) {
            var override = supplied[fresh.id];
            if (!override) return fresh;
            return Object.keys(fresh).reduce(
              function (out, key) {
                out[key] = key in override && override[key] != null
                  ? override[key]
                  : fresh[key];
                return out;
              },
              { pIndex: fresh.pIndex, authorKeys: fresh.authorKeys },
            );
          });
          var refById = {};
          effectiveRefs.forEach(function (r) {
            refById[r.id] = r;
          });

          var relsPath = "word/_rels/document.xml.rels";
          var relsFile = zip.file(relsPath);
          var relsXmlPromise = relsFile
            ? relsFile.async("string")
            : Promise.resolve(
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' +
                  REL_NS +
                  '"></Relationships>',
              );

          return Promise.all([
            relsXmlPromise,
            loadOrCreateFootnotesXml(zip),
            loadOrCreateFootnotesRels(zip),
          ]).then(function (parts) {
            var relsDoc = parseXml(parts[0], "the document relationships");
            var footnotesDoc = parts[1];
            var footnotesRelsDoc = parts[2];
            var footnotesRoot = footnotesDoc.documentElement;

            var nextRelIdRef = { value: getMaxRelId(relsDoc) + 1 };
            var footnotesNextRelIdRef = { value: getMaxRelId(footnotesRelsDoc) + 1 };
            var nextFootnoteId = getMaxFootnoteId(footnotesDoc) + 1;

            var convertedIndexes = [];
            var footnotesCreated = 0;
            var newFootnoteEls = [];
            var unmatched = [];

            for (var i = 0; i < ctx.bodyEndIdx; i++) {
              var p = paragraphs[i];
              // Read the paragraph once and compute every group against that
              // single snapshot. Editing between matches would shift the
              // offsets the later groups were measured against.
              var text = paragraphText(p);
              if (!text.trim()) continue;

              var found = core.findCitationGroups(text, effectiveRefs);
              found.unmatched.forEach(function (u) {
                unmatched.push({ text: u.text, paragraphIndex: i });
              });
              if (!found.groups.length) continue;

              // Create the footnote entries in reading order first — that is
              // what fixes the visible numbering — then apply the body edits
              // right-to-left so earlier replacements can't move later spans.
              found.groups.forEach(function (g) {
                g.footnoteIds = g.refIds.map(function (refId, idx) {
                  var ref = refById[refId];
                  var content = buildFootnoteContent(
                    footnotesDoc,
                    ref,
                    footnotesRelsDoc,
                    footnotesNextRelIdRef,
                    (g.locators || [])[idx] || "",
                  );
                  var footnoteId = nextFootnoteId++;
                  var footnoteEl = buildFootnoteElement(
                    footnotesDoc,
                    footnoteId,
                    content,
                  );
                  footnotesRoot.appendChild(footnoteEl);
                  newFootnoteEls.push(footnoteEl);
                  footnotesCreated++;
                  if (ref && ref.pIndex != null && convertedIndexes.indexOf(ref.pIndex) === -1) {
                    convertedIndexes.push(ref.pIndex);
                  }
                  return footnoteId;
                });
              });

              found.groups
                .slice()
                .sort(function (a, b) {
                  return b.start - a.start;
                })
                .forEach(function (g) {
                  replaceSpanWithFootnoteRefs(doc, p, g.start, g.end, g.footnoteIds);
                });
            }

            var collapse = { collapsed: false, removed: 0, kept: 0 };
            if (ctx.hasRefsSection && convertedIndexes.length) {
              collapse = collapseReferenceSection(
                doc,
                paragraphs,
                ctx.refIdx,
                ctx.refEndIdx,
                convertedIndexes,
              );
            }

            // Give the document's own footnotes the same blank line the new
            // ones get. Without this, a document that had three footnotes and
            // gained nine ends up with two visibly different spacings at the
            // foot of the page. Only done when something was actually
            // converted — a document nobody asked to change isn't reformatted.
            var spacedExisting = 0;
            if (footnotesCreated) {
              elements(footnotesDoc, W_NS, "footnote").forEach(function (fn) {
                if (!isRealFootnote(fn)) return;
                if (newFootnoteEls.indexOf(fn) !== -1) return;
                if (addFootnoteSpacing(footnotesDoc, fn)) spacedExisting++;
              });
            }

            ensureDocumentRelHasFootnotes(relsDoc, nextRelIdRef);

            var stylesFile = zip.file("word/styles.xml");
            var stylesPromise = stylesFile
              ? stylesFile.async("string").then(function (stylesXml) {
                  var stylesDoc = parseXml(stylesXml, "the styles part");
                  ensureRequiredStyles(stylesDoc);
                  zip.file("word/styles.xml", serialize(stylesDoc));
                })
              : Promise.resolve();

            return stylesPromise
              .then(function () {
                return ensureContentTypeOverride(
                  zip,
                  "/word/footnotes.xml",
                  FOOTNOTES_CONTENT_TYPE,
                );
              })
              .then(function () {
                zip.file("word/document.xml", serialize(doc));
                zip.file(relsPath, serialize(relsDoc));
                zip.file("word/footnotes.xml", serialize(footnotesDoc));
                zip.file("word/_rels/footnotes.xml.rels", serialize(footnotesRelsDoc));

                return {
                  zip: zip,
                  summary: {
                    footnotesCreated: footnotesCreated,
                    referencesConverted: convertedIndexes.length,
                    referencesKept: collapse.kept,
                    referenceSectionCollapsed: collapse.collapsed,
                    existingFootnotesRespaced: spacedExisting,
                    unmatched: unmatched,
                  },
                };
              });
          });
        });
    }

    return {
      W_NS: W_NS,
      R_NS: R_NS,
      REL_NS: REL_NS,
      PLACEHOLDER_TEXT: PLACEHOLDER_TEXT,

      // Exposed for the test suite and for anyone extending the pipeline.
      parseXml: parseXml,
      serialize: serialize,
      getBody: getBody,
      listParagraphs: listParagraphs,
      isTocParagraph: isTocParagraph,
      isPlaceholderEntry: isPlaceholderEntry,
      paragraphText: paragraphText,
      paragraphStyleId: paragraphStyleId,
      headingLevel: headingLevel,
      collectRunEntries: collectRunEntries,
      collectFootnoteMarks: collectFootnoteMarks,
      readExistingFootnotes: readExistingFootnotes,
      addFootnoteSpacing: addFootnoteSpacing,
      isRealFootnote: isRealFootnote,
      isSectionHeadingParagraph: isSectionHeadingParagraph,
      isNestedSubheading: isNestedSubheading,
      findReferencesSectionIndex: findReferencesSectionIndex,
      findReferencesSectionEnd: findReferencesSectionEnd,
      buildRelsMap: buildRelsMap,
      getParagraphHyperlinkUrl: getParagraphHyperlinkUrl,
      getMaxRelId: getMaxRelId,
      getMaxFootnoteId: getMaxFootnoteId,
      addHyperlinkRelationship: addHyperlinkRelationship,
      parseReferences: parseReferences,
      makeTextRun: makeTextRun,
      makeFootnoteReferenceRun: makeFootnoteReferenceRun,
      replaceSpanWithFootnoteRefs: replaceSpanWithFootnoteRefs,
      buildFootnoteElement: buildFootnoteElement,
      buildFootnoteContent: buildFootnoteContent,
      ensureRequiredStyles: ensureRequiredStyles,
      collapseReferenceSection: collapseReferenceSection,
      clearParagraphContent: clearParagraphContent,

      analyzeDocx: analyzeDocx,
      buildConvertedDocx: buildConvertedDocx,
    };
  };
});
