"use strict";

/**
 * Test helpers: build .docx packages in memory and read them back.
 *
 * Fixtures are constructed here rather than checked in as binary files so that
 * every test states, in the test itself, exactly which structural quirk it is
 * exercising — a run split mid-citation, a hyperlink-wrapped URL, a generated
 * table of contents — instead of hiding it inside an opaque Word document.
 */

const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const JSZip = require("jszip");
const path = require("path");

const core = require(path.join(__dirname, "..", "..", "miwim.core.js"));
const createDocxEngine = require(path.join(__dirname, "..", "..", "miwim.docx.js"));

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

function makeEngine() {
  return createDocxEngine({ DOMParser, XMLSerializer, JSZip, core });
}

const engine = makeEngine();

// ---------------------------------------------------------------------------
// Building
// ---------------------------------------------------------------------------

function esc(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function runXml(text) {
  return '<w:r><w:t xml:space="preserve">' + esc(text) + "</w:t></w:r>";
}

/**
 * Build one <w:p>.
 *
 * `content` is a string, or an array of pieces. A string piece becomes its own
 * <w:r>, which is how a citation gets deliberately split across several runs —
 * exactly what Word does whenever spell-check or a revision boundary lands
 * mid-word. An object piece may carry `{ text, rid }` to wrap that run in a
 * <w:hyperlink>, or `{ instrText }` to emit a field code.
 */
function p(content, opts) {
  opts = opts || {};
  const pieces = Array.isArray(content) ? content : [content];

  let pPr = "";
  const props = [];
  if (opts.style) props.push('<w:pStyle w:val="' + esc(opts.style) + '"/>');
  if (opts.numId != null) {
    props.push(
      '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + opts.numId + '"/></w:numPr>',
    );
  }
  if (props.length) pPr = "<w:pPr>" + props.join("") + "</w:pPr>";

  const body = pieces
    .map(function (piece) {
      if (typeof piece === "string") return runXml(piece);
      if (piece.footnoteRef != null) {
        // A footnote marker the document already contains, exactly as Word
        // writes one: a FootnoteReference-styled run wrapping w:footnoteReference.
        return (
          '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>' +
          '<w:footnoteReference w:id="' + piece.footnoteRef + '"/></w:r>'
        );
      }
      if (piece.instrText) {
        return (
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:instrText xml:space="preserve">' +
          esc(piece.instrText) +
          "</w:instrText></w:r>" +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          runXml(piece.text || "") +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        );
      }
      if (piece.rid) {
        return (
          '<w:hyperlink r:id="' + esc(piece.rid) + '" w:history="1">' +
          runXml(piece.text) +
          "</w:hyperlink>"
        );
      }
      return runXml(piece.text);
    })
    .join("");

  return "<w:p>" + pPr + body + "</w:p>";
}

/**
 * Build a word/footnotes.xml containing pre-existing notes, alongside the two
 * synthetic separator entries every document carries.
 *
 * Each note is `{ id, text }`, or `{ id, url, rid }` to write the note as a
 * hyperlink — the shape a previously converted document has. Pass
 * `trailingBlank: true` to give a note the blank paragraph MIWIM adds, so
 * a test can check that running twice doesn't keep adding them.
 */
function footnotesPart(notes) {
  const body = (notes || [])
    .map(function (note) {
      let content;
      if (note.rid) {
        content =
          '<w:hyperlink r:id="' + note.rid + '" w:history="1">' +
          '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>' +
          esc(note.url) +
          "</w:t></w:r></w:hyperlink>";
      } else {
        content = runXml(note.text || "");
      }
      const blank = note.trailingBlank
        ? '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr></w:p>'
        : "";
      return (
        '<w:footnote w:id="' + note.id + '">' +
        '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' +
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
        '<w:r><w:t xml:space="preserve"> </w:t></w:r>' +
        content +
        "</w:p>" +
        blank +
        "</w:footnote>"
      );
    })
    .join("");

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:footnotes xmlns:w="' + W_NS + '" xmlns:r="' + R_NS + '">' +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
    body +
    "</w:footnotes>"
  );
}

/** Relationships for a footnotes part whose notes link out. */
function footnotesRels(rels) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="' + REL_NS + '">' +
    (rels || [])
      .map(function (rel) {
        return (
          '<Relationship Id="' + rel.id + '" Type="' + R_NS + '/hyperlink" ' +
          'Target="' + esc(rel.target) + '" TargetMode="External"/>'
        );
      })
      .join("") +
    "</Relationships>"
  );
}

/** Wrap paragraphs in a single-cell table, to test nested-paragraph discovery. */
function table(paragraphsXml) {
  return (
    "<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w='9000' w:type='dxa'/></w:tcPr>" +
    paragraphsXml.join("") +
    "</w:tc></w:tr></w:tbl>"
  );
}

/** Wrap paragraphs in a content control, as a generated TOC is stored. */
function sdt(paragraphsXml) {
  return (
    "<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val='Table of Contents'/>" +
    "</w:docPartObj></w:sdtPr><w:sdtContent>" +
    paragraphsXml.join("") +
    "</w:sdtContent></w:sdt>"
  );
}

const NS_DECL =
  'xmlns:w="' + W_NS + '" xmlns:r="' + R_NS + '"';

const DEFAULT_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="' + W_NS + '">' +
  '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  "</w:styles>";

const EMPTY_FOOTNOTES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:footnotes xmlns:w="' + W_NS + '" xmlns:r="' + R_NS + '">' +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  "</w:footnotes>";

/**
 * Assemble a .docx.
 *
 * options:
 *   body            array of XML strings (use p(), table(), sdt())
 *   rels            [{ id, target }] external hyperlink relationships
 *   footnotesXml    override word/footnotes.xml (null omits the part entirely)
 *   stylesXml       override word/styles.xml   (null omits the part entirely)
 *   contentTypes    when false, omit the footnotes override
 */
function buildDocx(options) {
  options = options || {};
  const bodyXml = (options.body || []).join("");

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    "<w:document " + NS_DECL + "><w:body>" +
    bodyXml +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
    "</w:body></w:document>";

  const relList = [
    '<Relationship Id="rId1" Type="' + R_NS + '/styles" Target="styles.xml"/>',
  ];
  if (options.footnotesXml !== null) {
    relList.push(
      '<Relationship Id="rId2" Type="' + R_NS + '/footnotes" Target="footnotes.xml"/>',
    );
  }
  (options.rels || []).forEach(function (rel) {
    relList.push(
      '<Relationship Id="' + rel.id + '" Type="' + R_NS + '/hyperlink" ' +
        'Target="' + esc(rel.target) + '" TargetMode="External"/>',
    );
  });

  const relsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="' + REL_NS + '">' + relList.join("") + "</Relationships>";

  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
  ];
  if (options.contentTypes !== false && options.footnotesXml !== null) {
    overrides.push(
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    );
  }

  const contentTypesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    overrides.join("") +
    "</Types>";

  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' + REL_NS + '">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      "</Relationships>",
  );
  zip.file("word/document.xml", documentXml);
  zip.file("word/_rels/document.xml.rels", relsXml);

  if (options.stylesXml !== null) {
    zip.file("word/styles.xml", options.stylesXml || DEFAULT_STYLES);
  }
  if (options.footnotesXml !== null) {
    zip.file("word/footnotes.xml", options.footnotesXml || EMPTY_FOOTNOTES);
    if (options.footnotesRelsXml) {
      zip.file("word/_rels/footnotes.xml.rels", options.footnotesRelsXml);
    }
  }

  return zip.generateAsync({ type: "nodebuffer" });
}

// ---------------------------------------------------------------------------
// Reading back
// ---------------------------------------------------------------------------

function textOf(node) {
  const ts = node.getElementsByTagNameNS(W_NS, "t");
  let out = "";
  for (let i = 0; i < ts.length; i++) out += ts[i].textContent || "";
  return out;
}

function attrNS(el, ns, qname) {
  const local = qname.split(":")[1];
  let v = null;
  try {
    v = el.getAttributeNS(ns, local);
  } catch (e) {
    v = null;
  }
  if (!v) v = el.getAttribute(qname);
  return v || null;
}

/**
 * Read a generated package back into a plain object suitable for assertions:
 * paragraph texts with footnote markers rendered inline as {n}, the footnote
 * bodies with their resolved link targets, and the raw parts for anything a
 * test needs to inspect directly.
 */
async function readDocx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new DOMParser();

  const docXml = await zip.file("word/document.xml").async("string");
  const doc = parser.parseFromString(docXml, "application/xml");

  const footnoteRels = {};
  const fnRelsFile = zip.file("word/_rels/footnotes.xml.rels");
  if (fnRelsFile) {
    const relsDoc = parser.parseFromString(
      await fnRelsFile.async("string"),
      "application/xml",
    );
    const rels = relsDoc.getElementsByTagName("Relationship");
    for (let i = 0; i < rels.length; i++) {
      footnoteRels[rels[i].getAttribute("Id")] = rels[i].getAttribute("Target");
    }
  }

  const documentRels = {};
  const docRelsFile = zip.file("word/_rels/document.xml.rels");
  if (docRelsFile) {
    const relsDoc = parser.parseFromString(
      await docRelsFile.async("string"),
      "application/xml",
    );
    const rels = relsDoc.getElementsByTagName("Relationship");
    for (let i = 0; i < rels.length; i++) {
      documentRels[rels[i].getAttribute("Id")] = {
        type: rels[i].getAttribute("Type"),
        target: rels[i].getAttribute("Target"),
      };
    }
  }

  // Paragraph text with footnote markers made visible as {n}.
  const paragraphs = [];
  const pEls = doc.getElementsByTagNameNS(W_NS, "p");
  for (let i = 0; i < pEls.length; i++) {
    const pEl = pEls[i];
    let text = "";
    const walk = function (node) {
      for (let c = 0; c < node.childNodes.length; c++) {
        const child = node.childNodes[c];
        if (child.nodeType !== 1) continue;
        if (child.localName === "t") text += child.textContent || "";
        else if (child.localName === "footnoteReference")
          text += "{" + attrNS(child, W_NS, "w:id") + "}";
        else if (child.localName !== "pPr" && child.localName !== "rPr") walk(child);
      }
    };
    walk(pEl);
    paragraphs.push(text);
  }

  const footnotes = [];
  const fnFile = zip.file("word/footnotes.xml");
  let footnotesDoc = null;
  if (fnFile) {
    footnotesDoc = parser.parseFromString(
      await fnFile.async("string"),
      "application/xml",
    );
    const fnEls = footnotesDoc.getElementsByTagNameNS(W_NS, "footnote");
    for (let i = 0; i < fnEls.length; i++) {
      const fn = fnEls[i];
      const type = attrNS(fn, W_NS, "w:type");
      if (type) continue; // separator / continuationSeparator
      const links = fn.getElementsByTagNameNS(W_NS, "hyperlink");
      const url =
        links.length > 0
          ? footnoteRels[attrNS(links[0], R_NS, "r:id")] || null
          : null;
      const paras = fn.getElementsByTagNameNS(W_NS, "p");
      let trailingBlanks = 0;
      for (let k = paras.length - 1; k >= 0; k--) {
        if (textOf(paras[k]).trim()) break;
        trailingBlanks++;
      }
      footnotes.push({
        id: attrNS(fn, W_NS, "w:id"),
        text: textOf(fn).trim(),
        paragraphCount: paras.length,
        trailingBlanks: trailingBlanks,
        url: url,
        styleId: (function () {
          const s = fn.getElementsByTagNameNS(W_NS, "pStyle")[0];
          return s ? attrNS(s, W_NS, "w:val") : null;
        })(),
        hasFootnoteRef: fn.getElementsByTagNameNS(W_NS, "footnoteRef").length > 0,
      });
    }
  }

  const stylesFile = zip.file("word/styles.xml");
  const styleIds = [];
  if (stylesFile) {
    const stylesDoc = parser.parseFromString(
      await stylesFile.async("string"),
      "application/xml",
    );
    const sEls = stylesDoc.getElementsByTagNameNS(W_NS, "style");
    for (let i = 0; i < sEls.length; i++) {
      styleIds.push(attrNS(sEls[i], W_NS, "w:styleId"));
    }
  }

  const contentTypes = await zip.file("[Content_Types].xml").async("string");

  return {
    zip: zip,
    doc: doc,
    footnotesDoc: footnotesDoc,
    paragraphs: paragraphs,
    footnotes: footnotes,
    footnoteRels: footnoteRels,
    documentRels: documentRels,
    styleIds: styleIds,
    contentTypes: contentTypes,
  };
}

/** analyze -> build -> read, the path a real conversion takes. */
async function convert(buffer, mutateRefs) {
  const analysis = await engine.analyzeDocx(buffer);
  if (mutateRefs) mutateRefs(analysis.refs);
  const built = await engine.buildConvertedDocx(buffer, analysis.refs);
  const out = await built.zip.generateAsync({ type: "nodebuffer" });
  const read = await readDocx(out);
  return { analysis: analysis, summary: built.summary, out: read, buffer: out };
}

/** Parse a standalone paragraph XML string, for unit-testing the OOXML helpers. */
function parseParagraph(xml) {
  const doc = new DOMParser().parseFromString(
    '<w:document ' + NS_DECL + "><w:body>" + xml + "</w:body></w:document>",
    "application/xml",
  );
  return doc.getElementsByTagNameNS(W_NS, "p")[0];
}

function parseBody(bodyXml) {
  const doc = new DOMParser().parseFromString(
    '<w:document ' + NS_DECL + "><w:body>" + bodyXml + "</w:body></w:document>",
    "application/xml",
  );
  return { doc: doc, body: doc.getElementsByTagNameNS(W_NS, "body")[0] };
}

module.exports = {
  W_NS,
  R_NS,
  REL_NS,
  core,
  engine,
  makeEngine,
  JSZip,
  DOMParser,
  XMLSerializer,
  p,
  table,
  sdt,
  footnotesPart,
  footnotesRels,
  buildDocx,
  readDocx,
  convert,
  parseParagraph,
  parseBody,
  esc,
};
