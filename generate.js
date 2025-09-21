// api/generate.js
// Vercel Serverless (CommonJS) + docx-templates
// -> Rich-Text-Unterstützung (Bold/Italic/Linebreaks/Bullets) via RAW XML-Felder.

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- CORS ---------------- */
function setCors(req, res) {
  const origin = req.headers["origin"] || "*";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  const reqMethods = req.headers["access-control-request-method"] || "POST,OPTIONS";
  const reqHeaders = req.headers["access-control-request-headers"] || "Content-Type,Authorization";
  res.setHeader("Access-Control-Allow-Methods", reqMethods);
  res.setHeader("Access-Control-Allow-Headers", reqHeaders);
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

/* -------- Helpers: Plain & Rich -------- */
// 1) HTML/Plain nach Text (Fallback)
function stripHtmlToText(input) {
  if (input == null) return "";
  let s = String(input);
  s = s.replace(/<br\s*\/?>/gi, "\n")
       .replace(/<\/p>\s*/gi, "\n\n")
       .replace(/<p[^>]*>/gi, "");
  s = s.replace(/<li[^>]*>\s*/gi, "- ").replace(/<\/li>/gi, "\n");
  s = s.replace(/<\/?(ul|ol)[^>]*>/gi, "");
  s = s.replace(/<[^>]+>/g, "");
  s = s.replace(/\r\n?/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  return s;
}

// 2) Mini-Markdown → RAW WordprocessingML
//   unterstützt: **bold**, *italic*, __bold__, _italic_, Zeilenumbrüche, Bullet-Listen (- , • ), Nummern (1. )
//   (Keine automatischen Word-Nummerierungen – wir präfixen optisch "• " bzw. "1. ")
function mdToRawXml(input) {
  if (input == null) return { _type: "rawXml", xml: "" };
  let s = String(input);
  // HTML -> rudimentäres Markdown
  s = s.replace(/<br\s*\/?>/gi, "\n")
       .replace(/<\/p>\s*/gi, "\n\n")
       .replace(/<p[^>]*>/gi, "")
       .replace(/<\/?(ul|ol)[^>]*>/gi, "")
       .replace(/<li[^>]*>\s*/gi, "• ").replace(/<\/li>/gi, "\n");
  s = s.replace(/<\/?strong>/gi, "**").replace(/<\/?b>/gi, "**");
  s = s.replace(/<\/?em>/gi, "*").replace(/<\/?i>/gi, "*");
  s = s.replace(/<[^>]+>/g, "");
  s = s.replace(/\r\n?/g, "\n");

  // Absätze sind durch Doppellinie getrennt
  const paragraphs = s.split(/\n{2,}/).map(p => p.replace(/[ \t]+\n/g, "\n"));

  const esc = (t) => String(t)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  // Tokenize inline (**bold**, *italic*)
  function runsFor(line) {
    const out = [];
    let rest = line;
    // einfache Sequenz: erst fett, dann kursiv
    const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/;
    while (rest.length) {
      const m = rest.match(re);
      if (!m) { out.push({ t: rest }); break; }
      const [match] = m;
      const i = m.index;
      if (i > 0) out.push({ t: rest.slice(0, i) });
      if (match.startsWith("**")) out.push({ t: match.slice(2, -2), b: true });
      else out.push({ t: match.slice(1, -1), i: true });
      rest = rest.slice(i + match.length);
    }
    return out;
  }

  // Baue WordprocessingML <w:p> … <w:r><w:t>…</w:t></w:r> …
  const w = (tag, inner) => `<w:${tag}>${inner}</w:${tag}>`;
  const rPr = (r) => {
    const bits = [];
    if (r.b) bits.push("<w:b/>");
    if (r.i) bits.push("<w:i/>");
    return bits.length ? w("rPr", bits.join("")) : "";
  };

  let xml = "";
  for (const p of paragraphs) {
    const lines = p.split("\n");

    for (let idx = 0; idx < lines.length; idx++) {
      let line = lines[idx];
      // einfache Bullet-/Nummern-Erkennung am Zeilenanfang
      if (/^\s*-\s+/.test(line)) line = "• " + line.replace(/^\s*-\s+/, "");
      // Nummern 1. 2. … lassen wir stehen
      const runs = runsFor(line);
      const rXml = runs.map(r => w("r", rPr(r) + w("t", `<w:noProof/>${esc(r.t)}`))).join("");
      xml += w("p", rXml);
    }
  }
  return { _type: "rawXml", xml };
}

// baut für eine gegebene Zeichenkette beide Varianten: plain + rich
function buildPlainAndRich(val) {
  const plain = stripHtmlToText(val);
  const rich  = mdToRawXml(val);
  return { plain, rich };
}

/* ---- Absatz-Schema beidseitig unterstützen ---- */
const MAX_P = 16;
function joinParagraphs(payload) {
  const parts = [];
  for (let i = 1; i <= MAX_P; i++) {
    const k = `article_text_paragraph${i}`;
    if (payload[k]) parts.push(String(payload[k]));
  }
  if (parts.length && !payload.article_text) payload.article_text = parts.join("\n\n");

  const vocab = [];
  for (let x = 1; x <= MAX_P; x++) {
    for (let y = 1; y <= 3; y++) {
      const k = `article_vocab_p${x}_${y}`;
      if (payload[k]) vocab.push(String(payload[k]));
    }
  }
  if (vocab.length && !payload.article_vocab) payload.article_vocab = vocab.join("|");
}
function splitParagraphs(payload) {
  if (payload.article_text) {
    const paras = String(payload.article_text).split(/\n{2,}/).map(s => s.trim()).filter(Boolean).slice(0, MAX_P);
    paras.forEach((txt, i) => {
      const k = `article_text_paragraph${i+1}`;
      if (!payload[k]) payload[k] = txt;
    });
  }
  if (payload.article_vocab) {
    const words = String(payload.article_vocab).split("|").map(s => s.trim()).filter(Boolean);
    let idx = 0;
    for (let x = 1; x <= MAX_P; x++) {
      for (let y = 1; y <= 3; y++) {
        const k = `article_vocab_p${x}_${y}`;
        if (!payload[k]) payload[k] = words[idx++] || "";
      }
    }
  }
}

/* --------------- Handler --------------- */
module.exports = async (req, res) => {
  setCors(req, res);
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method Not Allowed" }); return; }

  try {
    let payload = req.body;
    if (typeof payload === "string") { try { payload = JSON.parse(payload); } catch { payload = {}; } }
    if (!payload || typeof payload !== "object") payload = {};

    // Defaults, Spiegelung
    if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
    if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // Absatz <-> Ein-Feld
    joinParagraphs(payload);
    splitParagraphs(payload);

    // Rich/Plain für Artikel + Vokabel
    const art = buildPlainAndRich(payload.article_text || "");
    payload.article_text = art.plain;
    payload.article_text_rich = art.rich;          // ← für RAW

    const voc = buildPlainAndRich((payload.article_vocab || "").split("|").join("\n"));
    payload.article_vocab = voc.plain.replace(/\n/g, " | "); // Plain weiterhin als Pipe
    payload.article_vocab_rich = mdToRawXml((payload.article_vocab || "").split("|").join("\n")); // hübsch als Zeilen

    // Rich/Plain für alle *_content Felder
    for (const k of Object.keys(payload)) {
      if (/_content$/.test(k)) {
        const v = buildPlainAndRich(payload[k]);
        payload[k] = v.plain;
        payload[`${k}_rich`] = v.rich;             // ← für RAW
      }
    }

    // Template laden & rendern
    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      rejectNullish: false,
      errorHandler: () => "" // lieber leer als 500
    });

    // DOCX zurück
    setCors(req, res);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    setCors(req, res);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
