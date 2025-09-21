// api/generate.js
// DOCX via docx-templates – rendert *_rich als Literal-XML (||...||).
// Fixes:
// 1) Help-Links: alle Varianten (…_1a/_1b/_2…) bekommen *_pretty="help".
// 2) Source-Label: aus URL Host → "Forbes", "CFR", "The Guardian" …
// 3) Idioms Wortbox: *_word_box_content_line aus vielen Alias-Feldern.
// 4) Active/Coop: **…** → <w:b/>, robuste Absätze, keine Doppel-BR.
// Hinweis: Damit Bold/Absätze ankommen, müssen *_rich-Platzhalter im
// Template als RAW gesetzt sein (z. B. {RAW active_exercise_1_content_rich}).

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- CORS ---------------- */
function setCors(req, res) {
  const origin = req.headers.origin || "*";
  const reqMethods = req.headers["access-control-request-method"] || "POST,OPTIONS";
  const reqHeaders = req.headers["access-control-request-headers"] || "Content-Type,Authorization";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  res.setHeader("Access-Control-Allow-Methods", reqMethods);
  res.setHeader("Access-Control-Allow-Headers", reqHeaders);
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

/* ---------------- URL → Publisher-Label ---------------- */
function hostToLabel(url) {
  try {
    const u = new URL(String(url));
    const host = u.hostname.replace(/^www\./, "");
    const map = {
      "cfr.org": "CFR",
      "forbes.com": "Forbes",
      "theguardian.com": "The Guardian",
      "nytimes.com": "The New York Times",
      "bbc.com": "BBC",
      "economist.com": "The Economist",
      "ft.com": "Financial Times",
      "wsj.com": "WSJ",
      "reuters.com": "Reuters",
      "apnews.com": "AP",
      "bloomberg.com": "Bloomberg",
    };
    if (map[host]) return map[host];

    // Generische Ableitung: erste Host-Component capitalizen (z. B. "vox.com" → "Vox")
    const base = host.split(".")[0];
    return base.charAt(0).toUpperCase() + base.slice(1);
  } catch {
    return "source";
  }
}

/* ---------- Tabellen/HTML → leichtes Markdown ---------- */
function normalizeTables(html) {
  if (html == null) return "";
  let s = String(html);
  s = s.replace(/<tr[^>]*>/gi, "");
  s = s.replace(/<\/tr>/gi, "\n");
  s = s.replace(/<t[hd][^>]*>/gi, "");
  s = s.replace(/<\/t[hd]>/gi, " | ");
  s = s.replace(/\s*\|\s*(\|\s*)+/g, " | ");
  return s;
}

function htmlToLightMd(input, { forActive = false } = {}) {
  if (input == null) return "";
  let s = normalizeTables(input);

  s = s.replace(/<br\s*\/?>/gi, "\n");
  s = s.replace(/<\/p>\s*/gi, "\n\n").replace(/<p[^>]*>/gi, "");
  s = s.replace(/<\/div>\s*/gi, "\n\n").replace(/<div[^>]*>/gi, "");

  // Headings → **…**
  s = s.replace(/<\/h[1-6]>\s*/gi, "\n\n")
       .replace(/<h[1-6][^>]*>/gi, "**")
       .replace(/<\/h[1-6]>/gi, "**");

  // Listen
  s = s.replace(/<li[^>]*>\s*/gi, "- ").replace(/<\/li>/gi, "\n");
  s = s.replace(/<\/?(ul|ol)[^>]*>/gi, "");

  // Inline
  s = s.replace(/<\/?strong>/gi, "**").replace(/<\/?b>/gi, "**");
  s = s.replace(/<\/?em>/gi, "*").replace(/<\/?i>/gi, "*");

  // Rest-Tags
  s = s.replace(/<[^>]+>/g, "");

  s = s.replace(/\r\n?/g, "\n");

  if (forActive) {
    // markante Abschnitte besser trennen
    s = s.replace(/(?:^|\n)\s*(Phase\s*\d+\s*:)/gi, (m, g1) => `\n\n**${g1.trim()}**\n`);
    if (!/\n{2,}/.test(s)) s = s.replace(/(?<=[.!?])\s+(?=[A-ZÄÖÜ])/g, "\n\n");
    s = s.replace(/(\s)(\d+\.\s+)/g, "$1\n$2");
  }

  // Pipes säubern
  s = s.split("\n").map(line => line.replace(/\s*\|\s*$/,"").trimEnd()).join("\n");

  s = s.replace(/\n{3,}/g, "\n\n").trim();
  return s;
}

/* ---------- Markdown Runs → Literal XML ---------- */
function escText(t){return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
function splitRuns(md){
  const out=[]; let rest=String(md);
  const re=/(\*\*[^*]+\*\*|\*[^*]+\*)/;
  while(rest.length){
    const m=rest.match(re);
    if(!m){out.push({t:rest}); break;}
    const [match]=m; const i=m.index;
    if(i>0) out.push({t:rest.slice(0,i)});
    if(match.startsWith("**")) out.push({t:match.slice(2,-2),b:true});
    else out.push({t:match.slice(1,-1),i:true});
    rest=rest.slice(i+match.length);
  }
  return out;
}
function runXml({t,b,i}) {
  const pr=(b||i)?`<w:rPr>${b?"<w:b/>":""}${i?"<w:i/>":""}</w:rPr>`:"";
  return `<w:r>${pr}<w:t xml:space="preserve">${escText(t)}</w:t></w:r>`;
}

/* ---------- Items/Rendering ---------- */
function collapseBlankItems(items){
  const out=[]; let prevBlank=false;
  for (const it of items){
    if (it.kind === "blank"){
      if (!prevBlank) { out.push(it); prevBlank = true; }
    } else { out.push(it); prevBlank = false; }
  }
  while (out[0] && out[0].kind === "blank") out.shift();
  while (out[out.length-1] && out[out.length-1].kind === "blank") out.pop();
  return out;
}
function itemsToRunsXml(items){
  const collapsed = collapseBlankItems(items);
  const parts=[];
  for (let idx=0; idx<collapsed.length; idx++){
    const it = collapsed[idx];
    const next = collapsed[idx+1];
    if (it.kind === "blank") { parts.push("<w:br/>"); continue; }
    splitRuns(it.line).forEach(r=>parts.push(runXml(r)));
    if (next && next.kind === "text") parts.push("<w:br/>");
  }
  return parts.join("");
}
function toLiteral(runsXml){ return `||${runsXml}||`; }

/* ---------- Helpers ---------- */
function sentenceUnderline(line){
  // nur expliziter Marker → lange Linie
  return String(line).replace(/___SENTENCE___/g, "_".repeat(80));
}
function stripLeadingLabels(line){
  return String(line).replace(/^\s*((?:[A-Za-z][\)\.]|[0-9]+\.)\s+|[-•]\s+)/, "");
}
function enumeratePreserveBlanks(lines){
  const items=[]; let n=0;
  for (const raw of lines) {
    const s = String(raw);
    if (s.trim().length === 0) { items.push({ kind: "blank" }); continue; }
    if (/^\s*([0-9]+\.)|[-•]\s+/.test(s)) { items.push({ kind: "text", line: s }); continue; }
    n += 1; items.push({ kind: "text", line: `${n}. ${s}` });
  }
  return items;
}
function choicesABC(lines){
  const labels="abcdefghijklmnopqrstuvwxyz".split("");
  const items=[]; let i=0;
  for (const raw of lines) {
    let s=String(raw);
    if (s.trim().length===0) { items.push({ kind: "blank" }); continue; }
    s = stripLeadingLabels(s);
    const label = labels[i] || String.fromCharCode(97+i);
    items.push({ kind:"text", line: `${label}) ${s}` });
    i++;
  }
  return items;
}

/* ---------- Robustes Einsammeln für Wortboxen ---------- */
function toArrayFlexible(v){
  if (v == null) return [];
  if (Array.isArray(v)) return v.map(x => String(x).trim()).filter(Boolean);
  const s = String(v);
  if (s.includes("|")) return s.split("|").map(t => t.trim()).filter(Boolean);
  if (s.includes("\n")) return s.split("\n").map(t => t.trim()).filter(Boolean);
  return s.split(",").map(t => t.trim()).filter(Boolean);
}
function harvestWordBox(payload, baseKey){
  // sammelt viele Aliase: *_word_box_content, *_word_box, *_wordbox, *_options, *_choices, *_words
  const candidates = [
    payload[`${baseKey}_word_box_content`],
    payload[`${baseKey}_word_box`],
    payload[`${baseKey}_wordbox`],
    payload[`${baseKey}_options`],
    payload[`${baseKey}_choices`],
    payload[`${baseKey}_words`],
  ];
  let arr = [];
  for (const c of candidates) {
    const a = toArrayFlexible(c);
    if (a.length) { arr = a; break; }
  }
  return arr;
}

/* ---------- Artikel/Idioms/Active ableiten ---------- */
const MAX_P=16;

function deriveArticle(payload){
  // Source-Label aus URL
  if (payload.source_link) payload.source_link_pretty = hostToLabel(payload.source_link);

  for(let i=1;i<=MAX_P;i++){
    const k=`article_text_paragraph${i}`;
    if(k in payload){
      const md = htmlToLightMd(payload[k]);
      const items = md
        .split(/\n/)
        .map(s => sentenceUnderline(s))
        .map(s => s.trim().length ? {kind:"text", line:s} : {kind:"blank"});
      payload[`${k}_rich`] = toLiteral(itemsToRunsXml(items));
    }
    const w1=(payload[`article_vocab_p${i}_1`]||"").toString().trim();
    const w2=(payload[`article_vocab_p${i}_2`]||"").toString().trim();
    const w3=(payload[`article_vocab_p${i}_3`]||"").toString().trim();
    const words=[w1,w2,w3].filter(Boolean);
    if(words.length){
      payload[`article_vocab_p${i}_line`] = words.join("   |   ");
      const it = words.map(s => ({kind:"text", line:s}));
      payload[`article_vocab_p${i}_rich`] = toLiteral(itemsToRunsXml(it));
    }
  }
  // optional: alles zusammen
  const paras=[];
  for(let i=1;i<=MAX_P;i++){
    const k=`article_text_paragraph${i}`;
    if(payload[k]) paras.push(htmlToLightMd(payload[k]));
  }
  if(paras.length){
    const lines = paras.flatMap(p=>p.split(/\n{2,}/)).flatMap(p=>p.split("\n"));
    const items = lines
      .map(s => sentenceUnderline(s))
      .map(s => s.trim().length ? {kind:"text", line:s} : {kind:"blank"});
    payload.article_text_all_rich = toLiteral(itemsToRunsXml(items));
  }
}

function deriveExercises(payload){
  // 1) Help-Links: alle Varianten auf *_pretty="help"
  for (const k of Object.keys(payload)) {
    if (/^help_link_/i.test(k)) {
      const url = (payload[k] || "").toString().trim();
      if (!url) continue;
      // Original-Key bekommt *_pretty
      payload[`${k}_pretty`] = "help";
      // Varianten spiegeln (…_1a → …_1 / …_1b; …_2 → …_2a)
      const m = k.match(/^(help_link_[a-z0-9]+_\d+)([ab])?$/i);
      if (m) {
        const base = m[1];
        const suffix = m[2] || "";
        if (suffix) {
          // hat a/b → spiegele auf Basis ohne a/b
          payload[`${base}_pretty`] = "help";
        } else {
          // hat kein a/b → spiegele auf a und b
          payload[`${base}a_pretty`] = "help";
          payload[`${base}b_pretty`] = "help";
        }
      }
    }
  }

  // 2) Inhalte + Wortboxen
  for(const k of Object.keys(payload)){
    // Wortboxen: explizit
    if(/_word_box_content$/i.test(k)){
      const base = k.replace(/_word_box_content$/i, "");
      const items = toArrayFlexible(payload[k]);
      payload[`${base}_word_box_content_line`] = items.join("   |   ");
      const it = items.map(s => ({kind:"text", line:s}));
      payload[`${base}_word_box_content_rich`] = toLiteral(itemsToRunsXml(it));
    }

    // Inhalte (HTML → md → runs)
    if(/_content$/i.test(k)){
      const isActive = /^active_/i.test(k);
      const isIdioms = /idioms/i.test(k);
      const isGrammar = /b1_|b2_|grammar/i.test(k) && !isIdioms;

      let md = htmlToLightMd(payload[k], { forActive: isActive });
      let lines = md.split(/\n/).map(s => s.replace(/\s*\|\s*/g, "   |   ").trimEnd());
      lines = lines.map(s => sentenceUnderline(s)); // nur Marker → lange Linie

      let items;
      if (isIdioms) {
        items = choicesABC(lines);
      } else if (isActive) {
        items = lines.map(s => s.trim().length ? {kind:"text", line:s} : {kind:"blank"});
      } else {
        items = enumeratePreserveBlanks(lines);
      }

      payload[`${k}_rich`] = toLiteral(itemsToRunsXml(items));

      // Falls dies ein Idioms-Block ist, aber KEINE explizite Wortbox vorlag:
      if (isIdioms) {
        const base = k.replace(/_content$/i, "");
        // Nur generieren, wenn noch keine *_word_box_content_line existiert:
        if (!payload.hasOwnProperty(`${base}_word_box_content_line`)) {
          const itemsWB = harvestWordBox(payload, base);
          if (itemsWB.length) {
            payload[`${base}_word_box_content_line`] = itemsWB.join("   |   ");
            const itWB = itemsWB.map(s => ({kind:"text", line:s}));
            payload[`${base}_word_box_content_rich`] = toLiteral(itemsToRunsXml(itWB));
          }
        }
      }
    }
  }
}

/* -------------------- Handler -------------------- */
module.exports = async (req, res) => {
  setCors(req, res);
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method Not Allowed" }); return; }

  try {
    let payload = req.body;
    if (typeof payload === "string"){ try{ payload=JSON.parse(payload);}catch{ payload={}; } }
    if (!payload || typeof payload !== "object") payload = {};

    // Defaults
    if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
    if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    deriveArticle(payload);
    deriveExercises(payload);

    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{","}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      errorHandler: () => ""
    });

    setCors(req, res);
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition",'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    setCors(req, res);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
