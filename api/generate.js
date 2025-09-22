// api/generate.js
// Kompatibel mit AI-Studio-Payload (selectedExercises[] …), mappt auf flache Template-Keys.
// Fixes:
// - Idioms: Wortbox aus ALLEN options aggregieren (Fallback, Titel egal)
// - Active/Cooperative: taskDescription → *_content_rich/_plain (RAW + Plain)
// - Grammar-Hinweise (obligation, possibility …) aus Items ziehen, wenn in description fehlen
// - Source-Link: klickbarer Ein-Wort-Hyperlink via fldSimple (kein Relationship-Zwang)
// - Liefert für jeden *_content sowohl *_content_rich (RAW) als auch *_content_plain

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

/* ---------------- Utils ---------------- */
function esc(t){return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
function stripInlineMd(s){ return String(s).replace(/\*\*(.+?)\*\*/g,"$1").replace(/\*([^*]+)\*/g,"$1").replace(/_([^_]+)_/g,"$1"); }
function normalizeWs(s){ return String(s).replace(/[ \t]+\n/g,"\n").replace(/\n{3,}/g,"\n\n").trim(); }
function sentenceUnderline(line){ return String(line).replace(/___SENTENCE___/g, "_".repeat(80)); }

function hostToLabel(url) {
  try {
    const u = new URL(String(url));
    const host = u.hostname.replace(/^www\./, "");
    const map = {
      "cfr.org":"CFR","forbes.com":"Forbes","theguardian.com":"The Guardian",
      "nytimes.com":"The New York Times","bbc.com":"BBC","economist.com":"The Economist",
      "ft.com":"Financial Times","wsj.com":"WSJ","reuters.com":"Reuters",
      "apnews.com":"AP","bloomberg.com":"Bloomberg"
    };
    if (map[host]) return map[host];
    const base = host.split(".")[0];
    return base ? base[0].toUpperCase()+base.slice(1) : "source";
  } catch { return "source"; }
}

/* ------------- Markdown → <w:r>… (für *_rich) ------------- */
function splitRuns(md){
  const out=[]; let rest=String(md);
  const re=/(\*\*[^*]+\*\*|\*[^*]+\*|_[^_]+_)/;
  while(rest.length){
    const m=rest.match(re);
    if(!m){ out.push({t:rest}); break; }
    const [match]=m; const i=m.index;
    if(i>0) out.push({t:rest.slice(0,i)});
    if(match.startsWith("**")) out.push({t:match.slice(2,-2), b:true});
    else if(match.startsWith("*")) out.push({t:match.slice(1,-1), i:true});
    else out.push({t:match.slice(1,-1), i:true}); // _italic_
    rest = rest.slice(i+match.length);
  }
  return out;
}
function runXml({t,b,i}) {
  const pr=(b||i)?`<w:rPr>${b?"<w:b/>":""}${i?"<w:i/>":""}</w:rPr>`:"";
  return `<w:r>${pr}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
}
function collapseBlanks(items){
  const out=[]; let prev=false;
  for(const it of items){
    if(it.kind==="blank"){ if(!prev){out.push(it); prev=true;} }
    else { out.push(it); prev=false; }
  }
  while(out[0]?.kind==="blank") out.shift();
  while(out[out.length-1]?.kind==="blank") out.pop();
  return out;
}
function itemsToRunsXml(items){
  const arr = collapseBlanks(items);
  const parts=[];
  for(let i=0;i<arr.length;i++){
    const it=arr[i], next=arr[i+1];
    if(it.kind==="blank"){ parts.push("<w:br/>"); continue; }
    splitRuns(it.line).forEach(r=>parts.push(runXml(r)));
    if(next && next.kind==="text") parts.push("<w:br/>");
  }
  return parts.join("");
}
const toLiteral = xml => `||${xml}||`;

/* ------------- Helpers: Lines & Lists ------------- */
function toArrayFlexible(v){
  if (v == null) return [];
  if (Array.isArray(v)) return v.map(x=>String(x).trim()).filter(Boolean);
  const s = String(v);
  if (s.includes("|")) return s.split("|").map(t=>t.trim()).filter(Boolean);
  if (s.includes("\n")) return s.split("\n").map(t=>t.trim()).filter(Boolean);
  return s.split(",").map(t=>t.trim()).filter(Boolean);
}
function stripLeadingLabels(line){
  return String(line).replace(/^\s*((?:[A-Za-z][\)\.]|[0-9]+\.)\s+|[-•]\s+)/,"");
}
function enumeratePreserve(lines){
  const items=[]; let n=0;
  for (const raw of lines){
    const s = String(raw);
    if (!s.trim()){ items.push({kind:"blank"}); continue; }
    if (/^\s*([0-9]+\.)|[-•]\s+/.test(s)){ items.push({kind:"text", line:s}); continue; }
    n += 1; items.push({kind:"text", line:`${n}. ${s}`});
  }
  return items;
}
function choicesABCD(lines){
  const labels="abcd".split(""); let idx=0;
  const items=[];
  for(const raw of lines){
    const s = String(raw).trimRight();
    if(!s){ items.push({kind:"blank"}); continue; }
    const clean = stripLeadingLabels(s);
    const lab = labels[Math.min(idx, labels.length-1)];
    items.push({kind:"text", line:`${lab}) ${clean}`});
    idx++;
  }
  return items;
}
function buildIdiomsBlocks(md){
  const raw = String(md).split(/\n/).map(l=>l.trimRight());
  const blocks=[]; let cur=null; let labIdx=0;
  for(const L of raw){
    const line = sentenceUnderline(L).trim();
    if(!line){ if(cur) cur.items.push({kind:"blank"}); continue; }
    if(/^\d+\./.test(line)){
      if(cur) blocks.push(cur);
      cur = {q: line, items: []}; labIdx=0;
      continue;
    }
    const lbl = "abcd".charAt(Math.min(labIdx,3));
    const opt = `${lbl}) ${stripLeadingLabels(line)}`;
    labIdx++;
    if(!cur) blocks.push({q:"", items:[{kind:"text", line:opt}]});
    else cur.items.push({kind:"text", line:opt});
  }
  if(cur) blocks.push(cur);
  const items=[];
  for(const b of blocks){
    if (b.q) items.push({kind:"text", line:b.q});
    b.items.forEach(it=>items.push(it));
    items.push({kind:"blank"});
  }
  return items;
}
function buildMatchingLines(md){
  const lines = String(md).split(/\n/).map(s=>s.trim()).filter(Boolean);
  const left=[], right=[];
  for(const s of lines){
    if(/^\d+\./.test(s)) left.push(s.replace(/^\d+\.\s*/,""));
    else if(/^[A-Za-z]\./.test(s)) right.push(s.replace(/^[A-Za-z]\.\s*/,""));
    else if (s.includes("|")){
      const [a,b] = s.split("|");
      left.push((a||"").trim()); right.push((b||"").trim());
    } else {
      return lines.map(t=>({kind:"text", line:t}));
    }
  }
  const n = Math.max(left.length,right.length);
  const items=[];
  for(let i=0;i<n;i++){
    items.push({kind:"text", line:`${left[i]||""}   |   ${right[i]||""}`});
  }
  return items;
}

/* ------------- Source Hyperlink (fldSimple) ------------- */
function buildHyperlinkField(label, url){
  const L = esc(String(label||"source"));
  const U = esc(String(url||"#"));
  // Unterstrichen + blau, klickbar per fldSimple
  const r = `<w:r><w:rPr><w:u w:val="single"/><w:color w:val="0000FF"/></w:rPr><w:t xml:space="preserve">${L}</w:t></w:r>`;
  return `||<w:fldSimple w:instr="HYPERLINK &quot;${U}&quot;">${r}</w:fldSimple>||`;
}

/* ------------- Mapping vom AI-Studio-Payload → Template-Keys ------------- */
function mdToItems(md, {mode}={}){
  // mode: "active"|"idioms"|"matching"|"plain"
  const text = normalizeWs(md);
  let lines = text.split(/\n/).map(s => sentenceUnderline(s).replace(/\s*\|\s*/g, "   |   ")).map(s=>s.trimEnd());
  if (mode === "idioms")       return buildIdiomsBlocks(text);
  if (mode === "matching")     return buildMatchingLines(text);
  if (mode === "active")       return lines.map(s=>s?{kind:"text", line:s}:{kind:"blank"});
  // default: nummerieren, Blankzeilen erhalten
  return enumeratePreserve(lines);
}

function collectOptionSet(ex){
  // Für Idioms/MC: sammle ALLE Optionen aus allen Items
  const content = ex.content || {};
  const items = Array.isArray(content.items) ? content.items : [];
  const set = new Set();
  for (const it of items){
    const opts = Array.isArray(it?.options) ? it.options : [];
    for (const o of opts){ const v=String(o||"").trim(); if(v) set.add(v); }
  }
  // evtl. vordefinierte wordBox berücksichtigen
  const wb = Array.isArray(content.wordBox) ? content.wordBox : [];
  for (const o of wb){ const v=String(o||"").trim(); if(v) set.add(v); }
  return Array.from(set);
}

function extractParentheticalLabelsFromItems(ex){
  // z. B. "(obligation)", "(possibility)" in question/sentence_with_gap
  const content = ex.content || {};
  const items = Array.isArray(content.items) ? content.items : content;
  const labels = new Set();
  const pushLabels = (s) => {
    const m = String(s||"").match(/\(([^)]+)\)/g);
    if (m) m.forEach(tok => labels.add(tok.replace(/[()]/g,"").trim()));
  };
  if (Array.isArray(items)){
    for(const it of items){
      for (const k of Object.keys(it||{})) pushLabels(it[k]);
    }
  } else {
    for (const k of Object.keys(items)) pushLabels(items[k]);
  }
  return Array.from(labels);
}

function mapExercises(payloadFromFrontend){
  const out = {};
  const exs = Array.isArray(payloadFromFrontend?.exercises) ? payloadFromFrontend.exercises : [];

  let idxB1=0, idxB2=0, idxIdioms=0, idxActive=0;

  for (const ex of exs){
    const title = String(ex.title||"").trim();
    const desc  = String(ex.description||"").trim();
    const cat   = String(ex.category||"").toUpperCase(); // "B1","B2","IDIOMS",...
    const type  = String(ex.type||"").toUpperCase();     // e.g., "MULTIPLE_CHOICE","COOPERATIVE_TASK",...

    // Hilfsfunktion je Zielgruppe
    function setBlock(prefix, idx, richMd, expl){
      const name = `${prefix}_exercise_${idx}`;
      out[`${name}`] = title || name;
      // Explanation: desc + (fehlende) Labels aus Items
      let explanation = desc;
      if (!/\(.*\)/.test(desc)) {
        const labels = extractParentheticalLabelsFromItems(ex);
        if (labels.length) explanation = desc ? `${desc} (${labels.join(", ")})` : `(${labels.join(", ")})`;
      }
      if (expl) explanation = expl;
      out[`${name}_explanation`] = explanation;

      // Inhalt
      const items = mdToItems(richMd, { mode: prefix==="active" ? "active"
                                     : (prefix==="idioms" && type==="MULTIPLE_CHOICE") ? "idioms"
                                     : (title.toLowerCase().includes("matching")||type==="MATCHING") ? "matching"
                                     : "plain" });
      const rich = toLiteral(itemsToRunsXml(items));
      out[`${name}_content_rich`]  = rich;
      out[`${name}_content_plain`] = stripInlineMd(normalizeWs(richMd));
      return name;
    }

    // CATEGORY/TYPE Routing
    if (cat==="IDIOMS" || /idiom/i.test(title) || type==="MULTIPLE_CHOICE" && (ex.content?.isWordBox===true)) {
      // Idioms: MC-Blocks + Wortbox
      idxIdioms += 1;
      const name = setBlock("idioms", idxIdioms, stringifyIdiomsMd(ex), null);
      // Wortbox aus allen Optionen aggregieren:
      const pool = collectOptionSet(ex);
      if (pool.length) out[`${name}_word_box_content_line`] = pool.join("   |   ");

    } else if (type==="COOPERATIVE_TASK") {
      // Active / Cooperative
      idxActive += 1;
      const md = String(ex?.content?.taskDescription || "").trim();
      setBlock("active", idxActive, md, desc);

    } else if (cat==="B2") {
      // B2: diverse Typen
      idxB2 += 1;
      const md = stringifyGenericMd(ex);
      const name = setBlock("b2", idxB2, md, null);
      // optionale Wortbox (falls gegeben, z. B. connectors)
      const box = extractWordBox(ex);
      if (box.length) out[`${name}_word_box_content_line`] = box.join("   |   ");

    } else if (cat==="B1") {
      // B1
      idxB1 += 1;
      const md = stringifyGenericMd(ex);
      const name = setBlock("b1", idxB1, md, null);
      const box = extractWordBox(ex);
      if (box.length) out[`${name}_word_box_content_line`] = box.join("   |   ");

    } else {
      // Fallback: behandle wie B1 (damit es nicht leer fällt)
      idxB1 += 1;
      const md = stringifyGenericMd(ex);
      const name = setBlock("b1", idxB1, md, null);
      const box = extractWordBox(ex);
      if (box.length) out[`${name}_word_box_content_line`] = box.join("   |   ");
    }
  }

  return out;
}

/* --------- MD-Erzeugung je Übungstyp aus AI-Studio-Content --------- */
function stringifyIdiomsMd(ex){
  // Erwartet content.items[].question + options; wir bauen "1."-Fragen + Optionen als Zeilen
  const content = ex.content || {};
  const items = Array.isArray(content.items) ? content.items : [];
  const parts=[];
  let qn=0;
  for (const it of items){
    const q = String(it?.question||"").trim();
    if (q){ qn+=1; parts.push(`${qn}. ${q}`); }
    const opts = Array.isArray(it?.options)? it.options : [];
    for (const o of opts){ const v=String(o||"").trim(); if (v) parts.push(v); }
    parts.push(""); // Leerzeile zwischen Blöcken
  }
  return parts.join("\n");
}

function extractWordBox(ex){
  // Für Fill/Connectors usw.
  const content = ex.content || {};
  // explizite wordBox (Array) bevorzugen
  if (Array.isArray(content.wordBox) && content.wordBox.length) {
    return content.wordBox.map(x=>String(x||"").trim()).filter(Boolean);
  }
  // aus items ggf. "options" (bei MC ohne idioms)
  const items = Array.isArray(content.items) ? content.items : [];
  const pool = [];
  for (const it of items){
    if (Array.isArray(it?.options)) for (const o of it.options) pool.push(String(o||"").trim());
    if (it?.hint) pool.push(String(it.hint).trim());
  }
  return pool.filter(Boolean);
}

function stringifyGenericMd(ex){
  const type = String(ex.type||"").toUpperCase();
  const content = ex.content || {};

  if (type==="FILL_IN_THE_BLANK") {
    const items = Array.isArray(content) ? content
                : Array.isArray(content.items) ? content.items : [];
    return items.map(it => String(it?.text||it?.sentence_with_gap||"").trim()).join("\n");

  } else if (type==="CONNECTORS_GAP_FILL") {
    const items = Array.isArray(content?.items) ? content.items : [];
    return items.map(it => String(it?.text||it?.sentence_with_gap||"").trim()).join("\n");

  } else if (type==="MULTIPLE_CHOICE") {
    // Nicht-Idioms MC → einfache Aufzählung (Frage + Optionen)
    const items = Array.isArray(content?.items) ? content.items : [];
    const parts=[];
    let qn=0;
    for (const it of items){
      const q = String(it?.question||"").trim();
      if (q){ qn+=1; parts.push(`${qn}. ${q}`); }
      const opts = Array.isArray(it?.options)? it.options : [];
      for (const o of opts){ const v=String(o||"").trim(); if(v) parts.push(v); }
      parts.push("");
    }
    return parts.join("\n");

  } else if (type==="MATCHING") {
    // term/definition → 1.. + A.. erzeugen
    const items = Array.isArray(content) ? content
                : Array.isArray(content.items) ? content.items : [];
    const left = [], right = [];
    let i=0;
    for (const it of items){
      const t = String(it?.term||"").trim();
      const d = String(it?.definition||"").trim();
      if (t) { i+=1; left.push(`${i}. ${t}`); }
      if (d) right.push(String.fromCharCode(65+right.length) + ". " + d);
    }
    return left.concat([""], right).join("\n");

  } else if (type==="SENTENCE_TRANSFORMATION") {
    const items = Array.isArray(content) ? content
                : Array.isArray(content.items) ? content.items : [];
    return items.map(it => String(it?.original||it?.original_sentence||"").trim()).join("\n");

  } else if (type==="OPEN_ENDED") {
    const items = Array.isArray(content) ? content
                : Array.isArray(content.items) ? content.items : [];
    return items.map(it => String(it?.task||"").trim()).join("\n");

  } else if (type==="COOPERATIVE_TASK") {
    // wird separat in mapExercises behandelt
    return String(ex?.content?.taskDescription||"");

  } else {
    // Fallback: versuche generisch
    if (Array.isArray(content?.items)) {
      return content.items.map(it => Object.values(it||{}).join(" ")).join("\n");
    }
    return Object.values(content||{}).join("\n");
  }
}

/* -------------------- Haupt-Handler -------------------- */
module.exports = async (req, res) => {
  setCors(req, res);
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method Not Allowed" }); return; }

  try {
    let body = req.body;
    if (typeof body === "string"){ try{ body=JSON.parse(body); }catch{ body={}; } }
    if (!body || typeof body !== "object") body = {};

    const data = {};

    // Meta aus AI-Studio
    const sourceUrl = body.sourceUrl || body.source_link || "";
    const topicName = body.topicName || body.themenbereich || "";
    const subTopic  = body.subTopicName || body.unterthema_des_themenbereichs || "";

    if (topicName) data.themenbereich = topicName;
    if (subTopic)  data.unterthema_des_themenbereichs = subTopic;
    if (sourceUrl){
      data.source_link = sourceUrl;
      data.source_link_pretty = hostToLabel(sourceUrl);
      data.source_link_hyperlink_raw = buildHyperlinkField(data.source_link_pretty, sourceUrl);
    }

    // Headline-Spiegelung (falls vorhanden)
    if (body.headline_article) data.headline_article = body.headline_article;
    if (body.headline_artikel) data.headline_artikel = body.headline_artikel;
    if (data.headline_article && !data.headline_artikel) data.headline_artikel = data.headline_article;
    if (data.headline_artikel && !data.headline_article) data.headline_article = data.headline_artikel;

    // Exercises mappen
    Object.assign(data, mapExercises(body));

    // Template laden & rendern
    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));
    const docBuffer = await createReport({
      template: templateBuffer,
      data,
      cmdDelimiter: ["{","}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      errorHandler: () => ""
    });

    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition",'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    res.status(500).json({ error: err?.message || String(err) });
  }
};
