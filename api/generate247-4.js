// api/generate.js
// DOCX via docx-templates. Liefert *_rich als Literal-XML (||...||) UND
// zusätzlich *_plain ohne **/HTML, sodass dein Template – egal ob RAW oder nicht –
// ein brauchbares Ergebnis bekommt.
//
// Behebt u. a.:
// - Active/Coop: Absätze & Fett (bei RAW), andernfalls Plain ohne **
// - Idioms (MC): blockweise a)–d) je Frage (1., 2., …), keine globalen a..t
// - Matching: 2-Spalten → "left   |   right"
// - Wortbox: *_word_box_content_line wird aus vielen Alias-Feldern erzeugt
// - Help: *_pretty = "help" für ALLE Varianten (1 / 1a / 1b / 2 …)
// - Source: Label aus URL (Forbes/CFR/…)
// - Gaps: nur ___SENTENCE___ → 80x "_" (kein globales U-Längen-Gebastel)

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- CORS ---------------- */
function setCors(req, res) {
  const origin = req.headers.origin || "*";
  const reqMethods = req.headers["access-control-request-method"] || "POST,OPTIONS";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", reqMethods);
  res.setHeader("Access-Control-Allow-Headers", req.headers["access-control-request-headers"] || "Content-Type");
  res.setHeader("Cache-Control", "no-store");
}

/* ---------------- Helpers ---------------- */
function escText(t){return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
function stripInlineStars(s){ return String(s).replace(/\*\*(.+?)\*\*/g,"$1").replace(/\*([^*]+)\*/g,"$1"); }
function stripHtmlEntities(s){
  return String(s)
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#160;/g, " ");
}
function sentenceUnderline(line){ return String(line).replace(/___SENTENCE___/g, "_".repeat(80)); }

function makeHyperlinkField(url, label){
  const u = escText(String(url||""));
  const t = escText(String(label||url||"link"));
  return `<w:fldSimple w:instr="HYPERLINK &quot;${u}&quot;">
    <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>${t}</w:t></w:r>
  </w:fldSimple>`;
}

function hostToLabel(url) {
  try {
    const u = new URL(String(url));
    const host = u.hostname.replace(/^www\./,"");
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

/* ---------------- HTML → light-md ---------------- */
function normalizeTables(html){
  if (html == null) return "";
  let s = String(html);
  s = s.replace(/<tr[^>]*>/gi,"");
  s = s.replace(/<\/tr>/gi,"\n");
  s = s.replace(/<t[hd][^>]*>/gi,"");
  s = s.replace(/<\/t[hd]>/gi," | ");
  s = s.replace(/\s*\|\s*(\|\s*)+/g," | ");
  return s;
}
function htmlToLightMd(input,{forActive=false}={}){
  if (input == null) return "";
  let s = normalizeTables(input);

  // Headings → eigene Zeilen mit **…**
  s = s.replace(/<h[1-6][^>]*>(.*?)<\/h[1-6]>/gi, (_,g1)=>`\n\n**${g1.trim()}**\n\n`);

  // Absätze / Zeilen
  s = s.replace(/<br\s*\/?>/gi,"\n");
  s = s.replace(/<\/p>\s*/gi,"\n\n").replace(/<p[^>]*>/gi,"");
  s = s.replace(/<\/div>\s*/gi,"\n\n").replace(/<div[^>]*>/gi,"");

  // Listen
  s = s.replace(/<li[^>]*>\s*/gi,"- ").replace(/<\/li>/gi,"\n");
  s = s.replace(/<\/?(ul|ol)[^>]*>/gi,"");

  // inline bold/italic → Markdown
  s = s.replace(/<\/?strong>/gi,"**").replace(/<\/?b>/gi,"**");
  s = s.replace(/<\/?em>/gi,"*").replace(/<\/?i>/gi,"*");

  // Rest-Tags killen
  s = s.replace(/<[^>]+>/g,"");
  s = stripHtmlEntities(s).replace(/\r\n?/g,"\n");

  if (forActive){
    // „Phase X:“ deutlich absetzen
    s = s.replace(/(?:^|\n)\s*(Phase\s*\d+\s*:)/gi,(m,g1)=>`\n\n**${g1.trim()}**\n`);
    // Falls kaum \n existieren: satzbasierte Trennung
    if (!/\n{2,}/.test(s)) s = s.replace(/(?<=[.!?])\s+(?=[A-ZÄÖÜ])/g,"\n\n");
    // 1. 2. 3. in neue Zeile
    s = s.replace(/(\s)(\d+\.\s+)/g,"$1\n$2");
  }

  // Pipes am Ende säubern + Mehrfach-Blankzeilen reduzieren
  s = s.split("\n").map(l=>l.replace(/\s*\|\s*$/,"").trimRight()).join("\n");
  s = s.replace(/\n{3,}/g,"\n\n").trim();
  return s;
}

/* ---------------- Markdown → Runs (für *_rich) ---------------- */
function splitRuns(md){
  const out=[]; let rest=String(md);
  const re=/(\*\*[^*]+\*\*|\*[^*]+\*)/;
  while (rest.length){
    const m = rest.match(re);
    if (!m){ out.push({b:false,i:false,t:rest}); break; }
    const idx = m.index;
    if (idx>0) out.push({b:false,i:false,t:rest.slice(0,idx)});
    const token = m[0];
    if (/^\*\*.*\*\*$/.test(token))      out.push({b:true,i:false,t:token.slice(2,-2)});
    else if (/^\*.*\*$/.test(token))     out.push({b:false,i:true,t:token.slice(1,-1)});
    rest = rest.slice(idx+token.length);
  }
  return out;
}
function runXml({b,i,t}){
  const esc = escText(t);
  let props = "<w:rPr>";
  if (b) props += "<w:b/><w:bCs/>";
  if (i) props += "<w:i/><w:iCs/>";
  props += '<w:lang w:val="en-US"/></w:rPr>';
  return `<w:r>${props}<w:t xml:space="preserve">${esc}</w:t></w:r>`;
}
function itemsToRunsXml(items){
  const parts=[];
  for (const it of items){
    if (it.kind==="blank"){ parts.push("<w:p/>"); continue; }
    if (it.kind==="text"){
      const runs = splitRuns(it.line);
      parts.push("<w:p>");
      parts.push(runs.map(runXml).join(""));
      parts.push("</w:p>");
      continue;
    }
    if (it.kind==="row"){
      // einfache „Tabelle“ als Textzeile mit Spaltentrenner
      const runs = splitRuns(it.line);
      parts.push("<w:p>");
      parts.push(runs.map(runXml).join(""));
      parts.push("</w:p>");
      continue;
    }
    if (it.kind==="hr"){ parts.push("<w:p><w:r><w:t> </w:t></w:r></w:p>"); continue; }
    if (it.kind==="br"){ parts.push("<w:r><w:br/></w:r>"); continue; }
    parts.push("<w:br/>");
  }
  return parts.join("");
}
const toLiteral = xml => `||${xml}||`;

/* ---------------- Listen/Parsing ---------------- */
function stripLeadingLabels(line){
  // A. / A) / 1. / 
  return String(line).replace(/^[a-z]\s*[\.\)]\s*/i,"").replace(/^\d+\s*[\.\)]\s*/,"");
}
function enumeratePreserveBlanks(lines){
  let idx=1; const out=[];
  for (const L of lines){
    const s=L.trimRight();
    if (!s){ out.push({kind:"blank"}); continue; }
    out.push({kind:"text", line: `${idx}. ${s}`}); idx++;
  }
  return out;
}
function toArrayFlexible(v){
  if (v==null) return [];
  if (Array.isArray(v)) return v;
  const s = String(v).trim();
  if (!s) return [];
  return s.split(/\s*\|\s*/).map(x=>x.trim()).filter(Boolean);
}

/* ----- Idioms: blockweise Optionen zu jeder Frage ----- */
function buildIdiomsBlocks(md){
  const rawLines = md.split(/\n/).map(l=>l.trimRight());
  const blocks=[]; let cur=null; let optIdx=0; const labels="abcd".split("");

  for (const L of rawLines){
    const line = sentenceUnderline(L).trim();
    if (!line){ if(cur) cur.items.push({kind:"blank"}); continue; }
    if (/^\d+\./.test(line)){ // neue Frage
      if (cur) blocks.push(cur);
      cur = { question: line, items: [] };
      optIdx = 0;
      continue;
    }
    const clean = stripLeadingLabels(line);
    const lab = labels[Math.min(optIdx, labels.length-1)];
    const opt = `${lab}) ${clean}`;
    cur ? cur.items.push({kind:"text", line: opt}) : null;
    optIdx++;
  }
  if (cur) blocks.push(cur);

  // flatten: Frage, dann deren a)–d), Leerzeile
  const items=[];
  for (const b of blocks){
    items.push({kind:"text", line:b.question});
    b.items.forEach(it=>items.push(it));
    items.push({kind:"blank"});
  }
  return items;
}

/* ----- Matching: 1..n + A..D → "left   |   right" ----- */
function buildMatchingLines(md){
  const lines = md.split(/\n/).map(s=>s.trim()).filter(Boolean);
  const left=[], right[];
  for (const s of lines){
    if (/^\d+\./.test(s)) left.push(s.replace(/^\d+\.\s*/,""));
    else if (/^[A-Za-z]\./.test(s)) right.push(s.replace(/^[A-Za-z]\.\s*/,""));
    else if (s.includes("|")) {
      const [a,b] = s.split("|");
      left.push((a||"").trim()); right.push((b||"").trim());
    } else {
      // nicht erkennbar → als Text zurückgeben
      return lines.map(t=>({kind:"text", line:t}));
    }
  }
  const n = Math.max(left.length,right.length);
  const out=[];
  for (let i=0;i<n;i++){
    const l = left[i]  || "";
    const r = right[i] || "";
    out.push({kind:"row", line: `${l}   |   ${r}`});
  }
  return out;
}

/* ---------------- Article & Vocab ---------------- */
const MAX_P=16;

function deriveArticle(payload){
  if (payload.source_link) payload.source_link_pretty = hostToLabel(payload.source_link);
  if (payload.source_link) payload.source_link_hyperlink_raw = toLiteral(makeHyperlinkField(payload.source_link, payload.source_link_pretty));

  for (let i=1;i<=MAX_P;i++){
    const k = `article_text_paragraph${i}`;
    if (k in payload){
      const md = htmlToLightMd(payload[k]);
      const items = md.split(/\n/)
        .map(s=>sentenceUnderline(s))
        .map(s=>s.trim()?{kind:"text", line:s}:{kind:"blank"});
      payload[`${k}_rich`]  = toLiteral(itemsToRunsXml(items));
      payload[`${k}_plain`] = stripInlineStars(md);
    }
    const w1=(payload[`article_vocab_p${i}_1`]||"").toString().trim();
    const w2=(payload[`article_vocab_p${i}_2`]||"").toString().trim();
    const w3=(payload[`article_vocab_p${i}_3`]||"").toString().trim();
    const words=[w1,w2,w3].filter(Boolean);
    if (words.length){
      payload[`article_vocab_p${i}_line`] = words.join("   |   ");
    }
  }

  const paras=[];
  for (let i=1;i<=MAX_P;i++){
    const k=`article_text_paragraph${i}`;
    if (payload[k]) paras.push(htmlToLightMd(payload[k]));
  }
  if (paras.length){
    const lines = paras.flatMap(p=>p.split(/\n{2,}/)).flatMap(p=>p.split("\n"))
      .map(s=>sentenceUnderline(s));
    const items = lines.map(s=>s.trim()?{kind:"text", line:s}:{kind:"blank"});
    payload.article_text_all_rich  = toLiteral(itemsToRunsXml(items));
    payload.article_text_all_plain = stripInlineStars(paras.join("\n\n"));
  }
}

/* ---------------- Exercises ---------------- */
function deriveExercises(payload){
  // Help-Links: alle Varianten spiegeln
  for (const k of Object.keys(payload)){
    if (/^help_link_/i.test(k)){
      const url = (payload[k]||"").toString().trim();
      if (!url) continue;
      payload[`${k}_pretty`] = "help";
      const m = k.match(/^(help_link_[a-z0-9]+_\d+)([ab])?$/i);
      if (m){
        const base=m[1], suffix=m[2]||"";
        // Spiegel-Felder (1a/1b zu 1, etc.)
        payload[`${base}${suffix}`] = url;
      }
    }
  }

  for (const k of Object.keys(payload)){
    // Inhalte
    if (/_content$/i.test(k)){
      const isActive   = /^active_/i.test(k);
      const isIdioms   = /idioms/i.test(k);
      const isMatching = /matching/i.test(k);

      let md = htmlToLightMd(payload[k], { forActive:isActive });
      let lines = md.split(/\n/).map(s=>s.replace(/\s*\|\s*/g,"   |   ").trimEnd());
      lines = lines.map(sentenceUnderline);

      let items;
      if (isIdioms){
        items = buildIdiomsBlocks(md); // pro Frage a)–d)
        // Wortbox falls noch nicht da:
        const base = k.replace(/_content$/i,"");
        if (!payload.hasOwnProperty(`${base}_word_box_content_line`)){
          const wb = harvestWordBox(payload, base);
          if (wb.length) payload[`${base}_word_box_content_line`] = wb.join("   |   ");
        }
      } else if (isMatching){
        items = buildMatchingLines(md);
      } else if (isActive){
        items = lines.map(s=>s.trim()?{kind:"text", line:s}:{kind:"blank"});
      } else {
        items = enumeratePreserveBlanks(lines);
      }

      // Ausgaben für beide Template-Fälle:
      payload[`${k}_rich`]  = toLiteral(itemsToRunsXml(items));            // für RAW-Platzhalter
      payload[`${k}_plain`] = stripInlineStars(lines.join("\n"));          // falls Template NICHT RAW nutzt
    }
  }

  // Article: schon oben verarbeitet
}

/* ----- Wortbox sammeln ----- */
function harvestWordBox(payload, base){
  const candidates = [
    payload[`${base}_word_box_content`],
    payload[`${base}_word_box`],
    payload[`${base}_wordbox`],
    payload[`${base}_options`],
    payload[`${base}_choices`],
    payload[`${base}_words`],
  ];
  for (const c of candidates){
    const a = toArrayFlexible(c);
    if (a.length) return a;
  }
  return [];
}

/* ---------------- API Handler ---------------- */
module.exports = async (req, res) => {
  try {
    setCors(req,res);
    if (req.method === "OPTIONS"){ res.status(200).end(); return; }

    const payload = req.body || {};
    deriveArticle(payload);
    deriveExercises(payload);

    const templatePath = path.join(process.cwd(), "template.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{","}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      errorHandler: () => ""
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    res.status(500).json({ error: err?.message || String(err) });
  }
};
