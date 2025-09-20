// api/generate.js
// Liefert ein echtes DOCX (Binary) statt JSON. CORS & OPTIONS korrekt.

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

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

module.exports = async (req, res) => {
  setCors(req, res);

  // Preflight (keinen Body senden)
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method Not Allowed" }); return; }

  try {
    // Body robust lesen
    let payload = req.body;
    if (typeof payload === "string") { try { payload = JSON.parse(payload); } catch { payload = {}; } }
    if (!payload || typeof payload !== "object") payload = {};

    // Minimal-Defaults gegen undefined-Ausdrücke im Template
    const ensure = (k, v = "") => { if (typeof payload[k] === "undefined") payload[k] = v; };
    ensure("midjourney_article_logo");  // falls im Template referenziert
    ensure("teacher_cloud_logo");
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // Template laden (liegt per includeFiles gebundled nebenan)
    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));

    // DOCX erzeugen – docx-templates nutzt SINGLE BRACES {key}
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      rejectNullish: false,
      errorHandler: () => ""   // statt 500 leeren String einsetzen
    });

    // Nur Binary senden (kein JSON-Wrapper!)
    setCors(req, res); // CORS auch auf Erfolgsantwort sicherstellen
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    setCors(req, res);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
