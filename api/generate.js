// api/generate.js
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---- CORS: dynamisch & überall ---- */
function setCors(req, res) {
  const origin = req.headers["origin"] || "*";
  const reqMethods = req.headers["access-control-request-method"] || "POST,OPTIONS";
  const reqHeaders = req.headers["access-control-request-headers"] || "Content-Type,Authorization";

  // Spiegel den Origin (kein "*", damit Credentials erlaubt sind)
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  res.setHeader("Access-Control-Allow-Methods", reqMethods);
  res.setHeader("Access-Control-Allow-Headers", reqHeaders);
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

module.exports = async (req, res) => {
  setCors(req, res);

  // Preflight sauber beenden
  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }

  if (req.method !== "POST") {
    setCors(req, res);
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    // Body robust lesen
    let payload = req.body;
    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch { payload = {}; }
    }
    if (!payload || typeof payload !== "object") payload = {};

    // Unkritische Defaults (verhindert Template-Fehler)
    if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
    if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // Template laden
    const template = fs.readFileSync(path.join(__dirname, "template.docx"));

    // DOCX erzeugen
    const buffer = await createReport({
      template,
      data: payload,
      cmdDelimiter: ["{", "}"],
      rejectNullish: false,
      errorHandler: () => "" // nie mit 500 sterben wegen einzelner Felder
    });

    // Binary senden + CORS beibehalten
    setCors(req, res);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(buffer));
  } catch (err) {
    setCors(req, res);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
