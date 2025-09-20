// api/generate.js
// Runtime: Vercel Node.js Serverless Function (CommonJS)
// Erwartet: vercel.json mit includeFiles: "api/template.docx"
// Abhängigkeit: "docx-templates" in package.json

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates");

module.exports = async (req, res) => {
  // --- CORS ---
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization");
  res.setHeader("Cache-Control", "no-store");

  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    // JSON-Body (AI Studio sendet application/json)
    const payload = req.body && typeof req.body === "object" ? req.body : {};

    // Template aus dem gebundleten Pfad lesen (dank includeFiles vorhanden)
    const templatePath = path.join(__dirname, "template.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    // DOCX generieren – docx-templates nutzt SINGLE BRACES {key}
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"]
    });

    // Binary-Response (kein JSON/kein Text!)
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');

    // WICHTIG: Nur Buffer senden
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    // Fehler klar als JSON zurückgeben (kein 200er mit HTML)
    res
      .status(500)
      .json({ error: err && err.message ? err.message : String(err) });
  }
};
