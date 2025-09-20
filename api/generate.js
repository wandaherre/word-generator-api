// api/generate.js
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates"); // <- docx-templates (single braces)

module.exports = async (req, res) => {
  // CORS (auch Preflight)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization");
  if (req.method === "OPTIONS") { res.status(200).end(); return; }

  if (req.method !== "POST") { res.status(405).end(); return; }

  try {
    const templatePath = path.join(__dirname, "template.docx"); // dank includeFiles vorhanden
    const template = fs.readFileSync(templatePath);

    const buffer = await createReport({
      template,
      data: req.body || {},
      cmdDelimiter: ["{", "}"] // explizit: single braces
    });

    res.setHeader("Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", "attachment; filename=generated.docx");
    res.status(200).send(Buffer.from(buffer)); // nur Binary, kein extra Text!
  } catch (err) {
    // Fehler als JSON (kein 200er mit HTML-Body!)
    res.status(500).json({ error: String(err && err.message || err) });
  }
};
