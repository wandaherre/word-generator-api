const fs = require('fs');
const path = require('path');
const { createReport } = require('docx-templates');

// Dies ist die Hauptfunktion, die Vercel ausführt
module.exports = async (req, res) => {
  // 1. Nur POST-Anfragen erlauben
  if (req.method !== 'POST') {
    res.status(405).send('Method Not Allowed');
    return;
  }

  try {
    // 2. Lade die Word-Vorlage
    //    path.resolve stellt sicher, dass der Pfad korrekt ist, auch auf dem Server
    const template = fs.readFileSync(path.resolve(__dirname, 'template.docx'));

    // 3. Erstelle das Word-Dokument im Speicher (Buffer)
    //    req.body enthält die JSON-Daten, die von Ihrer App gesendet wurden
    const buffer = await createReport({
      template,
      data: req.body, // Hier werden die JSON-Daten übergeben
    });

    // 4. Setze die HTTP-Header, damit der Browser die Datei herunterlädt
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=Lernmaterial.docx');
    
    // 5. Sende die fertige Datei zurück
    res.status(200).send(buffer);

  } catch (error) {
    console.error(error);
    res.status(500).send(`Error generating document: ${error.message}`);
  }
};