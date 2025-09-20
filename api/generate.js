const fs = require('fs');
const path = require('path');
const { createReport } = require('docx-templates');

module.exports = async (req, res) => {
  // CORS-Header setzen (zusätzlich zur vercel.json)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With, Content-Type, Accept, Authorization');

  // OPTIONS-Request für Preflight
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  // Nur POST erlauben
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method Not Allowed' });
    return;
  }

  try {
    const template = fs.readFileSync(path.resolve(__dirname, 'template.docx'));
    
    const buffer = await createReport({
      template,
      data: req.body,
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=Lernmaterial.docx');
    
    res.status(200).send(buffer);

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: `Error generating document: ${error.message}` });
  }
};
