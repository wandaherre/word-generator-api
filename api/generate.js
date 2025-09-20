const fs = require('fs');
const path = require('path');
const { createReport } = require('docx-templates');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const templatePath = path.resolve(__dirname, 'template.docx');
    
    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ error: 'Template not found' });
    }

    const template = fs.readFileSync(templatePath);
    
    // Use simple test data if no body provided
    const data = Object.keys(req.body || {}).length === 0 ? {
      title: 'Test Document',
      content: 'This is generated content.'
    } : req.body;

    console.log('Generating document with data:', JSON.stringify(data, null, 2));

    const buffer = await createReport({
      template,
      data,
      cmdDelimiter: ['{{', '}}']
    });

    console.log('Document generated successfully, size:', buffer.length);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=generated.docx');
    
    return res.status(200).send(buffer);

  } catch (error) {
    console.error('Generation error:', error.message);
    console.error('Stack:', error.stack);
    
    return res.status(500).json({
      error: 'Document generation failed',
      details: error.message
    });
  }
};
