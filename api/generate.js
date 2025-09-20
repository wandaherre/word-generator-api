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
    
    const data = Object.keys(req.body || {}).length === 0 ? {
      title: 'Test Document',
      content: 'This is test content.'
    } : req.body;

    const buffer = await createReport({
      template,
      data,
      cmdDelimiter: ['{{', '}}']
    });

    console.log('Generated buffer size:', buffer.length);

    // Convert to base64 and send as JSON
    const base64 = buffer.toString('base64');
    
    return res.status(200).json({
      success: true,
      document: base64,
      filename: 'generated.docx',
      size: buffer.length
    });

  } catch (error) {
    console.error('Generation error:', error.message);
    return res.status(500).json({
      error: 'Document generation failed',
      details: error.message
    });
  }
};
