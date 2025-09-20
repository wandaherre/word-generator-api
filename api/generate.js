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
    console.log('Template loaded, size:', template.length);

    // Force minimal data
    const data = {
      title: 'Hardcoded Title',
      content: 'Hardcoded Content'
    };

    console.log('Using hardcoded data:', data);

    const buffer = await createReport({
      template,
      data,
      cmdDelimiter: ['{{', '}}'],
      literalXmlDelimiter: ['||', '||'],
      processLineBreaks: false,
      noSandBox: true
    });

    console.log('Document generated, size:', buffer.length);

    // Validate the buffer is not empty
    if (!buffer || buffer.length === 0) {
      throw new Error('Generated buffer is empty');
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=forced-minimal.docx');
    
    return res.status(200).send(buffer);

  } catch (error) {
    console.error('DETAILED ERROR:');
    console.error('Type:', error.constructor.name);
    console.error('Message:', error.message);
    console.error('Stack:', error.stack);
    
    return res.status(500).json({
      error: 'Generation failed',
      type: error.constructor.name,
      message: error.message
    });
  }
};
