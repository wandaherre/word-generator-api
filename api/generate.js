const fs = require('fs');
const path = require('path');
const { createReport } = require('docx-templates');
const { put } = require('@vercel/blob');

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
    const template = fs.readFileSync(templatePath);
    
    const data = req.body || { title: 'Test', content: 'Test content' };

    const buffer = await createReport({
      template,
      data,
      cmdDelimiter: ['{{', '}}']
    });

    // Upload to Vercel Blob Storage
    const filename = `document-${Date.now()}.docx`;
    const blob = await put(filename, buffer, {
      access: 'public',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    return res.status(200).json({
      success: true,
      downloadUrl: blob.url,
      filename: filename,
      size: buffer.length
    });

  } catch (error) {
    console.error('Error:', error.message);
    return res.status(500).json({
      error: 'Generation failed',
      details: error.message
    });
  }
};
