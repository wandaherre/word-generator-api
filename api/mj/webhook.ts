// Vercel Serverless Function: POST /api/mj/webhook
// Zweck: Kie.ai-Callback empfangen (data.resultUrls[]), weiterverarbeiten (Storage/DB)
// Hinweis: Kie.ai publiziert keine festen Callback-IPs → nutze Secret-Header zur Verifikation.

import type { VercelRequest, VercelResponse } from '@vercel/node';

function withCORS(req: VercelRequest, res: VercelResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Webhook-Secret');
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return true;
  }
  return false;
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (withCORS(req, res)) return;
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    // Optional: eigenes Secret prüfen
    const expected = process.env.KIE_WEBHOOK_SECRET;
    const got = (req.headers['x-webhook-secret'] || req.headers['X-Webhook-Secret']) as string | undefined;
    if (expected && got !== expected) {
      return res.status(401).json({ error: 'Invalid webhook secret' });
    }

    const payload = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    // Erwarteter Body (vereinfachtes Schema):
    // { code, msg, data: { taskId, promptJson, resultUrls: string[] } }
    const urls = payload?.data?.resultUrls || [];

    // TODO: Hier Bild-URLs in deinen Storage/DB übernehmen
    // z.B. fetch(urls[0]) -> in S3/R2 hochladen; oder in DB referenzieren.

    return res.status(200).json({ ok: true, received: { taskId: payload?.data?.taskId, urlsCount: urls.length } });
  } catch (err: any) {
    return res.status(500).json({ error: err?.message || 'Webhook error' });
  }
}
