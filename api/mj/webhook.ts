// /api/mj/webhook.ts
// Zweck: Kie.ai-Callback empfangen (data.resultUrls[])
// Methode: POST
// Sicherheit: optionales eigenes Secret im Header "X-Webhook-Secret"

import type { VercelRequest, VercelResponse } from '@vercel/node';

function withCORS(req: VercelRequest, res: VercelResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Webhook-Secret');
  if (req.method === 'OPTIONS') { res.status(200).end(); return true; }
  return false;
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (withCORS(req, res)) return;
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const expected = process.env.KIE_WEBHOOK_SECRET;
    const got = (req.headers['x-webhook-secret'] || req.headers['X-Webhook-Secret']) as string | undefined;
    if (expected && got !== expected) {
      return res.status(401).json({ error: 'Invalid webhook secret' });
    }

    const payload = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    // Erwartet: { code, msg, data: { taskId, promptJson, resultUrls: string[] } }
    const urls: string[] = payload?.data?.resultUrls || [];
    console.log('[KIEAI][WEBHOOK] taskId=%s urls=%d', payload?.data?.taskId, urls.length);

    // TODO: URLs in DB/Storage übernehmen (optional)
    return res.status(200).json({ ok: true, received: { taskId: payload?.data?.taskId, urls } });
  } catch (err: any) {
    console.error('[KIEAI][WEBHOOK][ERROR]', err);
    return res.status(500).json({ error: err?.message || 'Webhook error' });
  }
}
