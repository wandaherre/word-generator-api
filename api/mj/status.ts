// Vercel Serverless Function: GET /api/mj/status?taskId=...
// Zweck: Task-Status abfragen und (falls fertig) resultUrls liefern

import type { VercelRequest, VercelResponse } from '@vercel/node';

function withCORS(req: VercelRequest, res: VercelResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return true;
  }
  return false;
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (withCORS(req, res)) return;
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  const API_KEY = process.env.KIE_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'Missing KIE_API_KEY env var' });

  try {
    const taskId = (req.query?.taskId ?? '') as string;
    if (!taskId) return res.status(422).json({ error: 'Missing "taskId" query param' });

    const url = new URL('https://api.kie.ai/api/v1/mj/record-info');
    url.searchParams.set('taskId', taskId);

    const upstream = await fetch(url.toString(), {
      method: 'GET',
      headers: { 'Authorization': `Bearer ${API_KEY}` },
    });

    const data = await upstream.json().catch(() => ({}));
    if (!upstream.ok) {
      return res.status(upstream.status).json(data || { error: `Upstream ${upstream.status}` });
    }

    // Erwartet u. a.:
    // { code:200, data:{ successFlag: boolean, resultInfoJson:{ resultUrls: string[] } } }
    return res.status(200).json(data);
  } catch (err: any) {
    return res.status(500).json({ error: err?.message || 'Internal error' });
  }
}
