import type { NextApiRequest, NextApiResponse } from 'next';

export const config = { api: { bodyParser: true } }; // Kie.ai sendet JSON per POST

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  // Beispiel-Callback laut Docs: { code, msg, data:{ taskId, promptJson, resultUrls:[] } }
  try {
    const payload = req.body;
    // TODO: Optional: Signaturprüfung, wenn verfügbar
    // Speichere Links in DB / triggert weiteren Prozess (z. B. Download → S3)
    // payload.data.resultUrls ist ein Array mit Bild- oder Video-URLs
    // -> Hier nur zurückspiegeln
    return res.status(200).json({ ok: true });
  } catch (e:any) {
    return res.status(500).json({ error: e?.message || 'Webhook error' });
  }
}
