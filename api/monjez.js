const ALLOWED_METHODS = new Set([
  'activateMonjez',
  'completeMonjezProfile',
  'resumeMonjezSession',
  'getMonjezData',
  'saveMonjezData',
]);

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'Method not allowed' });
  }

  const gasUrl = process.env.MONJEZ_GAS_URL;
  if (!gasUrl) {
    return res.status(500).json({ ok: false, error: 'إعداد MONJEZ_GAS_URL مفقود في Vercel.' });
  }

  const body = req.body || {};
  const method = String(body.method || '');
  const args = Array.isArray(body.args) ? body.args : [];
  if (!ALLOWED_METHODS.has(method)) {
    return res.status(400).json({ ok: false, error: 'طلب API غير مسموح.' });
  }

  try {
    const upstream = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ method, args }),
      redirect: 'follow',
    });
    const raw = await upstream.text();
    let payload;
    try {
      payload = JSON.parse(raw);
    } catch (_) {
      throw new Error('استجابة Google Apps Script ليست JSON. تأكد من نشر Web App والسماح بالوصول.');
    }
    if (!upstream.ok) {
      return res.status(502).json({ ok: false, error: payload.error || 'تعذر الاتصال بخدمة البيانات.' });
    }
    return res.status(200).json(payload);
  } catch (error) {
    return res.status(502).json({ ok: false, error: error.message || 'تعذر الاتصال بالخادم.' });
  }
};
