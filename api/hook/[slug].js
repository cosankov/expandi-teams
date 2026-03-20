const { buildTeamsPayload } = require('../../lib/format');

module.exports = async function handler(req, res) {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // Decode the Teams webhook URL from the path
  const { slug } = req.query;
  if (!slug) {
    return res.status(400).json({ error: 'Missing webhook slug.' });
  }

  let teamsUrl;
  try {
    teamsUrl = Buffer.from(slug, 'base64url').toString('utf-8');
  } catch {
    return res.status(400).json({ error: 'Invalid webhook slug.' });
  }

  // Validate it's a Teams webhook URL (legacy connectors or new Workflows)
  const isLegacy = teamsUrl.includes('.webhook.office.com/');
  const isWorkflow = teamsUrl.includes('.logic.azure.com/');
  if (!isLegacy && !isWorkflow) {
    return res.status(400).json({ error: 'Invalid Microsoft Teams webhook URL.' });
  }

  // GET = health check (Expandi test may send GET)
  if (req.method === 'GET') {
    return res.status(200).json({ ok: true, status: 'ready' });
  }

  // Parse the Expandi payload
  const payload = req.body;
  if (!payload || !payload.hook) {
    return res.status(200).json({ ok: true, status: 'received', note: 'No hook object — test ping acknowledged.' });
  }

  // Format the Teams message
  const teamsPayload = buildTeamsPayload(payload);

  // POST to Teams
  try {
    const response = await fetch(teamsUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(teamsPayload),
    });

    if (!response.ok) {
      const body = await response.text();
      console.error('Teams API error:', response.status, body);
      return res.status(502).json({
        error: 'Teams rejected the message.',
        status: response.status,
        detail: body,
      });
    }

    return res.status(200).json({
      ok: true,
      event: payload.hook.event,
      contact: [payload.contact?.first_name, payload.contact?.last_name].filter(Boolean).join(' ') || null,
    });
  } catch (err) {
    console.error('Failed to POST to Teams:', err.message);
    return res.status(502).json({ error: 'Failed to reach Teams.', detail: err.message });
  }
};
