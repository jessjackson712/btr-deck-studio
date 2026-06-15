export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const webhookUrl = process.env.SLACK_WEBHOOK_URL;
  if (!webhookUrl) return res.status(500).json({ error: 'SLACK_WEBHOOK_URL not configured' });

  const {
    clientName, clientType, deckTypeLabel, deckTypeIcon,
    requestedBy, dueDate, specialInstructions,
    reportLabels = [], transcriptCount = 0, requestId
  } = req.body;

  const appUrl = `https://btrdecks.app`;
  const dueLine = dueDate
    ? `Due ${new Date(dueDate).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} · `
    : '';
  const requestorLine = requestedBy ? `Requested by ${requestedBy}` : 'No requestor specified';

  const reportSection = reportLabels.length > 0
    ? reportLabels.map(l => `• ${l}`).join('\n')
    : '_No reports attached_';

  const blocks = [
    {
      type: 'header',
      text: { type: 'plain_text', text: `${deckTypeIcon} New Deck Request — ${clientName}`, emoji: true }
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `*${deckTypeLabel}* · ${dueLine}${requestorLine}\n*Client type:* ${clientType === 'prospect' ? 'Prospect' : 'Active Client'}`
      }
    },
    { type: 'divider' },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `*📁 Reports (${reportLabels.length}):*\n${reportSection}${transcriptCount > 0 ? `\n\n*📝 Transcripts:* ${transcriptCount} included` : ''}`
      }
    },
  ];

  if (specialInstructions) {
    blocks.push({
      type: 'section',
      text: { type: 'mrkdwn', text: `*📌 Special Instructions:*\n${specialInstructions}` }
    });
  }

  blocks.push(
    { type: 'divider' },
    {
      type: 'actions',
      elements: [{
        type: 'button',
        text: { type: 'plain_text', text: '→ Open in BTR Deck Studio', emoji: true },
        url: appUrl,
        style: 'primary'
      }]
    },
    {
      type: 'context',
      elements: [{ type: 'mrkdwn', text: `Request ID: ${requestId}` }]
    }
  );

  try {
    const slackRes = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ blocks })
    });
    if (!slackRes.ok) throw new Error(`Slack error: ${slackRes.status}`);
    res.status(200).json({ ok: true });
  } catch (err) {
    console.error('Slack notify error:', err);
    res.status(500).json({ error: err.message });
  }
}
