// api/roomStatus.js
import fetch from 'node-fetch';
import { DateTime } from 'luxon'; // Vi använder luxon för korrekt tidszonshantering

export default async function handler(req, res) {
  // Tillåt alla origins
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const roomEmail = 'motesrumtest@hissen.se';

    // Hämta access token
    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        scope: 'https://graph.microsoft.com/.default',
        client_secret: clientSecret,
        grant_type: 'client_credentials'
      })
    });

    const tokenData = await tokenRes.json();
    if (!tokenData.access_token) return res.status(500).json({ error: 'Failed to get token', details: tokenData });

    const accessToken = tokenData.access_token;

    const now = new Date().toISOString();
    const end = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();

    const graphUrl = `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${now}&enddatetime=${end}&$orderby=start/dateTime&$top=5`;
    const graphRes = await fetch(graphUrl, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    const graphData = await graphRes.json();

    const meetings = (graphData.value || []).map(m => {
      // Konvertera till Europe/Stockholm med luxon
      const startLocal = DateTime.fromISO(m.start.dateTime, { zone: m.start.timeZone }).setZone('Europe/Stockholm').toISO();
      const endLocal = DateTime.fromISO(m.end.dateTime, { zone: m.end.timeZone }).setZone('Europe/Stockholm').toISO();

      return {
        subject: m.subject,
        start: { dateTime: startLocal },
        end: { dateTime: endLocal },
        attendees: (m.attendees || []).filter(a => a.emailAddress.name !== 'Mötesrum test'),
        isOnlineMeeting: m.isOnlineMeeting,
      };
    });

    res.status(200).json({ meetings });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error', details: err.message });
  }
}
