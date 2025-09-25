// api/roomStatus.js
import fetch from 'node-fetch';

export default async function handler(req, res) {
  // Tillåt alla origins
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const roomEmail = 'motesrumtest@hissen.se';

    // Hämta access token
    const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        scope: 'https://graph.microsoft.com/.default',
        client_secret: clientSecret,
        grant_type: 'client_credentials'
      })
    });
    const tokenData = await tokenResponse.json();
    if (!tokenData.access_token) {
      return res.status(500).json({ error: 'Failed to get access token', details: tokenData });
    }
    const accessToken = tokenData.access_token;

    const now = new Date().toISOString();
    const end = new Date(Date.now() + 1 * 24 * 60 * 60 * 1000).toISOString();

    // Hämta upp till 10 kommande möten
    const graphUrl = `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${now}&enddatetime=${end}&$orderby=start/dateTime&$top=10`;
    const graphResponse = await fetch(graphUrl, { headers: { Authorization: `Bearer ${accessToken}` } });
    const graphData = await graphResponse.json();
    const meetings = graphData.value || [];

    // Pågående möte
    const currentMeeting = meetings.find(m => new Date(m.start.dateTime) <= new Date() && new Date() <= new Date(m.end.dateTime));

    // Nästa möte (börjar efter nu)
    const nextMeeting = meetings.find(m => new Date(m.start.dateTime) > new Date());

    // Filtrera bort "Mötesrum test" som deltagare
    const filterAttendees = m => {
      if (!m) return [];
      return (m.attendees || []).filter(a => a.emailAddress.name !== 'Mötesrum test');
    };

    res.status(200).json({ currentMeeting: currentMeeting ? { ...currentMeeting, attendees: filterAttendees(currentMeeting) } : null,
                            nextMeeting: nextMeeting ? { ...nextMeeting, attendees: filterAttendees(nextMeeting) } : null });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error', details: err.message });
  }
}
