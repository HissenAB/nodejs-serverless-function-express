// api/roomStatus.js
import fetch from 'node-fetch';

export default async function handler(req, res) {
  // Tillåt alla origins
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  // Hantera preflight request
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    console.log('Börjar köra roomStatus...');
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const roomEmail = 'motesrumtest@hissen.se';

    console.log('Tenant:', tenantId);
    console.log('ClientId:', clientId);
    console.log('ClientSecret satt:', !!clientSecret);

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

    // Hämta nästa möte
    const now = new Date().toISOString();
    const end = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
    const graphUrl = `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${now}&enddatetime=${end}&$orderby=start/dateTime&$top=1`;

    const graphResponse = await fetch(graphUrl, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    const graphData = await graphResponse.json();
    let nextMeeting = graphData.value && graphData.value.length > 0 ? graphData.value[0] : null;

    if (nextMeeting) {
      // Ta bort mötesrummet som deltagare
      nextMeeting.attendees = (nextMeeting.attendees || []).filter(
        a => a.emailAddress.address.toLowerCase() !== 'motesrumtest@hissen.se'
      );
      // Ta bort platsen
      nextMeeting.location = null;
    }

    res.status(200).json({ nextMeeting });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error', details: err.message });
  }
}
