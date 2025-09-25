// api/roomStatus.js
import fetch from 'node-fetch';
import { DateTime } from 'luxon'; // För tidszonshantering

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
    const roomEmail = 'motesrumtest@hissen.se'; // <-- kontrollera att detta är rätt adress

    // ----- Hämta access token -----
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

    // ----- Hämta möten för idag och imorgon -----
    const today = DateTime.now().setZone('Europe/Stockholm').startOf('day');
    const tomorrow = today.plus({ days: 1 }).endOf('day');

    // Viktigt: Graph kräver UTC i querystring
    const graphUrl =
      `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?` +
      `startdatetime=${today.toUTC().toISO()}&enddatetime=${tomorrow.toUTC().toISO()}&$orderby=start/dateTime`;

    const graphRes = await fetch(graphUrl, { headers: { Authorization: `Bearer ${accessToken}` } });
    const graphData = await graphRes.json();

    if (!graphRes.ok) {
      return res.status(500).json({ error: 'Graph error', details: graphData });
    }

    const meetings = graphData.value || [];

    // ----- Funktion för att acceptera möte -----
    async function acceptMeeting(eventId) {
      const url = `https://graph.microsoft.com/v1.0/users/${roomEmail}/events/${eventId}/accept`;
      const resp = await fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ sendResponse: true })
      });

      if (!resp.ok) {
        const err = await resp.text();
        console.error(`Kunde inte acceptera mötet (${eventId}):`, err);
      } else {
        console.log(`✅ Accepterade möte: ${eventId}`);
      }
    }

    // ----- Kolla möten och acceptera om ledigt -----
    const acceptedMeetings = meetings.filter(m => m.responseStatus?.response === "accepted");

    for (const m of meetings) {
      if (m.responseStatus?.response === "accepted") continue; // redan accepterat

      const start = DateTime.fromISO(m.start.dateTime);
      const end = DateTime.fromISO(m.end.dateTime);

      // Kontrollera konflikter
      const conflict = acceptedMeetings.some(other =>
        DateTime.fromISO(other.start.dateTime) < end &&
        DateTime.fromISO(other.end.dateTime) > start
      );

      if (!conflict) {
        await acceptMeeting(m.id);
        acceptedMeetings.push(m);
      }
    }

    // ----- Returnera möten i samma format som tidigare -----
    const formattedMeetings = meetings.map(m => {
      const startLocal = DateTime.fromISO(m.start.dateTime, { zone: m.start.timeZone })
        .setZone('Europe/Stockholm')
        .toISO();
      const endLocal = DateTime.fromISO(m.end.dateTime, { zone: m.end.timeZone })
        .setZone('Europe/Stockholm')
        .toISO();

      return {
        subject: m.subject,
        start: { dateTime: startLocal },
        end: { dateTime: endLocal },
        attendees: (m.attendees || []).filter(a => a.emailAddress.name !== 'Mötesrum test'),
        isOnlineMeeting: m.isOnlineMeeting,
      };
    });

    res.status(200).json({ meetings: formattedMeetings });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error', details: err.message });
  }
}
