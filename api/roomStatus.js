// api/roomStatus.js
import fetch from 'node-fetch';
import { DateTime } from 'luxon';

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
    const roomEmail = 'vastberga.mote@hissen.se'; // <-- ändrat här

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
    if (!tokenData.access_token) {
      console.error('Failed to get token:', tokenData);
      return res.status(500).json({ error: 'Failed to get token', details: tokenData });
    }
    const accessToken = tokenData.access_token;

    // ----- Hämta möten för idag + imorgon -----
    const todayStart = DateTime.now().setZone('Europe/Stockholm').startOf('day').toUTC().toISO();
    const tomorrowEnd = DateTime.now().setZone('Europe/Stockholm').plus({ days: 1 }).endOf('day').toUTC().toISO();

    const graphUrl = `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${todayStart}&enddatetime=${tomorrowEnd}&$orderby=start/dateTime`;
    const graphRes = await fetch(graphUrl, { headers: { Authorization: `Bearer ${accessToken}` } });
    const graphData = await graphRes.json();

    if (!graphRes.ok) {
      console.error('Graph API error:', graphData);
      return res.status(500).json({ error: 'Graph error', details: graphData });
    }

    const meetings = graphData.value || [];
    console.log(`Hämtade ${meetings.length} möten från kalendern`);

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
        console.error(`❌ Kunde inte acceptera mötet (${eventId}):`, err);
      } else {
        console.log(`✅ Accepterade möte: ${eventId}`);
      }
    }

    // ----- Acceptera alla möten som inte redan accepterats och inte krockar -----
    const acceptedMeetings = meetings.filter(m => m.responseStatus?.response === "accepted");

    for (const m of meetings) {
      if (m.responseStatus?.response === "accepted") continue;

      const start = DateTime.fromISO(m.start.dateTime);
      const end = DateTime.fromISO(m.end.dateTime);

      const conflict = acceptedMeetings.some(other =>
        DateTime.fromISO(other.start.dateTime) < end &&
        DateTime.fromISO(other.end.dateTime) > start
      );

      if (!conflict) {
        console.log(`Försöker acceptera möte: ${m.subject} (${m.id})`);
        await acceptMeeting(m.id);
        acceptedMeetings.push(m);
      }
    }

    // ----- Filtrera bort möten som redan passerat idag -----
    const now = DateTime.now().setZone('Europe/Stockholm');
    const upcomingMeetings = meetings.filter(m => {
      const endLocal = DateTime.fromISO(m.end.dateTime, { zone: m.end.timeZone }).setZone('Europe/Stockholm');
      return endLocal >= now;
    });

    // ----- Formatera möten -----
    const formattedMeetings = upcomingMeetings.map(m => {
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

    res.status(200).json({ meetings: formattedMeetings });

  } catch (err) {
    console.error('Server error:', err);
    res.status(500).json({ error: 'Server error', details: err.message });
  }
}
