// api/roomStatus.js
import fetch from 'node-fetch';
import { DateTime } from 'luxon';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const roomEmail = 'vastberga.mote@hissen.se';

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
    if (!tokenData.access_token) {
      console.error('Failed to get token:', tokenData);
      return res.status(500).json({ error: 'Failed to get token', details: tokenData });
    }
    const accessToken = tokenData.access_token;

    // Hämta möten för idag + imorgon
    const now = DateTime.now().setZone('Europe/Stockholm');
    const todayStartUTC = now.startOf('day').toUTC().toISO();
    const tomorrowEndUTC = now.plus({ days: 1 }).endOf('day').toUTC().toISO();

    const graphUrl = `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${todayStartUTC}&enddatetime=${tomorrowEndUTC}&$orderby=start/dateTime`;
    const graphRes = await fetch(graphUrl, { headers: { Authorization: `Bearer ${accessToken}` } });
    const graphData = await graphRes.json();

    if (!graphRes.ok) {
      console.error('Graph API error:', graphData);
      return res.status(500).json({ error: 'Graph error', details: graphData });
    }

    const meetings = graphData.value || [];
    console.log(`Hämtade möten: ${meetings.length}`);

    // Funktion för att acceptera möte
    async function acceptMeeting(eventId, subject) {
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
        console.error(`❌ Kunde inte acceptera mötet (${subject}):`, err);
      } else {
        console.log(`✅ Accepterade möte: ${subject}`);
      }
    }

    // Lista över accepterade möten för att kolla konflikter
    const acceptedMeetings = meetings.filter(m => m.responseStatus?.response === "accepted");

    for (const m of meetings) {
      const start = DateTime.fromISO(m.start.dateTime).setZone('Europe/Stockholm');
      const end = DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm');

      if (m.responseStatus?.response === "accepted") {
        console.log(`Möte redan accepterat: ${m.subject}`);
        continue;
      }

      // Kolla krockar med redan accepterade möten
      const conflict = acceptedMeetings.some(other => {
        const otherStart = DateTime.fromISO(other.start.dateTime).setZone('Europe/Stockholm');
        const otherEnd = DateTime.fromISO(other.end.dateTime).setZone('Europe/Stockholm');
        return start < otherEnd && end > otherStart;
      });

      if (!conflict) {
        console.log(`Försöker acceptera möte: ${m.subject}`);
        await acceptMeeting(m.id, m.subject);
        acceptedMeetings.push(m);
      } else {
        console.log(`Krock med annat möte, accepterar inte: ${m.subject}`);
      }
    }

    // Filtrera bort möten som redan är passerade idag
    const upcomingMeetings = meetings.filter(m => {
      const endLocal = DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm');
      return endLocal >= now;
    });

    // Formatera möten
    const formattedMeetings = upcomingMeetings.map(m => {
      return {
        subject: m.subject,
        start: { dateTime: DateTime.fromISO(m.start.dateTime).setZone('Europe/Stockholm').toISO() },
        end: { dateTime: DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm').toISO() },
        attendees: (m.attendees || []).filter(a => a.emailAddress.name !== 'Mötesrum test'),
        isOnlineMeeting: m.isOnlineMeeting
      };
    });

    res.status(200).json({ meetings: formattedMeetings });

  } catch (err) {
    console.error('Server error:', err);
    res.status(500).json({ error: 'Server error', details: err.message });
  }
}
