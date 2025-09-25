import fetch from 'node-fetch';
import { DateTime } from 'luxon';

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const REFRESH_TOKEN = process.env.AZURE_REFRESH_TOKEN; // Refresh token från användaren
const ROOM_EMAIL = 'vastberga.mote@hissen.se';

async function getAccessToken() {
  // Byt refresh token mot nytt access token
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default offline_access',
    grant_type: 'refresh_token',
    refresh_token: REFRESH_TOKEN,
    client_secret: CLIENT_SECRET
  });

  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params
  });

  const data = await res.json();
  if (!data.access_token) throw new Error('Could not get access token: ' + JSON.stringify(data));
  return data.access_token;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const accessToken = await getAccessToken();

    // 1️⃣ Hämta kalender för idag + imorgon
    const now = DateTime.now().setZone('Europe/Stockholm');
    const todayStartUTC = now.startOf('day').toUTC().toISO();
    const tomorrowEndUTC = now.plus({ days: 1 }).endOf('day').toUTC().toISO();

    const calendarRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/calendarview?startdatetime=${todayStartUTC}&enddatetime=${tomorrowEndUTC}&$orderby=start/dateTime`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const calendarData = await calendarRes.json();
    const meetings = calendarData.value || [];

    // 2️⃣ Filtrera bort passerade möten
    const upcomingMeetings = meetings.filter(m => {
      const endLocal = DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm');
      return endLocal >= now;
    });

    // 3️⃣ Hämta redan accepterade möten
    const acceptedMeetings = upcomingMeetings.filter(m => m.responseStatus?.response === 'accepted');

    // 4️⃣ Auto-acceptera möten som inte krockar
    for (const m of upcomingMeetings) {
      if (m.responseStatus?.response === 'accepted') continue;

      const start = DateTime.fromISO(m.start.dateTime);
      const end = DateTime.fromISO(m.end.dateTime);

      const conflict = acceptedMeetings.some(other =>
        DateTime.fromISO(other.start.dateTime) < end &&
        DateTime.fromISO(other.end.dateTime) > start
      );

      if (!conflict) {
        const acceptRes = await fetch(`https://graph.microsoft.com/v1.0/me/events/${m.id}/accept`, {
          method: 'POST',
          headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ sendResponse: true })
        });

        if (!acceptRes.ok) {
          const err = await acceptRes.text();
          console.error(`❌ Kunde inte acceptera mötet (${m.subject}):`, err);
        } else {
          console.log(`✅ Accepterade mötet: ${m.subject}`);
          acceptedMeetings.push(m);
        }
      }
    }

    // 5️⃣ Formatera möten för front-end
    const formatted = upcomingMeetings.map(m => ({
      subject: m.subject,
      start: { dateTime: DateTime.fromISO(m.start.dateTime).setZone('Europe/Stockholm').toISO() },
      end: { dateTime: DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm').toISO() },
      attendees: (m.attendees || []).filter(a => a.emailAddress.name !== 'Mötesrum test'),
      isOnlineMeeting: m.isOnlineMeeting
    }));

    res.status(200).json({ meetings: formatted });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error', details: err.message });
  }
}
