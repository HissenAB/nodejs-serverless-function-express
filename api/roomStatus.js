import fetch from 'node-fetch';
import { DateTime } from 'luxon';

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// Lista med tillåtna rum och deras displaynamn
const allowedRooms = {
  "vastberga.mote1@hissen.se": "Mötesrum 1 - Västberga",
  "vastberga.mote2@hissen.se": "Mötesrum 2 - Västberga",
  "storakonferensrummet@hissen.se": "Stora konferensrummet"
};

// Hämta access token från Microsoft Graph
async function getAccessToken() {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: CLIENT_SECRET,
    grant_type: 'client_credentials'
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

// API handler för Vercel
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const roomEmail = req.query.room;

  // Kontrollera om rummet är tillåtet
  if (!roomEmail || !allowedRooms[roomEmail]) {
    return res.status(403).json({ error: "Rummet är inte tillåtet", displayName: roomEmail });
  }

  try {
    const accessToken = await getAccessToken();
    const now = DateTime.now().setZone('Europe/Stockholm');
    const todayStartUTC = now.startOf('day').toUTC().toISO();
    const tomorrowEndUTC = now.plus({ days: 1 }).endOf('day').toUTC().toISO();

    // Hämta möten för idag + imorgon
    const calendarRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${todayStartUTC}&enddatetime=${tomorrowEndUTC}&$orderby=start/dateTime`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const calendarData = await calendarRes.json();
    const meetings = calendarData.value || [];

    const upcomingMeetings = meetings.filter(m => {
      const endLocal = DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm');
      return endLocal >= now;
    });

    const formatted = upcomingMeetings.map(m => ({
      subject: m.subject,
      start: { dateTime: DateTime.fromISO(m.start.dateTime).setZone('Europe/Stockholm').toISO() },
      end: { dateTime: DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm').toISO() },
      organizer: m.organizer,
      isOnlineMeeting: m.isOnlineMeeting
    }));

    res.status(200).json({ meetings: formatted, displayName: allowedRooms[roomEmail] });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error', details: err.message, displayName: allowedRooms[roomEmail] });
  }
}
