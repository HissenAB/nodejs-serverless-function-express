import fetch from 'node-fetch';
import { DateTime } from 'luxon';

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// Rumsmejl och visningsnamn
const allowedRooms = {
  "kajen@hissen.se": "Kajen - Västberga",
  "stadshuset@hissen.se": "Stadshuset - Västberga",
  "storakonferensrummet@hissen.se": "Stora konferensrummet",
  "centralen@hissen.se": "Centralen - Västberga",
};

// Cache: { [roomEmail]: { timestamp: Date, data: object } }
const roomCache = {};
const CACHE_TTL_MS = 60 * 1000; // 1 minut

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

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const roomEmail = req.query.room;
  if (!roomEmail || !allowedRooms[roomEmail]) {
    return res.status(403).json({ error: "Rummet är inte tillåtet", displayName: roomEmail });
  }

  try {
    // --- Kolla cache ---
    const cached = roomCache[roomEmail];
    const now = Date.now();
    if (cached && (now - cached.timestamp < CACHE_TTL_MS)) {
      return res.status(200).json(cached.data);
    }

    // --- Hämta från Graph API ---
    const accessToken = await getAccessToken();
    const todayStartUTC = DateTime.now().setZone('Europe/Stockholm').startOf('day').toUTC().toISO();
    const weekEndUTC = DateTime.now().setZone('Europe/Stockholm').plus({ days: 7 }).endOf('day').toUTC().toISO();

    const calendarRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${todayStartUTC}&enddatetime=${weekEndUTC}&$orderby=start/dateTime`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const calendarData = await calendarRes.json();
    const meetings = calendarData.value || [];

    // Filtrera bort möten som redan är slut
    const upcomingMeetings = meetings.filter(m => {
      const endLocal = DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm');
      return endLocal >= DateTime.now().setZone('Europe/Stockholm');
    });

    const formatted = upcomingMeetings.map(m => ({
      subject: m.subject,
      start: { dateTime: DateTime.fromISO(m.start.dateTime).setZone('Europe/Stockholm').toISO() },
      end: { dateTime: DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm').toISO() },
      organizer: m.organizer,
      isOnlineMeeting: m.isOnlineMeeting
    }));

    const responseData = { meetings: formatted, displayName: allowedRooms[roomEmail] };

    // --- Spara i cache ---
    roomCache[roomEmail] = { timestamp: now, data: responseData };

    res.status(200).json(responseData);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error', details: err.message, displayName: allowedRooms[roomEmail] });
  }
}
