import fetch from "node-fetch";
import { DateTime } from "luxon";

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// Tillåtna rumsadresser
const allowedRooms = [
  "vastberga.mote1@hissen.se",
  "vastberga.mote2@hissen.se",
  "storakonferensrummet@hissen.se"
];

// Hämta access token från Microsoft Graph
async function getAccessToken() {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    scope: "https://graph.microsoft.com/.default",
    client_secret: CLIENT_SECRET,
    grant_type: "client_credentials",
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params,
    }
  );

  const data = await res.json();
  if (!data.access_token) {
    throw new Error("Kunde inte hämta access token: " + JSON.stringify(data));
  }
  return data.access_token;
}

// API handler för Vercel
export default async function handler(req, res) {
  // CORS (för GitHub Pages)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    const roomEmail = (req.query.room || "vastberga.mote1@hissen.se").toLowerCase();

    // Kontrollera att rummet är tillåtet
    if (!allowedRooms.includes(roomEmail)) {
      return res.status(403).json({ error: "Room not allowed" });
    }

    const accessToken = await getAccessToken();

    const now = DateTime.now().setZone("Europe/Stockholm");
    const todayStartUTC = now.startOf("day").toUTC().toISO();
    const tomorrowEndUTC = now.plus({ days: 1 }).endOf("day").toUTC().toISO();

    // Hämta möten för idag + imorgon
    const calendarRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${todayStartUTC}&enddatetime=${tomorrowEndUTC}&$orderby=start/dateTime`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const calendarData = await calendarRes.json();
    const meetings = calendarData.value || [];

    // Filtrera bort möten som redan är slut
    const upcomingMeetings = meetings.filter((m) => {
      const endLocal = DateTime.fromISO(m.end.dateTime).setZone("Europe/Stockholm");
      return endLocal >= now;
    });

    // Formatera möten (bara ämne, tider, och vem som bokat)
    const formatted = upcomingMeetings.map((m) => ({
      subject: m.subject,
      start: {
        dateTime: DateTime.fromISO(m.start.dateTime)
          .setZone("Europe/Stockholm")
          .toISO(),
      },
      end: {
        dateTime: DateTime.fromISO(m.end.dateTime)
          .setZone("Europe/Stockholm")
          .toISO(),
      },
      organizer: m.organizer, // ✅ bara organisatören, inte alla deltagare
      isOnlineMeeting: m.isOnlineMeeting,
    }));

    res.status(200).json({ meetings: formatted });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Server error", details: err.message });
  }
}
