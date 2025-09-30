import { DateTime } from "luxon";

async function getAccessToken() {
  const res = await fetch("https://login.microsoftonline.com/YOUR_TENANT_ID/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
  });

  const data = await res.json();
  return data.access_token;
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    const { room } = req.query;
    if (!room || !["vastberga.mote1@hissen.se", "vastberga.mote2@hissen.se"].includes(room)) {
      return res.status(400).json({ error: "Invalid room. Use vastberga.mote1@hissen.se or vastberga.mote2@hissen.se" });
    }

    const accessToken = await getAccessToken();

    const now = DateTime.now().setZone("Europe/Stockholm");
    const todayStartUTC = now.startOf("day").toUTC().toISO();
    const tomorrowEndUTC = now.plus({ days: 1 }).endOf("day").toUTC().toISO();

    const calendarRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${room}/calendarview?startdatetime=${todayStartUTC}&enddatetime=${tomorrowEndUTC}&$orderby=start/dateTime`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const calendarData = await calendarRes.json();
    const meetings = calendarData.value || [];

    const upcomingMeetings = meetings.filter(m => {
      const endLocal = DateTime.fromISO(m.end.dateTime).setZone("Europe/Stockholm");
      return endLocal >= now;
    });

    const formatted = upcomingMeetings.map(m => ({
      subject: m.subject || "Bokat möte",
      start: { dateTime: DateTime.fromISO(m.start.dateTime).setZone("Europe/Stockholm").toISO() },
      end: { dateTime: DateTime.fromISO(m.end.dateTime).setZone("Europe/Stockholm").toISO() },
      organizer: m.organizer?.emailAddress?.name || "Okänd bokare",
      isOnlineMeeting: m.isOnlineMeeting
    }));

    res.status(200).json({ meetings: formatted });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Server error", details: err.message });
  }
}
