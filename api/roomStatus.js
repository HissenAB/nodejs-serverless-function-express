// api/roomStatus.js
import fetch from "node-fetch";
import { DateTime } from "luxon";

export default async function handler(req, res) {
  // Tillåt alla origins (kan ändra till specifik domän)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  try {
    console.log("👉 Kör roomStatus...");
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;
    const roomEmail = "motesrumtest@hissen.se";

    //hämta access token
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId,
          scope: "https://graph.microsoft.com/.default",
          client_secret: clientSecret,
          grant_type: "client_credentials",
        }),
      }
    );

    const tokenData = await tokenResponse.json();
    if (!tokenData.access_token) {
      return res
        .status(500)
        .json({ error: "Failed to get access token", details: tokenData });
    }

    const accessToken = tokenData.access_token;

    // hämta nästa möte
    const now = new Date().toISOString();
    const end = new Date(Date.now() + 1 * 24 * 60 * 60 * 1000).toISOString();
    const graphUrl = `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${now}&enddatetime=${end}&$orderby=start/dateTime&$top=1`;

    const graphResponse = await fetch(graphUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    const graphData = await graphResponse.json();
    let nextMeeting =
      graphData.value && graphData.value.length > 0 ? graphData.value[0] : null;

    if (nextMeeting) {
      //Konvertera start/slut till Europe/Stockholm
      const stockholmStart = DateTime.fromISO(nextMeeting.start.dateTime, {
        zone: "utc",
      }).setZone("Europe/Stockholm");
      const stockholmEnd = DateTime.fromISO(nextMeeting.end.dateTime, {
        zone: "utc",
      }).setZone("Europe/Stockholm");

      nextMeeting.start.localDateTime = stockholmStart.toISO();
      nextMeeting.end.localDateTime = stockholmEnd.toISO();

      //Filtrera bort rummet från attendees
      if (nextMeeting.attendees) {
        nextMeeting.attendees = nextMeeting.attendees.filter(
          (a) =>
            a.emailAddress?.address.toLowerCase() !==
            roomEmail.toLowerCase()
        );
      }

      // Ta bort plats
      delete nextMeeting.location;
      delete nextMeeting.locations;
    }

    res.status(200).json({ nextMeeting });
  } catch (err) {
    console.error("Fel i roomStatus:", err);
    res.status(500).json({ error: "Server error", details: err.message });
  }
}
