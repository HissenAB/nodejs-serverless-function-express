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

    // 1️⃣ Hämta access token
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

    // 2️⃣ Hämta senaste 50 mejlen från inboxen
    const mailRes = await fetch(`https://graph.microsoft.com/v1.0/users/${roomEmail}/mailFolders/Inbox/messages?$top=50`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    const mailData = await mailRes.json();
    const messages = mailData.value || [];

    // 3️⃣ Filtrera mötesinbjudningar i JavaScript
    const meetingRequests = messages.filter(m => m.meetingMessageType === 'meetingRequest');
    console.log(`Hämtade ${meetingRequests.length} mötesinbjudningar från inboxen`);

    // 4️⃣ Acceptera varje möte direkt via messageId
    for (const msg of meetingRequests) {
      try {
        console.log(`Accepterar möte: ${msg.subject}`);
        const acceptRes = await fetch(`https://graph.microsoft.com/v1.0/users/${roomEmail}/messages/${msg.id}/accept`, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ sendResponse: true })
        });

        if (!acceptRes.ok) {
          const err = await acceptRes.text();
          console.error(`❌ Kunde inte acceptera mötet (${msg.subject}):`, err);
        } else {
          console.log(`✅ Accepterade mötet: ${msg.subject}`);
        }
      } catch (err) {
        console.error(`Fel vid accept av möte (${msg.subject}):`, err);
      }
    }

    // 5️⃣ Hämta kalendern för idag och imorgon
    const now = DateTime.now().setZone('Europe/Stockholm');
    const todayStartUTC = now.startOf('day').toUTC().toISO();
    const tomorrowEndUTC = now.plus({ days: 1 }).endOf('day').toUTC().toISO();

    const calendarRes = await fetch(`https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${todayStartUTC}&enddatetime=${tomorrowEndUTC}&$orderby=start/dateTime`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    const calendarData = await calendarRes.json();
    const meetings = calendarData.value || [];

    // 6️⃣ Filtrera bort möten som redan passerat
    const upcomingMeetings = meetings.filter(m => {
      const endLocal = DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm');
      return endLocal >= now;
    });

    // 7️⃣ Formatera möten
    const formattedMeetings = upcomingMeetings.map(m => ({
      subject: m.subject,
      start: { dateTime: DateTime.fromISO(m.start.dateTime).setZone('Europe/Stockholm').toISO() },
      end: { dateTime: DateTime.fromISO(m.end.dateTime).setZone('Europe/Stockholm').toISO() },
      attendees: (m.attendees || []).filter(a => a.emailAddress.name !== 'Mötesrum test'),
      isOnlineMeeting: m.isOnlineMeeting
    }));

    res.status(200).json({ meetings: formattedMeetings });

  } catch (err) {
    console.error('Server error:', err);
    res.status(500).json({ error: 'Server error', details: err.message });
  }
}
