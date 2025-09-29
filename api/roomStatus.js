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

    // ----- Hämta interval (idag → imorgon slut) -----
    const todayStart = DateTime.now().setZone('Europe/Stockholm').startOf('day').toUTC().toISO();
    const tomorrowEnd = DateTime.now().setZone('Europe/Stockholm').plus({ days: 1 }).endOf('day').toUTC().toISO();

    const graphUrl =
      `https://graph.microsoft.com/v1.0/users/${roomEmail}/calendarview?startdatetime=${todayStart}&enddatetime=${tomorrowEnd}&$orderby=start/dateTime`;

    const graphRes = await fetch(graphUrl, { headers: { Authorization: `Bearer ${accessToken}` } });
    const graphData = await graphRes.json();

    const meetings = graphData.value || [];

    // ----- Filtrera möten i lokalt datum -----
    const today = DateTime.now().setZone('Europe/Stockholm').toFormat('yyyy-MM-dd');
    const tomorrow = DateTime.now().setZone('Europe/Stockholm').plus({ days: 1 }).toFormat('yyyy-MM-dd');

    const todaysMeetings = meetings.filter(m =>
      DateTime.fromISO(m.start.dateTime, { zone: m.start.timeZone }).setZone('Europe/Stockholm').toFormat('yyyy-MM-dd') === today
    );

    const tomorrowsMeetings = meetings.filter(m =>
