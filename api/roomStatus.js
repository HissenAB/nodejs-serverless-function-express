// roomStatus.js

async function fetchRoomData(roomEmail) {
    try {
        const res = await fetch(`https://VERCEL_APP_URL.vercel.app/api/roomStatus?room=${encodeURIComponent(roomEmail)}`);
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return await res.json();
    } catch (err) {
        console.error("Fel vid hämtning av möten:", err);
        return { meetings: [] };
    }
}

async function updateStatus() {
    const urlParams = new URLSearchParams(window.location.search);
    const roomEmail = urlParams.get("room") || "vastberga.mote1@hissen.se"; // default

    const statusEl = document.getElementById("status");
    const todayEl = document.getElementById("todayMeetings");
    const tomorrowEl = document.getElementById("tomorrowMeetings");
    const headerEl = document.querySelector("header");

    // Rubrik beroende på rum
    headerEl.textContent =
        roomEmail.includes("mote1") ? "Mötesrum – Västberga 1" : "Mötesrum – Västberga 2";

    try {
        const data = await fetchRoomData(roomEmail);
        const now = new Date();
        const meetings = data.meetings || [];

        // Pågående möte
        const ongoing = meetings.find(
            m => new Date(m.start.dateTime) <= now && now <= new Date(m.end.dateTime)
        );
        statusEl.textContent = ongoing ? "Upptaget" : "Ledigt";
        statusEl.className = "status " + (ongoing ? "upptaget" : "ledigt");

        // Dela upp i idag och imorgon
        const todayMeetingsArr = meetings.filter(m => {
            const start = new Date(m.start.dateTime);
            return start.toDateString() === now.toDateString() && start > now;
        });

        const tomorrow = new Date(now);
        tomorrow.setDate(tomorrow.getDate() + 1);
        const tomorrowMeetingsArr = meetings.filter(m => {
            const start = new Date(m.start.dateTime);
            return start.toDateString() === tomorrow.toDateString();
        });

        // Rendera idag
        todayEl.innerHTML = "";
        if (ongoing) {
            const div = document.createElement("div");
            div.className = "meeting-card ongoing";
            div.innerHTML = `
                <div class="meeting-time">Pågående: ${new Date(
                    ongoing.start.dateTime
                ).toLocaleTimeString("sv-SE", { hour: "2-digit", minute: "2-digit" })} - 
                ${new Date(ongoing.end.dateTime).toLocaleTimeString("sv-SE", { hour: "2-digit", minute: "2-digit" })}</div>
                <div>${ongoing.subject}</div>
                <div class="attendees">Bokad av: ${ongoing.organizer?.emailAddress?.name || "Okänd"}</div>
            `;
            todayEl.appendChild(div);
        }

        todayMeetingsArr.forEach(m => {
            const div = document.createElement("div");
            div.className = "meeting-card";
            div.innerHTML = `
                <div class="meeting-time">${new Date(m.start.dateTime).toLocaleTimeString("sv-SE", { hour: "2-digit", minute: "2-digit" })} - 
                ${new Date(m.end.dateTime).toLocaleTimeString("sv-SE", { hour: "2-digit", minute: "2-digit" })}</div>
                <div>${m.subject}</div>
                <div class="attendees">Bokad av: ${m.organizer?.emailAddress?.name || "Okänd"}</div>
            `;
            todayEl.appendChild(div);
        });

        // Rendera imorgon
        tomorrowEl.innerHTML = "";
        tomorrowMeetingsArr.forEach(m => {
            const div = document.createElement("div");
            div.className = "meeting-card";
            div.innerHTML = `
                <div class="meeting-time">${new Date(m.start.dateTime).toLocaleTimeString("sv-SE", { hour: "2-digit", minute: "2-digit" })} - 
                ${new Date(m.end.dateTime).toLocaleTimeString("sv-SE", { hour: "2-digit", minute: "2-digit" })}</div>
                <div>${m.subject}</div>
                <div class="attendees">Bokad av: ${m.organizer?.emailAddress?.name || "Okänd"}</div>
            `;
            tomorrowEl.appendChild(div);
        });
    } catch (err) {
        console.error("Fel vid renderingen:", err);
        statusEl.textContent = "Fel vid laddning";
        todayEl.innerHTML = "";
        tomorrowEl.innerHTML = "";
    }
}

// Starta uppdatering direkt + varje minut
updateStatus();
setInterval(updateStatus, 60000);
