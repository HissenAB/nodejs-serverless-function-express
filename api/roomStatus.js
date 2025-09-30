// roomStatus.js

async function getRoomStatus(roomEmail) {
    try {
        const res = await fetch(
            `https://nodejs-serverless-function-express-beta-dusky.vercel.app/api/roomStatus?room=${encodeURIComponent(roomEmail)}`
        );
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

    // Ändra rubriken beroende på rummet
    if (roomEmail.includes("mote1")) {
        headerEl.textContent = "Mötesrum – Västberga 1";
    } else if (roomEmail.includes("mote2")) {
        headerEl.textContent = "Mötesrum – Västberga 2";
    }

    try {
        const data = await getRoomStatus(roomEmail);
        const now = new Date();
        const meetings = data.meetings || [];

        // Nuvarande möte
        const ongoing = meetings.find(
            m => new Date(m.start.dateTime) <= now && now <= new Date(m.end.dateTime)
        );
        statusEl.textContent = ongoing ? "Upptaget" : "Ledigt";
        statusEl.className = "status " + (ongoing ? "upptaget" : "ledigt");

        // Idag
        const todayMeetingsArr = meetings.filter(m => {
            const start = new Date(m.start.dateTime);
            return start.toDateString() === now.toDateString() && start > now;
        });

        // Imorgon
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
                ${new Date(ongoing.end.dateTime).toLocaleTimeString("sv-SE", {
                    hour: "2-digit",
                    minute: "2-digit",
                })}</div>
                <div>${ongoing.subject}</div>
            `;
            if (ongoing.organizer && ongoing.organizer.emailAddress) {
                div.innerHTML += `<div class="attendees">Bokad av: ${ongoing.organizer.emailAddress.name}</div>`;
            }
            todayEl.appendChild(div);
        }

        todayMeetingsArr.forEach(m => {
            const div = document.createElement("div");
            div.className = "meeting-card";
            div.innerHTML = `
                <div class="meeting-time">${new Date(m.start.dateTime).toLocaleTimeString(
                    "sv-SE",
                    { hour: "2-digit", minute: "2-digit" }
                )} - 
                ${new Date(m.end.dateTime).toLocaleTimeString("sv-SE", {
                    hour: "2-digit",
                    minute: "2-digit",
                })}</div>
                <div>${m.subject}</div>
            `;
            if (m.organizer && m.organizer.emailAddress) {
                div.innerHTML += `<div class="attendees">Bokad av: ${m.organizer.emailAddress.name}</div>`;
            }
            todayEl.appendChild(div);
        });

        // Rendera imorgon
        tomorrowEl.innerHTML = "";
        tomorrowMeetingsArr.forEach(m => {
            const div = document.createElement("div");
            div.className = "meeting-card";
            div.innerHTML = `
                <div class="meeting-time">${new Date(m.start.dateTime).toLocaleTimeString(
                    "sv-SE",
                    { hour: "2-digit", minute: "2-digit" }
                )} - 
                ${new Date(m.end.dateTime).toLocaleTimeString("sv-SE", {
                    hour: "2-digit",
                    minute: "2-digit",
                })}</div>
                <div>${m.subject}</div>
            `;
            if (m.organizer && m.organizer.emailAddress) {
                div.innerHTML += `<div class="attendees">Bokad av: ${m.organizer.emailAddress.name}</div>`;
            }
            tomorrowEl.appendChild(div);
        });
    } catch (err) {
        console.error("Fel vid renderingen:", err);
        statusEl.textContent = "Fel vid laddning";
        todayEl.innerHTML = "";
        tomorrowEl.innerHTML = "";
    }
}

updateStatus();
setInterval(updateStatus, 60000);
