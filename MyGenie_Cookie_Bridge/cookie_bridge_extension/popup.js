const DOMAIN   = "mygenieplus-ir1.onbmc.com";
const RECEIVER = "http://localhost:17731/cookies";

function setStatus(text, color) {
  document.getElementById("statusText").textContent = text;
  const dot = document.getElementById("statusDot");
  dot.className = "status-dot " + (color || "");
}

async function sendCookies() {
  const btn = document.getElementById("sendBtn");
  btn.disabled = true;
  btn.textContent = "⏳ Sending…";
  setStatus("Reading cookies from Edge…", "yellow");

  try {
    // Get all cookies for the MyGenie domain
    const cookies = await chrome.cookies.getAll({ domain: DOMAIN });

    if (!cookies || cookies.length === 0) {
      setStatus("No cookies found. Are you logged in to MyGenie?", "red");
      btn.disabled = false;
      btn.textContent = "🔗 Send Session to Report App";
      return;
    }

    // Build a simple key:value object
    const cookieMap = {};
    cookies.forEach(c => { cookieMap[c.name] = c.value; });

    // POST to the local Python receiver
    const response = await fetch(RECEIVER, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ cookies: cookieMap, domain: DOMAIN }),
    });

    if (response.ok) {
      const data = await response.json();
      setStatus(`✅ Done! Sent ${Object.keys(cookieMap).length} cookie(s).`, "green");
      btn.textContent = "✅ Sent Successfully";

      // Auto-close popup after 2s
      setTimeout(() => window.close(), 2000);
    } else {
      setStatus("Report app not running? Start it first, then retry.", "red");
      btn.disabled = false;
      btn.textContent = "🔗 Send Session to Report App";
    }

  } catch (err) {
    if (err.message && err.message.includes("fetch")) {
      setStatus("Cannot reach app. Is the report app running?", "red");
    } else {
      setStatus("Error: " + err.message, "red");
    }
    btn.disabled = false;
    btn.textContent = "🔗 Send Session to Report App";
  }
}

// Auto-check on popup open: ping the receiver to see if app is running
window.addEventListener("load", async () => {
  try {
    const r = await fetch("http://localhost:17731/ping", { method: "GET" });
    if (r.ok) {
      setStatus("Report app detected — ready to send.", "green");
    } else {
      setStatus("Start the report app first, then click Send.", "yellow");
    }
  } catch {
    setStatus("Start the report app first, then click Send.", "yellow");
  }
});
