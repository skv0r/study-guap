async function sendCommand(url) {
    const response = await fetch(url);
    const data = await response.json();
    if (!response.ok) {
        throw new Error(JSON.stringify(data));
    }
    return data;
}

function setResult(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

document.getElementById("btnTemp").addEventListener("click", async () => {
    const value = document.getElementById("cmdTemp").value.trim();
    try {
        const data = await sendCommand(
            `/command/temperature?value=${encodeURIComponent(value)}`
        );
        setResult("cmdTempResult", JSON.stringify(data));
    } catch (e) {
        setResult("cmdTempResult", String(e.message));
    }
});

document.getElementById("btnLight").addEventListener("click", async () => {
    const enabled = document.getElementById("cmdLightOn").checked;
    const brightness = document.getElementById("cmdLightBright").value;
    const qs = new URLSearchParams({
        enabled: String(enabled),
        brightness: String(brightness)
    });
    try {
        const data = await sendCommand(`/command/light?${qs.toString()}`);
        setResult("cmdLightResult", JSON.stringify(data));
    } catch (e) {
        setResult("cmdLightResult", String(e.message));
    }
});

document.getElementById("btnLock").addEventListener("click", async () => {
    const locked = document.getElementById("cmdLockLocked").checked;
    const qs = new URLSearchParams({ locked: String(locked) });
    try {
        const data = await sendCommand(`/command/lock?${qs.toString()}`);
        setResult("cmdLockResult", JSON.stringify(data));
    } catch (e) {
        setResult("cmdLockResult", String(e.message));
    }
});
