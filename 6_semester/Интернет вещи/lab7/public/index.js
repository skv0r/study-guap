function $(id) {
    return document.getElementById(id);
}

async function updateTemperature() {
    const data = await fetchJson(CONNECT.temperature);
    document.getElementById("tempValue").textContent = `${data.value} ${data.unit}`;
}

async function updateLight() {
    const data = await fetchJson(CONNECT.light);
    $("lightEnabled").textContent = data.enabled ? "ВКЛ" : "ВЫКЛ";
    $("lightBrightness").textContent = String(data.brightnessPercent);
}

async function updateLock() {
    const data = await fetchJson(CONNECT.lock);
    $("lockState").textContent = data.locked ? "ЗАКРЫТ" : "ОТКРЫТ";
}

async function updateHeater() {
    const data = await fetchJson(CONNECT.heater);
    $("heaterPower").textContent = data.heaterPower === "On" ? "ВКЛ" : "ВЫКЛ";
}

async function tick() {
    await Promise.all([updateTemperature(), updateLight(), updateLock(), updateHeater()]);
}

tick();
setInterval(tick, 1000);
