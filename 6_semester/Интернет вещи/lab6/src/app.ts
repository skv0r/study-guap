import express from "express";
import path from "node:path";
import { TemperatureSensor, LightController, SmartLock, HeaterController } from "./things";

const app = express();
const PORT = Number(process.env.PORT) || 5006;

const temp = new TemperatureSensor("sensor-1", "Температурный датчик", 22.5);
const light = new LightController("light-1", "Свет в гостиной");
const lock = new SmartLock("lock-1", "Замок входной двери");
const heater = new HeaterController("heater-1", "Обогреватель", 25);

light.setLight(true, 65);
lock.setLocked(true);

function syncHeaterFromSensor(): void {
    heater.autoPower(temp.getTemperature());
}

app.use((req, _res, next) => {
    if (req.path.startsWith("/connect") || req.path.startsWith("/command")) {
        console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
    }
    next();
});

app.use(express.static(path.join(process.cwd(), "public")));

app.get("/connect/temperature", (_req, res) => {
    const payload = temp.connect();
    syncHeaterFromSensor();
    res.json(payload);
});

app.get("/connect/light", (_req, res) => res.json(light.connect()));
app.get("/connect/lock", (_req, res) => res.json(lock.connect()));

/** Отдельный опрос состояния обогревателя (аналог /connect_heater в пособии). */
app.get("/connect/heater", (_req, res) => res.json(heater.connect()));

app.get("/command/temperature", (req, res) => {
    const q = req.query as Record<string, string | undefined>;
    const result = temp.applyCommand(q);
    if (!result.ok) {
        return res.status(400).json(result);
    }
    syncHeaterFromSensor();
    return res.json({ ok: true, state: temp.snapshot(), heater: heater.snapshot() });
});

app.get("/command/light", (req, res) => {
    const q = req.query as Record<string, string | undefined>;
    const result = light.applyCommand(q);
    if (!result.ok) {
        return res.status(400).json(result);
    }
    return res.json({ ok: true, state: light.snapshot() });
});

app.get("/command/lock", (req, res) => {
    const q = req.query as Record<string, string | undefined>;
    const result = lock.applyCommand(q);
    if (!result.ok) {
        return res.status(400).json(result);
    }
    return res.json({ ok: true, state: lock.snapshot() });
});

app.listen(PORT, () => {
    console.log(`http://127.0.0.1:${PORT}/ — ЛР6 (обогреватель, /connect/heater)`);
});
