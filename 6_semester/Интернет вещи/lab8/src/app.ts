import express from "express";
import path from "node:path";
import { TemperatureSensor, LightController, SmartLock, HeaterController } from "./things";
import { IoTLogger } from "./logger";

const app = express();
const PORT = Number(process.env.PORT) || 5008;

const temp = new TemperatureSensor("sensor-1", "Температурный датчик", 22.5);
const light = new LightController("light-1", "Свет в гостиной");
const lock = new SmartLock("lock-1", "Замок входной двери");
const heater = new HeaterController("heater-1", "Обогреватель", 25);
const logger = new IoTLogger();

light.setLight(true, 65);
lock.setLocked(true);

function syncHeaterFromSensor(): void {
    const before = heater.getPower();
    heater.autoPower(temp.getTemperature());
    const after = heater.getPower();
    void logger.insertTemperature(temp.getTemperature());
    if (before !== after) {
        void logger.insertHeaterSnapshot(after);
    }
}

app.use((req, _res, next) => {
    if (
        req.path.startsWith("/connect") ||
        req.path.startsWith("/command") ||
        req.path.startsWith("/analysis")
    ) {
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

/** ЛР8: агрегаты по данным в MongoDB. */
app.get("/analysis/stats", async (_req, res) => {
    const temperature = await logger.getTemperatureMeanAndMax();
    const heaterLog = await logger.getHeaterOnOffCounts();
    res.json({
        loggerEnabled: logger.isEnabled(),
        temperature,
        heaterSnapshots: heaterLog
    });
});

void (async function main() {
    app.listen(PORT, () => {
        console.log(`http://127.0.0.1:${PORT}/ — ЛР8 (анализ: GET /analysis/stats)`);
    });
    await logger.init();
    if (!logger.isEnabled()) {
        console.warn("MongoDB недоступна — анализ и журнал будут пустыми, UI доступен");
    }
})().catch((e) => {
    console.error(e);
    process.exit(1);
});
