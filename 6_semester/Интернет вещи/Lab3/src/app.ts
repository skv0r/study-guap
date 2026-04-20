import express from "express";
import path from "node:path"
import { TemperatureSensor, LightController, SmartLock } from "./things";

const app = express();
const PORT = 5000;

const temp = new TemperatureSensor("sensor-1", "Температурный датчик", 22.5);
const light = new LightController("light-1", "Свет в гостиной");
const lock = new SmartLock("lock-1", "Замок входной двери");


app.use((req, _res, next) => {
    if (req.path.startsWith("/connect")) {
        console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
    }
    next();
});

app.use(express.static(path.join(process.cwd(), "public")));
   
app.get("/connect/temperature", (_req, res) => res.json(temp.connect()))
app.get("/connect/light", (_req, res) => res.json(light.connect()))
app.get("/connect/lock", (_req, res) => res.json(lock.connect()))

app.listen(PORT, () => {
    console.log(`http://127.0.0.1:${PORT}/`)
});

