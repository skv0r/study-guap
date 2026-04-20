import express from "express";
import { runDemo } from "./things"

const app = express();
const PORT = 5000;

app.get("/", (_req, res) => {
    console.log("[GET /]")
    const {parentText, childText, sensorRecords } = runDemo();

    const html = `<!DOCTYPE html>
<html lang="ru">
<head><meta charset="utf-8"><title>ЛР2</title></head>
<body>
  <h1>Умный дом</h1>
  <h2>Родитель</h2><pre>${escapeHtml(parentText)}</pre>
  <h2>Ребёнок</h2><pre>${escapeHtml(childText)}</pre>
  <h2>Store (sensor-1)</h2><pre>${escapeHtml(sensorRecords.join("\n") || "(пусто)")}</pre>
</body>
</html>`;

    res.type("html").send(html);
});

function escapeHtml(s: string): string {
    return s.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
}


app.listen(PORT, () => {
    console.log(`http://127.0.0.1:${PORT}/`)
});