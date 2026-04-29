/* global Chart */

async function loadChart() {
    const canvas = document.getElementById("tempChart");
    const r = await fetch("/chart/temperature-series?limit=60");
    const { labels, values } = await r.json();

    if (!labels.length) {
        const ctx = canvas.getContext("2d");
        if (ctx) {
            ctx.font = "16px system-ui";
            ctx.fillStyle = "#64748b";
            ctx.fillText("Нет данных в MongoDB — подождите накопления опросом /connect/temperature", 8, 40);
        }
        return;
    }

    const shortLabels = labels.map((s) => (s.length > 19 ? s.slice(11, 19) : s));

    new Chart(canvas, {
        type: "line",
        data: {
            labels: shortLabels,
            datasets: [
                {
                    label: "Температура, °C",
                    data: values,
                    borderColor: "#2563eb",
                    backgroundColor: "rgba(37, 99, 235, 0.12)",
                    borderWidth: 2,
                    fill: true,
                    tension: 0.35,
                    pointRadius: 3,
                    pointHoverRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { intersect: false, mode: "index" },
            plugins: {
                title: {
                    display: true,
                    text: "Температура по журналу (последние точки)",
                    font: { size: 16, weight: "600" }
                },
                legend: { position: "top" }
            },
            scales: {
                x: {
                    grid: { color: "rgba(0,0,0,0.06)" },
                    ticks: { maxRotation: 45, minRotation: 0 }
                },
                y: {
                    grid: { color: "rgba(0,0,0,0.08)" },
                    title: { display: true, text: "°C" }
                }
            }
        }
    });
}

loadChart().catch((e) => {
    document.body.appendChild(document.createTextNode(String(e)));
});
