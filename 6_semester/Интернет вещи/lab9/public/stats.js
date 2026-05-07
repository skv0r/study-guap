document.getElementById("btnStats").addEventListener("click", async () => {
    const el = document.getElementById("statsOut");
    try {
        const r = await fetch("/analysis/stats");
        const data = await r.json();
        el.textContent = JSON.stringify(data, null, 2);
    } catch (e) {
        el.textContent = String(e);
    }
});
