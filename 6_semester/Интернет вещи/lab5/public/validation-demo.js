function showBad(url) {
    const out = document.getElementById("badDemoOut");
    fetch(url)
        .then(async (r) => {
            const text = await r.text();
            try {
                const j = JSON.parse(text);
                out.textContent = `HTTP ${r.status}\n${JSON.stringify(j, null, 2)}`;
            } catch {
                out.textContent = `HTTP ${r.status}\n${text}`;
            }
        })
        .catch((e) => {
            out.textContent = String(e);
        });
}

document.getElementById("badTemp").addEventListener("click", () => {
    showBad("/command/temperature?value=abc");
});

document.getElementById("badLight").addEventListener("click", () => {
    showBad("/command/light?enabled=maybe&brightness=50");
});

document.getElementById("badLock").addEventListener("click", () => {
    showBad("/command/lock?locked=скорее_нет");
});
