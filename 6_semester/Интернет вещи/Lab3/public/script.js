const CONNECT = {
    temperature: "/connect/temperature",
    light: "/connect/light",
    lock: "/connect/lock",
};

async function fetchJson(url) {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status} для ${url}`);
    }
    return response.json();
}