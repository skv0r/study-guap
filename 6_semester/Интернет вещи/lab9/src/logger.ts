import { MongoClient, type Db } from "mongodb";

const DEFAULT_URI = process.env.MONGODB_URI ?? "mongodb://127.0.0.1:27017";

/** ЛР7–8: логирование и выборка для анализа (пособие). */
export class IoTLogger {
    private client: MongoClient | null = null;
    private db: Db | null = null;
    private lastLoggedTemperature: number | null = null;
    private enabled = false;

    async init(dbName = "iot_logger_db"): Promise<void> {
        try {
            this.client = new MongoClient(DEFAULT_URI, {
                serverSelectionTimeoutMS: 4000,
                connectTimeoutMS: 4000
            });
            await this.client.connect();
            this.db = this.client.db(dbName);
            this.enabled = true;
            console.log(`[IoTLogger] MongoDB: ${DEFAULT_URI}, БД «${dbName}»`);
        } catch (err) {
            console.warn("[IoTLogger] подключение к MongoDB не удалось, логирование отключено:", err);
            this.enabled = false;
            this.db = null;
            if (this.client) {
                try {
                    await this.client.close();
                } catch {
                    /* ignore */
                }
                this.client = null;
            }
        }
    }

    isEnabled(): boolean {
        return this.enabled && this.db !== null;
    }

    async insertTemperature(value: number): Promise<void> {
        if (!this.db || !this.enabled) return;
        if (this.lastLoggedTemperature === value) {
            console.log("[IoTLogger] температура не изменилась — запись не дублируется");
            return;
        }
        this.lastLoggedTemperature = value;
        await this.db.collection("Temperature").insertOne({
            timeStamp: new Date().toISOString(),
            Temperature: value
        });
    }

    async insertHeaterSnapshot(power: string): Promise<void> {
        if (!this.db || !this.enabled) return;
        await this.db.collection("Heater").insertOne({
            timeStamp: new Date().toISOString(),
            power
        });
    }

    /** ЛР8: среднее и максимум по коллекции Temperature. */
    async getTemperatureMeanAndMax(): Promise<{ mean: number | null; max: number | null; count: number }> {
        if (!this.db || !this.enabled) {
            return { mean: null, max: null, count: 0 };
        }
        const docs = await this.db
            .collection("Temperature")
            .find({})
            .project({ _id: 0, Temperature: 1 })
            .toArray();
        const vals = docs.map((d) => Number(d.Temperature)).filter((n) => !Number.isNaN(n));
        if (vals.length === 0) {
            return { mean: null, max: null, count: 0 };
        }
        const sum = vals.reduce((a, b) => a + b, 0);
        return { mean: sum / vals.length, max: Math.max(...vals), count: vals.length };
    }

    /** ЛР8: вторая характеристика — сколько раз в логе обогреватель был On / Off. */
    async getHeaterOnOffCounts(): Promise<{ on: number; off: number }> {
        if (!this.db || !this.enabled) {
            return { on: 0, off: 0 };
        }
        const on = await this.db.collection("Heater").countDocuments({ power: "On" });
        const off = await this.db.collection("Heater").countDocuments({ power: "Off" });
        return { on, off };
    }

    /** ЛР9: ряды для графика (метки времени + температура). */
    async getTemperatureChartSeries(limit: number): Promise<{ labels: string[]; values: number[] }> {
        if (!this.db || !this.enabled) {
            return { labels: [], values: [] };
        }
        const lim = Math.min(Math.max(limit, 1), 500);
        const docs = await this.db
            .collection("Temperature")
            .find({})
            .sort({ _id: -1 })
            .limit(lim)
            .toArray();
        docs.reverse();
        return {
            labels: docs.map((d) => String(d.timeStamp ?? "")),
            values: docs.map((d) => Number(d.Temperature)).filter((n) => !Number.isNaN(n))
        };
    }

    async close(): Promise<void> {
        await this.client?.close();
    }
}
