import { MongoClient, type Db } from "mongodb";

const DEFAULT_URI = process.env.MONGODB_URI ?? "mongodb://127.0.0.1:27017";

/**
 * ЛР7: логирование в MongoDB, два метода записи в разные коллекции (пособие — класс Logger).
 */
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

    /** Коллекция Temperature: без дублей подряд одинакового значения (как в пособии). */
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

    /** Вторая коллекция — снимки состояния обогревателя. */
    async insertHeaterSnapshot(power: string): Promise<void> {
        if (!this.db || !this.enabled) return;
        await this.db.collection("Heater").insertOne({
            timeStamp: new Date().toISOString(),
            power
        });
    }

    async close(): Promise<void> {
        await this.client?.close();
    }
}
