abstract class Thing {
  constructor(
    public readonly id: string,
    public readonly name: string,
    public isOnline: boolean = true
  ) {}

  abstract getStatus(): string;
}

class TemperatureSensor extends Thing {
  private currentTemperatureC: number;

  constructor(id: string, name: string, initialTemperatureC: number) {
    super(id, name, true);
    this.currentTemperatureC = initialTemperatureC;
  }

  setTemperature(valueC: number): void {
    this.currentTemperatureC = valueC;
  }

  getTemperature(): number {
    return this.currentTemperatureC;
  }

  getStatus(): string {
    return `${this.name}: ${this.currentTemperatureC.toFixed(1)} °C`;
  }
}

class LightController extends Thing {
  private brightnessPercent = 0;
  private isEnabled = false;

  setLight(isEnabled: boolean, brightnessPercent: number): void {
    this.isEnabled = isEnabled;
    this.brightnessPercent = Math.max(0, Math.min(100, brightnessPercent));
  }

  getStatus(): string {
    return `${this.name}: ${this.isEnabled ? "ВКЛ" : "ВЫКЛ"}, яркость=${this.brightnessPercent}%`;
  }
}

class SmartLock extends Thing {
  private locked = true;

  setLocked(value: boolean): void {
    this.locked = value;
  }

  getStatus(): string {
    return `${this.name}: ${this.locked ? "ЗАКРЫТ" : "ОТКРЫТ"}`;
  }
}

class DeviceDataStore {
  private readonly dataByThingId = new Map<string, string[]>();

  saveRecord(thing: Thing, payload: string): void {
    const records = this.dataByThingId.get(thing.id) ?? [];
    records.push(payload);
    this.dataByThingId.set(thing.id, records);
  }

  getRecords(thingId: string): string[] {
    return this.dataByThingId.get(thingId) ?? [];
  }
}

class MainControlUnit {
  private readonly things: Thing[] = [];

  constructor(private readonly dataStore: DeviceDataStore) {}

  registerThing(thing: Thing): void {
    this.things.push(thing);
  }

  collectMonitoringData(): void {
    for (const thing of this.things) {
      // Сохраняем снимок текущего состояния каждого устройства с временной меткой.
      const line = `[${new Date().toISOString()}] ${thing.getStatus()}`;
      this.dataStore.saveRecord(thing, line);
    }
  }

  getAllStatuses(): string[] {
    return this.things.map((thing) => thing.getStatus());
  }
}

class ParentInterface {
  render(statuses: string[]): string {
    return [
      "Интерфейс родителя (администратора)",
      "Возможности: полный обзор и управление",
      ...statuses
    ].join("\n");
  }
}

class ChildInterface {
  render(statuses: string[]): string {
    const safeLines = statuses.filter(
      (line) => line.includes("Температурный датчик") || line.includes("Свет")
    );
    return [
      "Интерфейс ребенка (ограниченный)",
      "Возможности: только просмотр температуры и света",
      ...safeLines
    ].join("\n");
  }
}

function demoVariant1SmartHome(): void {
  const db = new DeviceDataStore();
  const mcu = new MainControlUnit(db);

  const temperature = new TemperatureSensor("sensor-1", "Температурный датчик", 22.5);
  const light = new LightController("light-1", "Свет в гостиной");
  const lock = new SmartLock("lock-1", "Замок входной двери");

  light.setLight(true, 65);
  lock.setLocked(true);

  mcu.registerThing(temperature);
  mcu.registerThing(light);
  mcu.registerThing(lock);
  mcu.collectMonitoringData();

  const statuses = mcu.getAllStatuses();
  const parentUi = new ParentInterface();
  const childUi = new ChildInterface();

  console.log(parentUi.render(statuses));
  console.log("-----");
  console.log(childUi.render(statuses));
  console.log("-----");
  console.log("Сохраненные записи для sensor-1:", db.getRecords("sensor-1"));
}

demoVariant1SmartHome();
