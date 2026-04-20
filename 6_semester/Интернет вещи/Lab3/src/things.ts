export abstract class Thing {
    constructor(
        public id: string,
        public name: string,
        public isOnline: boolean = true
    ) {}

    protected abstract emulate(): void
    protected abstract buildPayload(): Record<string, unknown>;
    
    connect(): Record<string, unknown> {
        console.log(`[Thing.connect] ${this.constructor.name} (${this.id})`);
        this.emulate();
        return this.buildPayload();
    }

    abstract getStatus(): string;
}

export interface IParent {
    render(statuses: string[]): string;
}

export interface IChild {
    render(statuses: string[]): string;
}

export class MainControlUnit {
    constructor(
        public readonly things: Thing[],
        public readonly dataStore: DeviceDataStore
    ) {}

    registerThing(thing: Thing): void {
        console.log(`[MainControlUnit.registerThing] ${thing.id}`);
        this.things.push(thing);
    }
    collectMonitoringData(): void {
        console.log(`[MainControlUnit.collectMonitoringData]` );
        for (const thing of this.things) {
            const line = `[${new Date().toISOString()}] ${thing.getStatus()}`
            this.dataStore.saveRecord(thing, line)
        }
    }
    getAllStatuses(): string[] {
        console.log(`[MainControlUnit.getAllStatuses]`);
        return this.things.map((t) => t.getStatus());
    }
}

export class DeviceDataStore {
    private readonly dataByThingId = new Map<string, string[]>();

    saveRecord(thing: Thing, payload: string): void {
        console.log(`[DeviceDataStore.saveRecord] ${thing.id}` );
        const list = this.dataByThingId.get(thing.id) ?? [];
        list.push(payload);
        this.dataByThingId.set(thing.id, list);
    }

    getRecords(thingId: string): string[] {
        console.log(`[DeviceDataStore.getRecords] ${thingId}` );
        return this.dataByThingId.get(thingId) ?? [];
    }
    
}

export class TemperatureSensor extends Thing {
    private currentTemperatureC: number;

    constructor(id: string, name: string, initialTemperatureC: number) {
        super(id, name, true);
        this.currentTemperatureC = initialTemperatureC;
    }

    setTemperature(valueC: number): void {
        console.log(`[TemperatureSensor.setTemperature] ${this.id}`);
        this.currentTemperatureC = valueC;
    }

    getTemperature(): number {
        return this.currentTemperatureC;
    }

    protected emulate(): void {
        const delta = (Math.random() - 0.5) * 1.2;
        this.currentTemperatureC = Math.round((this.currentTemperatureC + delta) * 10) / 10;
    }

    protected buildPayload(): Record<string, unknown> {
        return {
          id: this.id,
          name: this.name,
          value: this.currentTemperatureC,
          unit: "°C"
        };
    }
    

    getStatus(): string {
        console.log(`[TemperatureSensor.getStatus] ${this.id}`);
        return `${this.name}: ${this.currentTemperatureC.toFixed(1)} °C`
    }
}

export class LightController extends Thing {
    private brightnessPercent: number = 0;
    private isEnabled: boolean = false;

    setLight(isEnabled: boolean, brightnessPercent: number): void {
        console.log(`[LightController.setLight] ${this.id}`);
        this.isEnabled = isEnabled;
        this.brightnessPercent = Math.max(0, Math.min(100, brightnessPercent))
    }

    protected emulate(): void {
        const delta = Math.floor((Math.random() - 0.5) * 12);
        this.brightnessPercent = Math.max(0, Math.min(100, this.brightnessPercent + delta));
    }

    protected buildPayload(): Record<string, unknown> {
        return {
          id: this.id,
          name: this.name,
          enabled: this.isEnabled,
          brightnessPercent: this.brightnessPercent
        };
    }

    getStatus(): string {
        console.log(`[LightController.getStatus] ${this.id}`);
        return `${this.name}: ${this.isEnabled ? "ВКЛ" : "ВЫКЛ"}, яркость=${this.brightnessPercent}%`;
    }
}

export class SmartLock extends Thing{
    private isLocked: boolean = true;

    setLocked(value: boolean): void {
        console.log(`[SmartLock.setLocked] ${this.id}`);
         this.isLocked = value;
    }

    protected emulate(): void {
        if (Math.random() < 0.08) this.isLocked = !this.isLocked;
      }
    
    protected buildPayload(): Record<string, unknown> {
        return {
          id: this.id,
          name: this.name,
          locked: this.isLocked
        };
    }

    getStatus(): string {
        console.log(`[SmartLock.getStatus] ${this.id}`);
        return `${this.name}: ${this.isLocked ? "ЗАКРЫТ" : "ОТКРЫТ"}`;
    }
}

export class ParentInterface implements IParent {
    render(statuses: string[]): string {
        console.log("[ParentInterface.render]");
        return ["Интерфейс родителя (полный доступ)", ...statuses].join("\n");
    }
}

export class ChildInterface implements IChild {
    render(statuses: string[]): string {
        console.log("[ChildInterface.render]");
        return ["Интерфейс ребенка (ограничен)", ...statuses].join("\n");
    }
}

export function runDemo(): {
    parentText: string;
    childText: string;
    sensorRecords: string[];
  } {
    const store = new DeviceDataStore();
    const things: Thing[] = [];
    const mcu = new MainControlUnit(things, store);
    const temp = new TemperatureSensor("sensor-1", "Температурный датчик", 22.5);
    const light = new LightController("light-1", "Свет в гостиной");
    const lock = new SmartLock("lock-1", "Замок входной двери");
    light.setLight(true, 65);
    lock.setLocked(true);
    mcu.registerThing(temp);
    mcu.registerThing(light);
    mcu.registerThing(lock);
    mcu.collectMonitoringData();
    const statuses = mcu.getAllStatuses();
    const parentText = new ParentInterface().render(statuses);
    const childText = new ChildInterface().render(statuses);
    return {
      parentText,
      childText,
      sensorRecords: store.getRecords("sensor-1")
    };
  }