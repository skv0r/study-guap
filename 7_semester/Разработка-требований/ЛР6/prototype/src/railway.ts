import { supabase } from "./supabase";

export type Answer = { ok: boolean; text: string };

type Station = { code: string; name: string };
type Scheme = { type: string; seats: string };
type RouteRec = { code: string; name: string; stations: string; minutes: string };
type Trip = {
  number: string;
  date: string;
  route: string;
  scheme: string;
  time: string;
  open: boolean;
  cancelled: boolean;
};
type Ticket = {
  number: string;
  trip: string;
  seat: string;
  passenger: string;
  price: string;
  boarded: boolean;
  refunded: boolean;
};
type Repair = { id?: number; object: string; text: string; closed: boolean };
type Seat = { trip: string; seat: string; free: boolean };

const fail = (error: { message: string } | null): Answer | null =>
  error ? { ok: false, text: `База не записала строку: ${error.message}` } : null;

const money = (value: string) => {
  const number = Number(value.replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(number) ? number : 0;
};

const showMoney = (value: number | string) =>
  new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 2 }).format(Number(value));

export const expandSeats = (spec: string) => {
  const seats: string[] = [];
  for (const part of spec.split(/[,;]/).map((item) => item.trim()).filter(Boolean)) {
    const range = part.split(/[–-]/).map((item) => item.trim());
    if (range.length === 2 && /^\d+$/.test(range[0]) && /^\d+$/.test(range[1])) {
      for (let number = Number(range[0]); number <= Number(range[1]); number += 1) seats.push(String(number));
    } else seats.push(part);
  }
  return seats;
};

const text = (value: FormDataEntryValue | null) => String(value ?? "").trim();

export class RailwayDb {
  stationOpen = true;
  role = "";
  shiftOpen = false;
  draft = false;
  stations: Station[] = [
    { code: "SVR", name: "Северная" },
    { code: "YUG", name: "Южная" },
  ];
  schemes: Scheme[] = [{ type: "Плацкарт", seats: "1–36" }];
  routes: RouteRec[] = [
    { code: "M-12", name: "Северная — Южная", stations: "Северная — Южная", minutes: "95" },
  ];
  trips: Trip[] = [
    {
      number: "102",
      date: "2026-10-20",
      route: "M-12 Северная — Южная",
      scheme: "Плацкарт",
      time: "08:40–10:15",
      open: true,
      cancelled: false,
    },
  ];
  tickets: Ticket[] = [];
  seats: Seat[] = [];
  repairs: Repair[] = [];
  inventories: { id?: number; date: string; object: string; fact: string }[] = [];
  ready = { cash: true, platform: true, hall: true, passage: true };
  serviceDay = "2026-10-20";
  busy = 0;

  private needsShift(role: string) {
    return role === "Кассир" || role === "Контролёр";
  }

  enter(role: string): Answer {
    this.role = role;
    if (this.needsShift(role)) {
      if (!this.stationOpen) {
        return { ok: false, text: "Станция не в обслуживании. Смену открыть нельзя." };
      }
      this.shiftOpen = true;
      return { ok: true, text: "Смена открыта." };
    }
    return { ok: true, text: "Вход выполнен." };
  }

  leave() {
    this.shiftOpen = false;
    this.role = "";
  }

  saveStation(code: string, name: string): Answer {
    const key = code.toUpperCase();
    if (!key || !name) return { ok: false, text: "Код и название станции обязательны." };
    if (this.stations.some((item) => item.code === key)) {
      return { ok: false, text: `Код ${key} уже есть. Вторая станция не записана.` };
    }
    this.stations.push({ code: key, name });
    return { ok: true, text: `Станция ${name} сохранена. Код ${key} уникален.` };
  }

  saveScheme(type: string, seats: string): Answer {
    if (!type || !seats) return { ok: false, text: "Тип вагона и номера мест обязательны." };
    const numbers = seats.split(/[,\s–-]+/).filter(Boolean);
    if (new Set(numbers).size !== numbers.length) {
      return { ok: false, text: "Номера мест повторяются. Схема не сохранена." };
    }
    const found = this.schemes.find((item) => item.type === type);
    if (found) found.seats = seats;
    else this.schemes.push({ type, seats });
    return { ok: true, text: "Схема сохранена. Номера мест не повторяются." };
  }

  saveRoute(code: string, name: string, stations: string, minutes: string): Answer {
    const parts = stations.split("—").map((item) => item.trim()).filter(Boolean);
    if (parts.length < 2) return { ok: false, text: "В маршруте нужны две станции." };
    if (new Set(parts).size !== parts.length) {
      return { ok: false, text: "Станция в маршруте повторяется. Маршрут не сохранён." };
    }
    if (this.routes.some((item) => item.code === code)) {
      return { ok: false, text: `Маршрут ${code} уже есть.` };
    }
    this.routes.push({ code, name, stations, minutes });
    return { ok: true, text: "Маршрут сохранён: две станции, без повтора." };
  }

  publishTrip(number: string, date: string, route: string, scheme: string, time: string): Answer {
    const found = this.trips.find((item) => item.number === number);
    if (found?.cancelled) return { ok: false, text: "Рейс отменён. Тем же номером не открывается." };
    if (found?.open) return { ok: false, text: "Рейс с этим номером уже открыт." };
    if (found) {
      found.open = true;
      found.date = date;
      found.route = route;
      found.scheme = scheme;
      found.time = time;
    } else {
      this.trips.push({ number, date, route, scheme, time, open: true, cancelled: false });
    }
    const plan = this.schemes.find((item) => item.type === scheme);
    for (const seat of expandSeats(plan?.seats ?? "")) {
      if (!this.seats.some((item) => item.trip === number && item.seat === seat)) {
        this.seats.push({ trip: number, seat, free: true });
      }
    }
    return { ok: true, text: `Рейс ${number} опубликован. Фонд мест создан.` };
  }

  cancelTrip(number: string): Answer {
    const found = this.trips.find((item) => item.number === number);
    if (!found) return { ok: false, text: "Такого рейса нет." };
    if (found.cancelled) return { ok: false, text: "Рейс уже отменён. Повторная отмена не создана." };
    const hadDraft = this.draft;
    found.cancelled = true;
    found.open = false;
    this.draft = false;
    return {
      ok: true,
      text: hadDraft
        ? "Рейс отменён. Черновик продажи не проведён, новые продажи закрыты."
        : "Рейс отменён. Новые продажи закрыты, билеты помечены.",
    };
  }

  sell(tripNo: string, seat: string, passenger: string, price: string, draft: boolean): Answer {
    if (!this.stationOpen || !this.shiftOpen) {
      return { ok: false, text: "Нет обслуживания станции или открытой смены." };
    }
    const trip = this.trips.find((item) => item.number === tripNo);
    if (!trip?.open) return { ok: false, text: "Рейс закрыт. Билет не создан." };
    if (draft) {
      this.draft = true;
      return { ok: true, text: "Черновик продажи записан. Билет ещё не проведён." };
    }
    if (this.tickets.some((item) => item.trip === tripNo && item.seat === seat && !item.refunded)) {
      return { ok: false, text: `Место ${seat} рейса ${tripNo} уже продано. Второй билет не создан.` };
    }
    const number = `${tripNo}-${seat}`;
    const previous = this.tickets.find((item) => item.number === number);
    if (previous) {
      previous.passenger = passenger;
      previous.price = price;
      previous.boarded = false;
      previous.refunded = false;
    } else {
      this.tickets.push({ number, trip: tripNo, seat, passenger, price, boarded: false, refunded: false });
    }
    const place = this.seats.find((item) => item.trip === tripNo && item.seat === seat);
    if (place) place.free = false;
    this.draft = false;
    return { ok: true, text: "Билет проведён, купон отправлен на печать." };
  }

  refund(number: string): Answer {
    const ticket = this.tickets.find((item) => item.number === number && !item.refunded);
    if (!ticket) return { ok: false, text: "Открытого билета с таким номером нет." };
    if (ticket.boarded) return { ok: false, text: "Посадка уже отмечена. Возврат не проведён." };
    ticket.refunded = true;
    const trip = this.trips.find((item) => item.number === ticket.trip);
    const place = this.seats.find((item) => item.trip === ticket.trip && item.seat === ticket.seat);
    if (place && trip?.open) place.free = true;
    return {
      ok: true,
      text: trip?.open ? "Билет возвращён, место свободно." : "Билет погашен. Место закрыто: рейс отменён.",
    };
  }

  board(number: string): Answer {
    if (!this.shiftOpen || !this.stationOpen) {
      return { ok: false, text: "Нет смены или станция не в обслуживании." };
    }
    const ticket = this.tickets.find((item) => item.number === number && !item.refunded);
    if (!ticket) return { ok: false, text: "Билет для посадки не найден." };
    if (ticket.boarded) return { ok: false, text: "Отметка посадки уже есть." };
    ticket.boarded = true;
    return { ok: true, text: "Посадка отмечена один раз." };
  }

  saveService(ready: RailwayDb["ready"], note: string): Answer {
    this.ready = ready;
    if (!Object.values(ready).every(Boolean)) {
      const objectName = note && note !== "Ограничений нет" ? note : "Объект станции";
      if (!this.repairs.some((item) => item.object === objectName && !item.closed)) {
        this.repairs.push({ object: objectName, text: note || "Не готов", closed: false });
      }
      return { ok: true, text: "Неготовый объект: заявка на ремонт создана или обновлена." };
    }
    if (this.shiftOpen && this.stationOpen) {
      return { ok: false, text: "Есть открытая смена. Сутки не закрыты." };
    }
    this.stationOpen = !this.stationOpen;
    return {
      ok: true,
      text: this.stationOpen
        ? "Сутки открыты. Станция в обслуживании."
        : "Сутки закрыты. Новые смены недоступны.",
    };
  }

  saveRepair(objectName: string, description: string, status: string, role: string): Answer {
    const same = this.repairs.find((item) => item.object === objectName && item.text === description);
    if (same?.closed) {
      return { ok: false, text: "Заявка уже закрыта. Той же неисправностью повторно не открывается." };
    }
    if (status === "Закрыта" && role !== "Инженер по ремонту") {
      return { ok: false, text: "Закрыть заявку может только инженер после отметки устранения." };
    }
    if (status === "Закрыта") {
      if (same) same.closed = true;
      else this.repairs.push({ object: objectName, text: description, closed: true });
      return { ok: true, text: "Заявка закрыта." };
    }
    if (!same) this.repairs.push({ object: objectName, text: description, closed: false });
    return { ok: true, text: "Заявка на ремонт открыта или обновлена." };
  }

  saveInventory(date: string, objectName: string, fact: string): Answer {
    if (this.inventories.some((item) => item.date === date && item.object === objectName)) {
      return { ok: false, text: `Инвентаризация ${objectName} за ${date} уже есть. Повтор закрыт.` };
    }
    this.inventories.push({ date, object: objectName, fact });
    return { ok: true, text: `Инвентаризация за ${date} проведена. Повтор на эту дату закрыт.` };
  }

  closeShift(): Answer {
    if (this.draft) return { ok: false, text: "Есть черновик продажи. Смена не закрыта." };
    this.shiftOpen = false;
    return { ok: true, text: "Смена закрыта." };
  }

  async pull(force = false): Promise<Answer> {
    if (!force && this.busy) return { ok: true, text: "" };
    const stations = await supabase.from("Станция").select("*").order("код");
    const schemes = await supabase.from("Схема вагона").select("*").order("тип");
    const routes = await supabase.from("Маршрут").select("*").order("код");
    const stops = await supabase.from("Состав маршрута").select("*").order("порядок");
    const trips = await supabase.from("Рейс").select("*").order("номер");
    const seats = await supabase.from("Место").select("*");
    const tickets = await supabase.from("Билет").select("*");
    const repairs = await supabase.from("Заявка на ремонт").select("*").order("ид");
    const inventories = await supabase.from("Инвентаризация").select("*").order("дата");
    const shift = await supabase.from("Смена").select("*").eq("код", "текущая").maybeSingle();
    const service = await supabase.from("Обслуживание станции").select("*").order("сутки", { ascending: false }).limit(1).maybeSingle();
    const error = [stations, schemes, routes, stops, trips, seats, tickets, repairs, inventories, shift, service].map((item) => item.error).find(Boolean);
    const failed = fail(error ?? null);
    if (failed) return failed;

    const stationName = new Map((stations.data ?? []).map((row) => [row.код, row.название]));
    const routeName = new Map((routes.data ?? []).map((row) => [row.код, row.название]));
    this.stations = (stations.data ?? []).map((row) => ({ code: row.код, name: row.название }));
    this.schemes = (schemes.data ?? []).map((row) => ({ type: row.тип, seats: row["номера мест"] }));
    this.routes = (routes.data ?? []).map((row) => ({
      code: row.код,
      name: row.название,
      stations: (stops.data ?? [])
        .filter((stop) => stop.маршрут === row.код)
        .sort((left, right) => left.порядок - right.порядок)
        .map((stop) => stationName.get(stop.станция) ?? stop.станция)
        .join(" — "),
      minutes: String(row["время хода"]),
    }));
    this.trips = (trips.data ?? []).map((row) => ({
      number: row.номер,
      date: String(row.дата).slice(0, 10),
      route: `${row.маршрут} ${routeName.get(row.маршрут) ?? ""}`.trim(),
      scheme: row.схема,
      time: row.время,
      open: row.открыт,
      cancelled: row["отменён"],
    }));
    this.seats = (seats.data ?? []).map((row) => ({ trip: row.рейс, seat: row["номер места"], free: row.свободно }));
    this.tickets = (tickets.data ?? []).map((row) => ({
      number: row.номер,
      trip: row.рейс,
      seat: row.место,
      passenger: row.пассажир,
      price: showMoney(row.цена),
      boarded: row.посадка,
      refunded: row.возврат,
    }));
    this.repairs = (repairs.data ?? []).map((row) => ({
      id: row.ид,
      object: row.объект,
      text: row.описание,
      closed: row.закрыта,
    }));
    this.inventories = (inventories.data ?? []).map((row) => ({
      id: row.ид,
      date: String(row.дата).slice(0, 10),
      object: row.объект,
      fact: row.факт,
    }));
    this.role = shift.data?.роль ?? "";
    this.shiftOpen = Boolean(shift.data?.открыта);
    this.draft = Boolean(shift.data?.черновик);
    if (service.data) {
      this.serviceDay = String(service.data.сутки).slice(0, 10);
      this.stationOpen = service.data.открыты;
      this.ready = {
        cash: service.data.кассы,
        platform: service.data.платформы,
        hall: service.data["зал ожидания"],
        passage: service.data.проходы,
      };
    }
    return { ok: true, text: "" };
  }

  async flush(): Promise<Answer> {
    const stations = await supabase.from("Станция").upsert(
      this.stations.map((item) => ({ код: item.code, название: item.name })),
      { onConflict: "код" },
    );
    const schemes = await supabase.from("Схема вагона").upsert(
      this.schemes.map((item) => ({ тип: item.type, "номера мест": item.seats })),
      { onConflict: "тип" },
    );
    const routes = await supabase.from("Маршрут").upsert(
      this.routes.map((item) => ({
        код: item.code,
        название: item.name,
        "время хода": Number(item.minutes) || 1,
      })),
      { onConflict: "код" },
    );
    const stationCode = new Map(this.stations.map((item) => [item.name, item.code]));
    const stops = this.routes.flatMap((route) =>
      route.stations
        .split("—")
        .map((item) => item.trim())
        .filter(Boolean)
        .map((name, index) => ({
          маршрут: route.code,
          станция: stationCode.get(name),
          порядок: index + 1,
        }))
        .filter((item): item is { маршрут: string; станция: string; порядок: number } => Boolean(item.станция)),
    );
    const clearStops = await supabase.from("Состав маршрута").delete().not("маршрут", "is", null);
    const writeStops = stops.length
      ? await supabase.from("Состав маршрута").insert(stops)
      : { error: null };
    const trips = await supabase.from("Рейс").upsert(
      this.trips.map((item) => ({
        номер: item.number,
        дата: item.date,
        маршрут: item.route.split(/\s+/)[0],
        схема: item.scheme,
        время: item.time,
        открыт: item.open,
        "отменён": item.cancelled,
      })),
      { onConflict: "номер" },
    );
    const seats = this.seats.length
      ? await supabase.from("Место").upsert(
          this.seats.map((item) => ({ рейс: item.trip, "номер места": item.seat, свободно: item.free })),
          { onConflict: "рейс,номер места" },
        )
      : { error: null };
    const tickets = this.tickets.length
      ? await supabase.from("Билет").upsert(
          this.tickets.map((item) => ({
            номер: item.number,
            рейс: item.trip,
            место: item.seat,
            пассажир: item.passenger,
            цена: money(item.price),
            посадка: item.boarded,
            возврат: item.refunded,
          })),
          { onConflict: "номер" },
        )
      : { error: null };
    const shift = await supabase.from("Смена").upsert(
      {
        код: "текущая",
        роль: this.role,
        сотрудник: this.role || "не назначен",
        открыта: this.shiftOpen,
        черновик: this.draft,
      },
      { onConflict: "код" },
    );
    const service = await supabase.from("Обслуживание станции").upsert(
      {
        сутки: this.serviceDay,
        кассы: this.ready.cash,
        платформы: this.ready.platform,
        "зал ожидания": this.ready.hall,
        проходы: this.ready.passage,
        открыты: this.stationOpen,
      },
      { onConflict: "сутки" },
    );
    const repairWrites = [];
    for (const item of this.repairs) {
      repairWrites.push(
        item.id
          ? supabase.from("Заявка на ремонт").update({ объект: item.object, описание: item.text, закрыта: item.closed }).eq("ид", item.id)
          : supabase.from("Заявка на ремонт").insert({ объект: item.object, описание: item.text, закрыта: item.closed }).select("*").single(),
      );
    }
    const repairResults = await Promise.all(repairWrites);
    repairResults.forEach((result, index) => {
      const id = result.data && "ид" in result.data ? result.data.ид : undefined;
      if (id) this.repairs[index].id = id;
    });
    const inventory = this.inventories.length
      ? await supabase.from("Инвентаризация").upsert(
          this.inventories.map((item) => ({ дата: item.date, объект: item.object, факт: item.fact })),
          { onConflict: "дата,объект" },
        )
      : { error: null };
    const error = [stations, schemes, routes, clearStops, writeStops, trips, seats, tickets, shift, service, inventory, ...repairResults]
      .map((item) => item.error)
      .find(Boolean);
    return fail(error ?? null) ?? { ok: true, text: "" };
  }

  async apply(task: () => Answer): Promise<Answer> {
    this.busy += 1;
    try {
      const loaded = await this.pull(true);
      if (!loaded.ok) return loaded;
      const answer = task();
      if (!answer.ok) {
        await this.pull(true);
        return answer;
      }
      const saved = await this.flush();
      if (!saved.ok) {
        await this.pull(true);
        return saved;
      }
      return answer;
    } finally {
      this.busy -= 1;
    }
  }
}

export const db = new RailwayDb();

export function read(data: FormData, key: string) {
  return text(data.get(key));
}
