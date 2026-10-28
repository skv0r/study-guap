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
type Repair = { object: string; text: string; closed: boolean };

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
  repairs: Repair[] = [];
  inventories: { date: string; object: string; fact: string }[] = [];
  ready = { cash: true, platform: true, hall: true, passage: true };

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
    this.tickets.push({ number, trip: tripNo, seat, passenger, price, boarded: false, refunded: false });
    this.draft = false;
    return { ok: true, text: "Билет проведён, купон отправлен на печать." };
  }

  refund(number: string): Answer {
    const ticket = this.tickets.find((item) => item.number === number && !item.refunded);
    if (!ticket) return { ok: false, text: "Открытого билета с таким номером нет." };
    if (ticket.boarded) return { ok: false, text: "Посадка уже отмечена. Возврат не проведён." };
    ticket.refunded = true;
    const trip = this.trips.find((item) => item.number === ticket.trip);
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
}

export const db = new RailwayDb();

export function read(data: FormData, key: string) {
  return text(data.get(key));
}
