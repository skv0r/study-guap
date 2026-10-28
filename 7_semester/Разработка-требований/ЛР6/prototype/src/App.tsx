import { useState, type ReactNode } from "react";
import { db, read, type Answer } from "./railway";

type Role =
  | "Диспетчер расписания"
  | "Кассир"
  | "Контролёр"
  | "Администратор справочников"
  | "Дежурный по станции"
  | "Инженер по ремонту"
  | "Инженер по инвентаризации";

type Screen =
  | "login"
  | "menu"
  | "stations"
  | "schemes"
  | "route"
  | "publish"
  | "cancel"
  | "sale"
  | "refund"
  | "boarding"
  | "service"
  | "repair"
  | "inventory"
  | "shift";

const roles: Role[] = [
  "Диспетчер расписания",
  "Кассир",
  "Контролёр",
  "Администратор справочников",
  "Дежурный по станции",
  "Инженер по ремонту",
  "Инженер по инвентаризации",
];

const menuByRole: Record<Role, { id: Screen; label: string }[]> = {
  "Диспетчер расписания": [
    { id: "route", label: "Составить маршрут" },
    { id: "publish", label: "Опубликовать рейс" },
    { id: "cancel", label: "Отменить рейс" },
  ],
  Кассир: [
    { id: "sale", label: "Продать билет" },
    { id: "refund", label: "Вернуть билет" },
    { id: "shift", label: "Закрыть смену" },
  ],
  Контролёр: [
    { id: "boarding", label: "Отметить посадку" },
    { id: "shift", label: "Закрыть смену" },
  ],
  "Администратор справочников": [
    { id: "stations", label: "Справочник станций" },
    { id: "schemes", label: "Схемы вагонов" },
  ],
  "Дежурный по станции": [
    { id: "service", label: "Обслуживание станции" },
    { id: "repair", label: "Заявка на ремонт" },
  ],
  "Инженер по ремонту": [{ id: "repair", label: "Заявка на ремонт" }],
  "Инженер по инвентаризации": [{ id: "inventory", label: "Инвентаризация" }],
};

const needsShift = (role: Role) => role === "Кассир" || role === "Контролёр";

const App = () => {
  const [screen, setScreen] = useState<Screen>("login");
  const [role, setRole] = useState<Role>("Кассир");
  const [notice, setNotice] = useState("");
  const [bad, setBad] = useState(false);
  const [tick, setTick] = useState(0);
  const [repairStatus, setRepairStatus] = useState("Открыта");

  function show(answer: Answer) {
    setNotice(answer.text);
    setBad(!answer.ok);
    setTick((value) => value + 1);
  }

  function enter() {
    const answer = db.enter(role);
    show(answer);
    if (answer.ok) setScreen("menu");
  }

  function leave() {
    db.leave();
    setScreen("login");
    setNotice("");
  }

  function back() {
    setScreen("menu");
    setNotice("");
  }

  const title =
    screen === "login"
      ? "Вход на рабочее место"
      : screen === "menu"
        ? "Меню роли"
        : (menuByRole[role].find((item) => item.id === screen)?.label ?? "Форма");

  return (
    <div className="shell">
      <header className="top">
        <div>
          <h1>АИС железной дороги</h1>
          <p>
            {role}
            {needsShift(role) ? (db.shiftOpen ? " · смена открыта" : " · смена закрыта") : ""}
            {db.stationOpen ? " · станция в обслуживании" : " · станция закрыта"}
          </p>
        </div>
        {screen !== "login" ? (
          <button type="button" onClick={leave}>
            Выйти
          </button>
        ) : null}
      </header>
      <main className="panel">
        <h2>{title}</h2>
        {notice ? <p className={bad ? "note bad" : "note"}>{notice}</p> : null}
        {screen === "login" ? (
          <form
            className="form"
            onSubmit={(event) => {
              event.preventDefault();
              enter();
            }}
          >
            <label>
              Роль
              <select value={role} onChange={(event) => setRole(event.target.value as Role)}>
                {roles.map((item) => (
                  <option key={item}>{item}</option>
                ))}
              </select>
            </label>
            <label>
              Станция
              <input name="station" defaultValue="Северная" />
            </label>
            <button className="primary" type="submit">
              {needsShift(role) ? "Войти и открыть смену" : "Войти"}
            </button>
          </form>
        ) : null}
        {screen === "menu" ? (
          <div className="menu">
            {menuByRole[role].map((item) => (
              <button key={item.id} type="button" onClick={() => { setScreen(item.id); setNotice(""); }}>
                {item.label}
              </button>
            ))}
          </div>
        ) : null}
        {screen === "stations" ? (
          <Form onBack={back} onSave={(data) => show(db.saveStation(read(data, "code"), read(data, "name")))}>
            <label>Код<input name="code" defaultValue="LUG" /></label>
            <label>Название<input name="name" defaultValue="Луговая" /></label>
            <Records title="Записанные станции" rows={db.stations.map((item) => `${item.code} — ${item.name}`)} mark={tick} />
          </Form>
        ) : null}
        {screen === "schemes" ? (
          <Form onBack={back} onSave={(data) => show(db.saveScheme(read(data, "type"), read(data, "seats")))}>
            <label>Тип вагона<input name="type" defaultValue="Купе" /></label>
            <label>Номера мест<input name="seats" defaultValue="1–18" /></label>
            <Records title="Схемы" rows={db.schemes.map((item) => `${item.type}: ${item.seats}`)} mark={tick} />
          </Form>
        ) : null}
        {screen === "route" ? (
          <Form
            onBack={back}
            onSave={(data) => show(db.saveRoute(read(data, "code"), read(data, "name"), read(data, "stations"), read(data, "minutes")))}
          >
            <label>Код<input name="code" defaultValue="M-14" /></label>
            <label>Название<input name="name" defaultValue="Северная — Луговая" /></label>
            <label>Станции по порядку<input name="stations" defaultValue="Северная — Луговая" /></label>
            <label>Время хода, мин<input name="minutes" defaultValue="40" /></label>
            <Records title="Маршруты" rows={db.routes.map((item) => `${item.code} ${item.stations}`)} mark={tick} />
          </Form>
        ) : null}
        {screen === "publish" ? (
          <Form
            onBack={back}
            onSave={(data) =>
              show(db.publishTrip(read(data, "number"), read(data, "date"), read(data, "route"), read(data, "scheme"), read(data, "time")))
            }
          >
            <label>Номер рейса<input name="number" defaultValue="102" /></label>
            <label>Дата<input name="date" type="date" defaultValue="2026-10-20" /></label>
            <label>Маршрут<input name="route" defaultValue="M-12 Северная — Южная" /></label>
            <label>Схема вагона<input name="scheme" defaultValue="Плацкарт" /></label>
            <label>Время отправления и прибытия<input name="time" defaultValue="08:40–10:15" /></label>
            <Records
              title="Рейсы"
              rows={db.trips.map((item) => `${item.number}: ${item.cancelled ? "отменён" : item.open ? "открыт" : "закрыт"}`)}
              mark={tick}
            />
          </Form>
        ) : null}
        {screen === "cancel" ? (
          <Form onBack={back} onSave={(data) => show(db.cancelTrip(read(data, "number")))}>
            <label>Рейс<input name="number" defaultValue="102" /></label>
            <label>Причина<input name="reason" defaultValue="Отмена перевозчиком" /></label>
          </Form>
        ) : null}
        {screen === "sale" ? (
          <Form
            onBack={back}
            onSave={(data) =>
              show(db.sell(read(data, "trip"), read(data, "seat"), read(data, "passenger"), read(data, "price"), data.get("draft") === "on"))
            }
          >
            <label>Рейс<input name="trip" defaultValue="102" /></label>
            <label>Станции и дата<input name="path" defaultValue="Северная — Южная, 20.10.2026" /></label>
            <label>Место<input name="seat" defaultValue="12" /></label>
            <label>Пассажир<input name="passenger" defaultValue="Иванов Иван" /></label>
            <label>Цена из прайса<input name="price" defaultValue="1 850" /></label>
            <label>
              <span>Черновик</span>
              <input name="draft" type="checkbox" />
            </label>
            <Records
              title="Проведённые билеты"
              rows={db.tickets.filter((item) => !item.refunded).map((item) => `${item.number}, место ${item.seat}, ${item.passenger}`)}
              mark={tick}
            />
          </Form>
        ) : null}
        {screen === "refund" ? (
          <Form onBack={back} onSave={(data) => show(db.refund(read(data, "number")))}>
            <label>Номер билета<input name="number" defaultValue="102-12" /></label>
          </Form>
        ) : null}
        {screen === "boarding" ? (
          <Form onBack={back} onSave={(data) => show(db.board(read(data, "number")))}>
            <label>Номер билета или купона<input name="number" defaultValue="102-12" /></label>
          </Form>
        ) : null}
        {screen === "service" ? (
          <Form
            onBack={back}
            onSave={(data) =>
              show(
                db.saveService(
                  {
                    cash: data.get("cash") === "on",
                    platform: data.get("platform") === "on",
                    hall: data.get("hall") === "on",
                    passage: data.get("passage") === "on",
                  },
                  read(data, "note"),
                ),
              )
            }
          >
            <div className="checks">
              <label><input name="cash" type="checkbox" defaultChecked /> Кассы</label>
              <label><input name="platform" type="checkbox" defaultChecked /> Платформы</label>
              <label><input name="hall" type="checkbox" defaultChecked /> Зал ожидания</label>
              <label><input name="passage" type="checkbox" defaultChecked /> Проходы</label>
            </div>
            <label>Замечание<textarea name="note" defaultValue="Ограничений нет" /></label>
          </Form>
        ) : null}
        {screen === "repair" ? (
          <Form
            onBack={back}
            onSave={(data) => show(db.saveRepair(read(data, "object"), read(data, "text"), read(data, "status"), role))}
          >
            <label>Объект<input name="object" defaultValue="Платформа 2" /></label>
            <label>Описание<input name="text" defaultValue="Не горит табло посадки" /></label>
            <label>
              Статус
              <select name="status" value={repairStatus} onChange={(event) => setRepairStatus(event.target.value)}>
                <option>Открыта</option>
                <option>Закрыта</option>
              </select>
            </label>
            <Records
              title="Заявки"
              rows={db.repairs.map((item) => `${item.object}: ${item.closed ? "закрыта" : "открыта"}`)}
              mark={tick}
            />
          </Form>
        ) : null}
        {screen === "inventory" ? (
          <Form
            onBack={back}
            onSave={(data) => show(db.saveInventory(read(data, "date"), read(data, "object"), read(data, "fact")))}
          >
            <label>Дата<input name="date" type="date" defaultValue="2026-10-20" /></label>
            <label>Объект<input name="object" defaultValue="Касса 1, принтер купонов" /></label>
            <label>Факт<input name="fact" defaultValue="В наличии" /></label>
            <Records title="Акты" rows={db.inventories.map((item) => `${item.date}: ${item.object}`)} mark={tick} />
          </Form>
        ) : null}
        {screen === "shift" ? (
          <Form
            onBack={back}
            onSave={() => {
              const answer = db.closeShift();
              show(answer);
              if (answer.ok) setScreen("login");
            }}
          >
            <p>Роль: {role}. Закрытие доступно, если нет черновика продажи.</p>
          </Form>
        ) : null}
      </main>
    </div>
  );
};

const Records = ({ title, rows }: { title: string; rows: string[]; mark: number }) => (
  <div className="records">
    <p>{title}</p>
    <ul>
      {rows.map((row) => (
        <li key={row}>{row}</li>
      ))}
    </ul>
  </div>
);

const Form = ({
  children,
  onBack,
  onSave,
}: {
  children: ReactNode;
  onBack: () => void;
  onSave: (data: FormData) => void;
}) => (
  <form
    className="form"
    onSubmit={(event) => {
      event.preventDefault();
      onSave(new FormData(event.currentTarget));
    }}
  >
    {children}
    <div className="row">
      <button className="primary" type="submit">Сохранить</button>
      <button type="button" onClick={onBack}>К меню</button>
    </div>
  </form>
);

export default App;
