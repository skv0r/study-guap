import { useState, type ReactNode } from "react";

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
  const [stationOpen, setStationOpen] = useState(true);
  const [shiftOpen, setShiftOpen] = useState(false);
  const [draft, setDraft] = useState(false);
  const [boarded, setBoarded] = useState(false);
  const [tripOpen, setTripOpen] = useState(true);
  const [cancelled, setCancelled] = useState(false);
  const [repairClosed, setRepairClosed] = useState(false);
  const [repairStatus, setRepairStatus] = useState("Открыта");
  const [ready, setReady] = useState({ cash: true, platform: true, hall: true, passage: true });
  const [notice, setNotice] = useState("");
  const [bad, setBad] = useState(false);

  function say(text: string, error = false) {
    setNotice(text);
    setBad(error);
  }

  function enter() {
    if (needsShift(role)) {
      if (!stationOpen) {
        say("Станция не в обслуживании. Смену открыть нельзя.", true);
        return;
      }
      setShiftOpen(true);
      say("Смена открыта.");
    } else {
      say("Вход выполнен.");
    }
    setScreen("menu");
  }

  function leave() {
    setScreen("login");
    say("");
  }

  function back() {
    setScreen("menu");
    say("");
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
            {needsShift(role) ? (shiftOpen ? " · смена открыта" : " · смена закрыта") : ""}
            {stationOpen ? " · станция в обслуживании" : " · станция закрыта"}
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
              <input defaultValue="Северная" />
            </label>
            <button className="primary" type="submit">
              {needsShift(role) ? "Войти и открыть смену" : "Войти"}
            </button>
          </form>
        ) : null}
        {screen === "menu" ? (
          <div className="menu">
            {menuByRole[role].map((item) => (
              <button key={item.id} type="button" onClick={() => { setScreen(item.id); say(""); }}>
                {item.label}
              </button>
            ))}
          </div>
        ) : null}
        {screen === "stations" ? (
          <Form onBack={back} onSave={() => say("Станция сохранена. Код уникален.")}>
            <label>Код<input defaultValue="SVR" /></label>
            <label>Название<input defaultValue="Северная" /></label>
          </Form>
        ) : null}
        {screen === "schemes" ? (
          <Form onBack={back} onSave={() => say("Схема сохранена. Номера мест не повторяются.")}>
            <label>Тип вагона<input defaultValue="Плацкарт" /></label>
            <label>Номера мест<input defaultValue="1–36" /></label>
          </Form>
        ) : null}
        {screen === "route" ? (
          <Form onBack={back} onSave={() => say("Маршрут сохранён: две станции, без повтора.")}>
            <label>Код<input defaultValue="M-12" /></label>
            <label>Название<input defaultValue="Северная — Южная" /></label>
            <label>Станции по порядку<input defaultValue="Северная — Южная" /></label>
            <label>Время хода, мин<input defaultValue="95" /></label>
          </Form>
        ) : null}
        {screen === "publish" ? (
          <Form
            onBack={back}
            onSave={() => {
              if (cancelled) {
                say("Рейс отменён. Тем же номером не открывается.", true);
                return;
              }
              if (tripOpen) {
                say("Рейс с этим номером уже открыт.", true);
                return;
              }
              setTripOpen(true);
              say("Рейс 102 опубликован. Фонд мест создан.");
            }}
          >
            <label>Номер рейса<input defaultValue="102" /></label>
            <label>Дата<input type="date" defaultValue="2026-10-20" /></label>
            <label>Маршрут<input defaultValue="M-12 Северная — Южная" /></label>
            <label>Схема вагона<input defaultValue="Плацкарт" /></label>
            <label>Время отправления и прибытия<input defaultValue="08:40–10:15" /></label>
          </Form>
        ) : null}
        {screen === "cancel" ? (
          <Form
            onBack={back}
            onSave={() => {
              if (cancelled) {
                say("Рейс уже отменён. Повторная отмена не создана.", true);
                return;
              }
              setCancelled(true);
              setTripOpen(false);
              setDraft(false);
              say(
                draft
                  ? "Рейс отменён. Черновик продажи не проведён, новые продажи закрыты."
                  : "Рейс отменён. Новые продажи закрыты, билеты помечены.",
              );
            }}
          >
            <label>Рейс<input defaultValue="102" /></label>
            <label>Причина<input defaultValue="Отмена перевозчиком" /></label>
          </Form>
        ) : null}
        {screen === "sale" ? (
          <Form
            onBack={back}
            onSave={() => {
              if (!stationOpen || !shiftOpen) {
                say("Нет обслуживания станции или открытой смены.", true);
                return;
              }
              if (!tripOpen) {
                say("Рейс закрыт. Билет не создан.", true);
                return;
              }
              setDraft(false);
              say("Билет проведён, купон отправлен на печать.");
            }}
          >
            <label>Станции и дата<input defaultValue="Северная — Южная, 20.10.2026" /></label>
            <label>Место<input defaultValue="12" /></label>
            <label>Пассажир<input defaultValue="Иванов Иван" /></label>
            <label>Цена из прайса<input defaultValue="1 850" /></label>
            <label>
              <span>Черновик</span>
              <input
                type="checkbox"
                checked={draft}
                onChange={(event) => setDraft(event.target.checked)}
              />
            </label>
          </Form>
        ) : null}
        {screen === "refund" ? (
          <Form
            onBack={back}
            onSave={() => {
              if (boarded) {
                say("Посадка уже отмечена. Возврат не проведён.", true);
                return;
              }
              say(tripOpen ? "Билет возвращён, место свободно." : "Билет погашен. Место закрыто: рейс отменён.");
            }}
          >
            <label>Номер билета<input defaultValue="102-12" /></label>
          </Form>
        ) : null}
        {screen === "boarding" ? (
          <Form
            onBack={back}
            onSave={() => {
              if (!shiftOpen || !stationOpen) {
                say("Нет смены или станция не в обслуживании.", true);
                return;
              }
              if (boarded) {
                say("Отметка посадки уже есть.", true);
                return;
              }
              setBoarded(true);
              say("Посадка отмечена один раз.");
            }}
          >
            <label>Номер билета или купона<input defaultValue="102-12" /></label>
          </Form>
        ) : null}
        {screen === "service" ? (
          <Form
            onBack={back}
            onSave={() => {
              const allReady = Object.values(ready).every(Boolean);
              if (!allReady) {
                say("Неготовый объект: заявка на ремонт создана или обновлена.");
                return;
              }
              if (shiftOpen && stationOpen) {
                say("Есть открытая смена. Сутки не закрыты.", true);
                return;
              }
              setStationOpen((value) => !value);
              say(stationOpen ? "Сутки закрыты. Новые смены недоступны." : "Сутки открыты. Станция в обслуживании.");
            }}
          >
            <div className="checks">
              <label>
                <input
                  type="checkbox"
                  checked={ready.cash}
                  onChange={(event) => setReady((value) => ({ ...value, cash: event.target.checked }))}
                />{" "}
                Кассы
              </label>
              <label>
                <input
                  type="checkbox"
                  checked={ready.platform}
                  onChange={(event) => setReady((value) => ({ ...value, platform: event.target.checked }))}
                />{" "}
                Платформы
              </label>
              <label>
                <input
                  type="checkbox"
                  checked={ready.hall}
                  onChange={(event) => setReady((value) => ({ ...value, hall: event.target.checked }))}
                />{" "}
                Зал ожидания
              </label>
              <label>
                <input
                  type="checkbox"
                  checked={ready.passage}
                  onChange={(event) => setReady((value) => ({ ...value, passage: event.target.checked }))}
                />{" "}
                Проходы
              </label>
            </div>
            <label>Замечание<textarea defaultValue="Ограничений нет" /></label>
          </Form>
        ) : null}
        {screen === "repair" ? (
          <Form
            onBack={back}
            onSave={() => {
              if (repairClosed) {
                say("Заявка уже закрыта. Той же неисправностью повторно не открывается.", true);
                return;
              }
              if (repairStatus === "Закрыта" && role !== "Инженер по ремонту") {
                say("Закрыть заявку может только инженер после отметки устранения.", true);
                return;
              }
              if (repairStatus === "Закрыта") {
                setRepairClosed(true);
                say("Заявка закрыта.");
                return;
              }
              say("Заявка на ремонт открыта или обновлена.");
            }}
          >
            <label>Объект<input defaultValue="Платформа 2" /></label>
            <label>Описание<input defaultValue="Не горит табло посадки" /></label>
            <label>
              Статус
              <select value={repairStatus} onChange={(event) => setRepairStatus(event.target.value)}>
                <option>Открыта</option>
                <option>Закрыта</option>
              </select>
            </label>
          </Form>
        ) : null}
        {screen === "inventory" ? (
          <Form onBack={back} onSave={() => say("Инвентаризация за 20.10.2026 проведена. Повтор на эту дату закрыт.")}>
            <label>Дата<input type="date" defaultValue="2026-10-20" /></label>
            <label>Объект<input defaultValue="Касса 1, принтер купонов" /></label>
            <label>Факт<input defaultValue="В наличии" /></label>
          </Form>
        ) : null}
        {screen === "shift" ? (
          <Form
            onBack={back}
            onSave={() => {
              if (draft) {
                say("Есть черновик продажи. Смена не закрыта.", true);
                return;
              }
              setShiftOpen(false);
              say("Смена закрыта.");
              setScreen("login");
            }}
          >
            <p>Роль: {role}. Закрытие доступно, если нет черновика продажи.</p>
          </Form>
        ) : null}
      </main>
    </div>
  );
};

const Form = ({
  children,
  onBack,
  onSave,
}: {
  children: ReactNode;
  onBack: () => void;
  onSave: () => void;
}) => (
  <form
    className="form"
    onSubmit={(event) => {
      event.preventDefault();
      onSave();
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
