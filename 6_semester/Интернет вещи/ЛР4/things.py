"""
Умный дом: мониторинг (connect) + управляющие команды (отдельные методы).
ЛР4: проверка типов минимальная; полная валидация — в ЛР5.
"""
from __future__ import annotations

import abc
import json
import logging
import random

logger = logging.getLogger(__name__)


class Thing(abc.ABC):
    def __init__(self, id_: str, name: str, is_online: bool = True) -> None:
        self.id = id_
        self.name = name
        self.is_online = is_online

    @abc.abstractmethod
    def get_status(self) -> str:
        pass

    @abc.abstractmethod
    def emulate(self) -> None:
        pass

    @abc.abstractmethod
    def connect(self) -> str:
        pass


class TemperatureSensor(Thing):
    def __init__(self, id_: str, name: str, initial_c: float = 22.0) -> None:
        super().__init__(id_, name)
        self.current_temperature_c = initial_c

    def get_status(self) -> str:
        return f"{self.name}: {self.current_temperature_c:.1f} °C"

    def emulate(self) -> None:
        self.current_temperature_c = round(random.uniform(18.0, 28.0), 1)

    def connect(self) -> str:
        logger.info("connect (мониторинг) TemperatureSensor %s", self.id)
        self.emulate()
        return json.dumps({"id": self.id, "value": self.current_temperature_c}, ensure_ascii=False)

    def command_set_temperature(self, value_raw: str) -> str:
        """Управление: задать температуру (число)."""
        logger.info("команда температура=%s для %s", value_raw, self.id)
        value = float(value_raw.replace(",", "."))
        self.current_temperature_c = round(value, 1)
        return json.dumps({"ok": True, "value": self.current_temperature_c}, ensure_ascii=False)


class LightController(Thing):
    def __init__(self, id_: str, name: str) -> None:
        super().__init__(id_, name)
        self.brightness_percent = 40
        self.is_enabled = True

    def get_status(self) -> str:
        state = "ВКЛ" if self.is_enabled else "ВЫКЛ"
        return f"{self.name}: {state}, яркость={self.brightness_percent}%"

    def emulate(self) -> None:
        if random.random() < 0.12:
            self.is_enabled = not self.is_enabled
        self.brightness_percent = random.randint(0, 100)

    def connect(self) -> str:
        logger.info("connect (мониторинг) LightController %s", self.id)
        self.emulate()
        return json.dumps(
            {"id": self.id, "enabled": self.is_enabled, "brightness": self.brightness_percent},
            ensure_ascii=False,
        )

    def command_set_light(self, enabled_raw: str, brightness_raw: str) -> str:
        """Управление: вкл/выкл (0/1) и яркость 0..100."""
        logger.info("команда свет on=%s br=%s для %s", enabled_raw, brightness_raw, self.id)
        self.is_enabled = enabled_raw in ("1", "true", "True", "да", "ДА")
        self.brightness_percent = max(0, min(100, int(brightness_raw)))
        return json.dumps(
            {"ok": True, "enabled": self.is_enabled, "brightness": self.brightness_percent},
            ensure_ascii=False,
        )


class SmartLock(Thing):
    def __init__(self, id_: str, name: str) -> None:
        super().__init__(id_, name)
        self.locked = True

    def get_status(self) -> str:
        state = "ЗАКРЫТ" if self.locked else "ОТКРЫТ"
        return f"{self.name}: {state}"

    def emulate(self) -> None:
        if random.random() < 0.08:
            self.locked = not self.locked

    def connect(self) -> str:
        logger.info("connect (мониторинг) SmartLock %s", self.id)
        self.emulate()
        return json.dumps({"id": self.id, "locked": self.locked}, ensure_ascii=False)

    def command_set_lock(self, locked_raw: str) -> str:
        """Управление: 1/0 — закрыть/открыть."""
        logger.info("команда замок locked=%s для %s", locked_raw, self.id)
        self.locked = locked_raw in ("1", "true", "True", "да", "ДА")
        return json.dumps({"ok": True, "locked": self.locked}, ensure_ascii=False)
