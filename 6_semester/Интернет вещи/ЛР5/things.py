"""
ЛР5: проверка входящих данных (число через float(), строка через регулярное выражение).
"""
from __future__ import annotations

import abc
import json
import logging
import random
import re

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
        """Числовой формат: приведение через float() в try/except."""
        logger.info("команда температура (с проверкой) value=%s", value_raw)
        try:
            value = float(str(value_raw).strip().replace(",", "."))
        except ValueError:
            logger.warning("Неверный тип: ожидалось число, получено: %r", value_raw)
            return json.dumps(
                {"ok": False, "error": "ожидается числовая температура"},
                ensure_ascii=False,
            )
        self.current_temperature_c = round(value, 1)
        return json.dumps({"ok": True, "value": self.current_temperature_c}, ensure_ascii=False)


class LightController(Thing):
    _ON_PATTERN = re.compile(r"^(1|0|true|false|вкл|выкл|да|нет)$", re.IGNORECASE)

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
        """
        Два формата: строковый шаблон для вкл/выкла и целое для яркости с try/except.
        """
        logger.info("команда свет (с проверкой) on=%s br=%s", enabled_raw, brightness_raw)
        s = str(enabled_raw).strip()
        if not self._ON_PATTERN.fullmatch(s):
            logger.warning("Строка включения не по шаблону: %r", enabled_raw)
            return json.dumps(
                {
                    "ok": False,
                    "error": "ожидается 1/0, true/false, вкл/выкл, да/нет",
                },
                ensure_ascii=False,
            )
        low = s.lower()
        self.is_enabled = low in ("1", "true", "вкл", "да")

        try:
            br = int(str(brightness_raw).strip())
        except ValueError:
            logger.warning("Яркость не целое число: %r", brightness_raw)
            return json.dumps({"ok": False, "error": "яркость должна быть целым числом"}, ensure_ascii=False)

        if not 0 <= br <= 100:
            return json.dumps({"ok": False, "error": "яркость от 0 до 100"}, ensure_ascii=False)
        self.brightness_percent = br
        return json.dumps(
            {"ok": True, "enabled": self.is_enabled, "brightness": self.brightness_percent},
            ensure_ascii=False,
        )


class SmartLock(Thing):
    _LOCK_WORD = re.compile(r"^(ЗАКРЫТ|ОТКРЫТ)$", re.IGNORECASE)

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

    def command_set_lock(self, mode_raw: str) -> str:
        """Строковый формат с обязательной проверкой регулярным выражением."""
        logger.info("команда замок (с проверкой) mode=%s", mode_raw)
        s = str(mode_raw).strip()
        if not self._LOCK_WORD.fullmatch(s):
            logger.warning("Недопустимое значение замка: %r", mode_raw)
            return json.dumps(
                {"ok": False, "error": "допустимы только слова ЗАКРЫТ или ОТКРЫТ"},
                ensure_ascii=False,
            )
        word = s.upper()
        self.locked = word == "ЗАКРЫТ"
        return json.dumps({"ok": True, "locked": self.locked}, ensure_ascii=False)
