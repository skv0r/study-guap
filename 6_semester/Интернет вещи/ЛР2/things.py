"""
Цифровые двойники «Умный дом» (вариант 1).
Иерархия согласована с первой лабораторной работой.
"""
from __future__ import annotations

import abc
import logging

logger = logging.getLogger(__name__)


class Thing(abc.ABC):
    """Базовая абстрактная «вещь» IoT."""

    def __init__(self, id_: str, name: str, is_online: bool = True) -> None:
        self.id = id_
        self.name = name
        self.is_online = is_online
        logger.info("Создан цифровой двойник: %s (%s)", name, id_)

    @abc.abstractmethod
    def get_status(self) -> str:
        """Текстовое состояние для мониторинга."""

    def ping(self) -> None:
        """Проверка связи; достаточно сообщения в лог (требования ЛР2)."""
        logger.info("Метод ping вызван у вещи %s", self.name)


class TemperatureSensor(Thing):
    """Датчик температуры."""

    def __init__(self, id_: str, name: str, initial_c: float = 22.0) -> None:
        super().__init__(id_, name)
        self.current_temperature_c = initial_c

    def set_temperature(self, value: float) -> None:
        logger.info("TemperatureSensor.set_temperature запущен, вещь=%s", self.name)
        self.current_temperature_c = value

    def get_temperature(self) -> float:
        logger.info("TemperatureSensor.get_temperature запущен, вещь=%s", self.name)
        return self.current_temperature_c

    def get_status(self) -> str:
        return f"{self.name}: {self.current_temperature_c:.1f} °C"


class LightController(Thing):
    """Управление освещением."""

    def __init__(self, id_: str, name: str) -> None:
        super().__init__(id_, name)
        self.brightness_percent = 0
        self.is_enabled = False

    def set_light(self, enabled: bool, brightness: int) -> None:
        logger.info("LightController.set_light запущен, вещь=%s", self.name)
        self.is_enabled = enabled
        self.brightness_percent = max(0, min(100, brightness))

    def get_status(self) -> str:
        state = "ВКЛ" if self.is_enabled else "ВЫКЛ"
        return f"{self.name}: {state}, яркость={self.brightness_percent}%"


class SmartLock(Thing):
    """Умный замок."""

    def __init__(self, id_: str, name: str) -> None:
        super().__init__(id_, name)
        self.locked = True

    def set_locked(self, locked: bool) -> None:
        logger.info("SmartLock.set_locked запущен, вещь=%s", self.name)
        self.locked = locked

    def get_status(self) -> str:
        state = "ЗАКРЫТ" if self.locked else "ОТКРЫТ"
        return f"{self.name}: {state}"


class DeviceDataStore:
    """Упрощённое хранилище записей мониторинга (сущность данных)."""

    def __init__(self) -> None:
        self._by_id: dict[str, list[str]] = {}

    def save_record(self, thing: Thing, payload: str) -> None:
        logger.info("DeviceDataStore.save_record запущен для %s", thing.id)
        self._by_id.setdefault(thing.id, []).append(payload)


class MainControlUnit:
    """Центральный логический узел."""

    def __init__(self, store: DeviceDataStore) -> None:
        self._store = store
        self._things: list[Thing] = []

    def register_thing(self, thing: Thing) -> None:
        logger.info("MainControlUnit.register_thing запущен: %s", thing.name)
        self._things.append(thing)

    def collect_monitoring_snapshot(self) -> None:
        logger.info("MainControlUnit.collect_monitoring_snapshot запущен")
        for t in self._things:
            line = f"{t.get_status()}"
            self._store.save_record(t, line)
