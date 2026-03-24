"""
Лабораторная работа №2: Flask-приложение и демонстрация методов классов.
"""
from __future__ import annotations

import logging

from flask import Flask, render_template_string

import things

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

app = Flask(__name__)

store = things.DeviceDataStore()
mcu = things.MainControlUnit(store)

temp_sensor = things.TemperatureSensor("sensor-1", "Температурный датчик", 22.5)
light = things.LightController("light-1", "Свет в гостиной")
lock = things.SmartLock("lock-1", "Замок входной двери")

mcu.register_thing(temp_sensor)
mcu.register_thing(light)
mcu.register_thing(lock)

_SIMPLE_PAGE = """
<!DOCTYPE html>
<html lang="ru">
<head><meta charset="UTF-8"><title>Умный дом — ЛР2</title></head>
<body>
  <h1>Лабораторная работа №2</h1>
  <p>Откройте консоль запуска Flask: там будут сообщения о вызовах методов.</p>
  <pre>{{ status }}</pre>
</body>
</html>
"""


@app.route("/")
def hello_world():
    # Демонстрация вызовов методов (достаточно лога сервера по методичке)
    temp_sensor.ping()
    temp_sensor.set_temperature(23.0)
    light.set_light(True, 65)
    lock.set_locked(True)
    mcu.collect_monitoring_snapshot()

    lines = [temp_sensor.get_status(), light.get_status(), lock.get_status()]
    return render_template_string(_SIMPLE_PAGE, status="\n".join(lines))


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
