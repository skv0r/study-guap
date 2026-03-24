"""
Лабораторная работа №4: передача управляющих команд (отдельные маршруты).
Маршруты мониторинга /connect/* сохранены как в ЛР3.
"""
from __future__ import annotations

import logging

from flask import Flask, render_template, request

import things

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

app = Flask(__name__)

temp_sensor = things.TemperatureSensor("sensor-1", "Температурный датчик")
light = things.LightController("light-1", "Свет в гостиной")
door_lock = things.SmartLock("lock-1", "Замок входной двери")


@app.route("/")
def index():
    return render_template("control.html")


# --- Мониторинг (без изменений логики ЛР3) ---


@app.route("/connect/temperature")
def connect_temperature():
    return temp_sensor.connect()


@app.route("/connect/light")
def connect_light():
    return light.connect()


@app.route("/connect/lock")
def connect_lock():
    return door_lock.connect()


# --- Управление (новые маршруты) ---


@app.route("/command/temperature")
def command_temperature():
    value = request.args.get("value", "")
    return temp_sensor.command_set_temperature(value)


@app.route("/command/light")
def command_light():
    on = request.args.get("on", "0")
    br = request.args.get("brightness", "0")
    return light.command_set_light(on, br)


@app.route("/command/lock")
def command_lock():
    locked = request.args.get("locked", "0")
    return door_lock.command_set_lock(locked)


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
