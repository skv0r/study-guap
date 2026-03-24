"""
Лабораторная работа №5: валидация входящих JSON-параметров (через try/except и regex).
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


@app.route("/connect/temperature")
def connect_temperature():
    return temp_sensor.connect()


@app.route("/connect/light")
def connect_light():
    return light.connect()


@app.route("/connect/lock")
def connect_lock():
    return door_lock.connect()


@app.route("/command/temperature")
def command_temperature():
    return temp_sensor.command_set_temperature(request.args.get("value", ""))


@app.route("/command/light")
def command_light():
    return light.command_set_light(
        request.args.get("on", ""),
        request.args.get("brightness", ""),
    )


@app.route("/command/lock")
def command_lock():
    return door_lock.command_set_lock(request.args.get("mode", ""))


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
