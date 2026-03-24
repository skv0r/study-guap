"""
Лабораторная работа №3: эмуляторы вещей, JSON и AJAX-опрос мониторинга.
"""
from __future__ import annotations

import logging

from flask import Flask, render_template

import things

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

app = Flask(__name__)

temp_sensor = things.TemperatureSensor("sensor-1", "Температурный датчик")
light = things.LightController("light-1", "Свет в гостиной")
door_lock = things.SmartLock("lock-1", "Замок входной двери")


@app.route("/")
def index():
    return render_template("dashboard.html")


@app.route("/connect/temperature")
def connect_temperature():
    return temp_sensor.connect()


@app.route("/connect/light")
def connect_light():
    return light.connect()


@app.route("/connect/lock")
def connect_lock():
    return door_lock.connect()


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
