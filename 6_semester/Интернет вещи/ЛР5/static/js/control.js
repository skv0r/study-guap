function pollTemperature() {
  $.ajax({
    type: "GET",
    url: "/connect/temperature",
    dataType: "json",
    success: function (r) {
      $("#temperature_value").val(String(r.value) + " °C");
    },
  });
}

function pollLight() {
  $.ajax({
    type: "GET",
    url: "/connect/light",
    dataType: "json",
    success: function (r) {
      $("#light_state").val(r.enabled ? "ВКЛ" : "ВЫКЛ");
      $("#light_brightness").val(String(r.brightness) + " %");
    },
  });
}

function pollLock() {
  $.ajax({
    type: "GET",
    url: "/connect/lock",
    dataType: "json",
    success: function (r) {
      $("#lock_state").val(r.locked ? "ЗАКРЫТ" : "ОТКРЫТ");
    },
  });
}

function sendTemperatureCommand() {
  $.ajax({
    type: "GET",
    url: "/command/temperature",
    dataType: "json",
    data: { value: document.getElementById("temperature_cmd").value },
    success: function (r) {
      if (r.ok) alert("Ок: " + r.value);
      else alert("Ошибка: " + r.error);
    },
  });
}

function sendLightCommand() {
  $.ajax({
    type: "GET",
    url: "/command/light",
    dataType: "json",
    data: {
      on: document.getElementById("light_on").value,
      brightness: document.getElementById("light_br").value,
    },
    success: function (r) {
      if (r.ok) alert("Ок: свет обновлён");
      else alert("Ошибка: " + r.error);
    },
  });
}

function sendLockCommand() {
  $.ajax({
    type: "GET",
    url: "/command/lock",
    dataType: "json",
    data: { mode: document.getElementById("lock_cmd").value },
    success: function (r) {
      if (r.ok) alert("Ок: замок " + (r.locked ? "закрыт" : "открыт"));
      else alert("Ошибка: " + r.error);
    },
  });
}

setInterval(pollTemperature, 1000);
setInterval(pollLight, 1000);
setInterval(pollLock, 1000);

$(function () {
  $("#btn_temperature_send").on("click", sendTemperatureCommand);
  $("#btn_light_send").on("click", sendLightCommand);
  $("#btn_lock_send").on("click", sendLockCommand);
});
