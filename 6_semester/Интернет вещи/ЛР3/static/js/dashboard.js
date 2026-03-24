/* Отдельный AJAX на каждую вещь — по методичке */

function pollTemperature() {
  $.ajax({
    type: "GET",
    url: "/connect/temperature",
    dataType: "json",
    data: {},
    success: function (response) {
      $("#temperature_value").val(String(response.value) + " °C");
    },
  });
}

function pollLight() {
  $.ajax({
    type: "GET",
    url: "/connect/light",
    dataType: "json",
    data: {},
    success: function (response) {
      $("#light_state").val(response.enabled ? "ВКЛ" : "ВЫКЛ");
      $("#light_brightness").val(String(response.brightness) + " %");
    },
  });
}

function pollLock() {
  $.ajax({
    type: "GET",
    url: "/connect/lock",
    dataType: "json",
    data: {},
    success: function (response) {
      $("#lock_state").val(response.locked ? "ЗАКРЫТ" : "ОТКРЫТ");
    },
  });
}

setInterval(pollTemperature, 1000);
setInterval(pollLight, 1000);
setInterval(pollLock, 1000);
