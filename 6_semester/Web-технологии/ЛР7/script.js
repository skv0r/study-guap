alert("Здравствуйте! Меня зовут Буренков Григорий");

// Часть 1. Три способа объявления функции (одна и та же задача)
function sumByDeclaration(a, b) {
    return a + b;
}

const sumByExpression = function (a, b) {
    return a + b;
};

const sumByArrow = (a, b) => a + b;

const functionResult = document.getElementById("functionResult");
const a = 12;
const b = 8;

document.getElementById("declarationBtn").addEventListener("click", () => {
    functionResult.textContent = `Function Declaration: ${a} + ${b} = ${sumByDeclaration(a, b)}`;
});

document.getElementById("expressionBtn").addEventListener("click", () => {
    functionResult.textContent = `Function Expression: ${a} + ${b} = ${sumByExpression(a, b)}`;
});

document.getElementById("arrowBtn").addEventListener("click", () => {
    functionResult.textContent = `Arrow Function: ${a} + ${b} = ${sumByArrow(a, b)}`;
});

// Часть 2, вариант 1, пункт 1: поочередные модальные окна alert
const remindersBtn = document.getElementById("remindersBtn");
const reminders = [
    "Напоминание: до 1-го числа продлить проездную карту.",
    "Напоминание: до 10-го числа оплатить жилищно-коммунальные услуги.",
    "Напоминание: до 25-го числа ввести показания счетчиков.",
    "Напоминание: записаться к врачу и проверить план обследований."
];

remindersBtn.addEventListener("click", () => {
    // Важно: alert блокирует выполнение, поэтому можно вывести все напоминания
    // в одном обработчике клика "поочередно".
    for (const reminder of reminders) {
        alert(reminder);
    }

    remindersBtn.disabled = true;
    remindersBtn.textContent = "Напоминания выполнены";
});

// Часть 2, вариант 1, пункт 2: клонирование узлов (создание второго контейнера)
const cloneContainer1 = document.getElementById("cloneContainer1");
const cloneBtn = document.getElementById("cloneBtn");

cloneBtn.addEventListener("click", () => {
    // Клонируем контейнер вместе с содержимым и меняем только второй абзац.
    const newContainer = cloneContainer1.cloneNode(true);
    newContainer.id = `cloneContainer_${Date.now()}`;

    const p2 = newContainer.querySelector(".clone-p2");
    if (p2) {
        p2.textContent = "Второй абзац: текст во втором контейнере (измененный).";
    }

    cloneContainer1.insertAdjacentElement("afterend", newContainer);
});
