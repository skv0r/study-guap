$(document).ready(() => {
    // Анимация jQuery с явными параметрами длительности и функции плавности.
    $("#runAnimationBtn").on("click", () => {
        const $card = $("#animatedCard");
        const $button = $("#runAnimationBtn");

        $button.stop(true, true);
        $button
            .animate({ left: "10px", top: "-4px" }, 150)
            .animate({ left: "0px", top: "0px" }, 180);

        $card.stop(true, true);
        $card.removeClass("glow");
        $card.css({ left: "10px", top: "80px", opacity: 1 });

        $card
            .animate(
                { left: "calc(100% - 270px)", top: "24px", opacity: 0.85 },
                {
                    duration: 900,
                    easing: "swing"
                }
            )
            .animate(
                { left: "40%", top: "130px", opacity: 0.65 },
                {
                    duration: 700,
                    easing: "linear"
                }
            )
            .animate(
                { left: "10px", top: "80px", opacity: 1 },
                {
                    duration: 900,
                    easing: "linear"
                }
            )
            .promise()
            .done(() => {
                $card.addClass("glow");
                setTimeout(() => $card.removeClass("glow"), 450);
            });
    });

    // Диаграмма Chart.js: самые популярные веб-браузеры в мире.
    const labels = ["Chrome", "Safari", "Edge", "Firefox", "Yandex", "Opera"];
    const values = [64, 20, 5, 3, 3, 2];

    const ctx = document.getElementById("jsChart");
    new Chart(ctx, {
        type: "bar",
        data: {
            labels,
            datasets: [
                {
                    label: "Популярность, %",
                    data: values,
                    backgroundColor: [
                        "#2563eb",
                        "#14b8a6",
                        "#f97316",
                        "#6366f1",
                        "#e11d48",
                        "#10b981"
                    ],
                    borderColor: "#1f2937",
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            plugins: {
                legend: {
                    display: true
                },
                title: {
                    display: true,
                    text: "Самые популярные веб-браузеры в мире"
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 70
                }
            }
        }
    });
});
