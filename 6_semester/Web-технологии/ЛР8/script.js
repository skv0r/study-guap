document.addEventListener("DOMContentLoaded", () => {
    // Задание 1: слайд-шоу из 3 изображений с интервалом 1 секунда.
    const slideshowImage = document.getElementById("slideshowImage");
    const slideshowLink = document.getElementById("slideshowLink");
    const slides = [
        { src: "../ЛР7/bmw.png", href: "../ЛР7/bmw.png", alt: "Слайд BMW" },
        { src: "../ЛР7/audi.png", href: "../ЛР7/audi.png", alt: "Слайд Audi" },
        { src: "../ЛР7/toyota.png", href: "../ЛР7/toyota.png", alt: "Слайд Toyota" },
        { src: "../ЛР7/mega-mercded.png", href: "../ЛР7/mega-mercded.png", alt: "Слайд Mercedes 1" },
        { src: "../ЛР7/mercdedec.png", href: "../ЛР7/mercdedec.png", alt: "Слайд Mercedes 2" },
        { src: "../ЛР7/mercedez.png", href: "../ЛР7/mercedez.png", alt: "Слайд Mercedes 3" },
        { src: "../ЛР7/bmw.png", href: "../ЛР7/bmw.png", alt: "Слайд BMW 2" }
    ];
    let slideIndex = 0;

    setInterval(() => {
        slideIndex = (slideIndex + 1) % slides.length;
        slideshowImage.src = slides[slideIndex].src;
        slideshowImage.alt = slides[slideIndex].alt;
        slideshowLink.href = slides[slideIndex].href;
    }, 1000);

    // Задание 2: выпадающие меню для 3 пунктов навигации.
    const menuItems = document.querySelectorAll(".menu-item");

    menuItems.forEach((item) => {
        const toggleButton = item.querySelector(".menu-toggle");
        toggleButton.addEventListener("click", (event) => {
            event.stopPropagation();

            menuItems.forEach((other) => {
                if (other !== item) {
                    other.classList.remove("open");
                }
            });

            item.classList.toggle("open");
        });
    });

    window.addEventListener("click", (event) => {
        const target = event.target;
        menuItems.forEach((item) => {
            if (!target.closest(".menu-item")) {
                item.classList.remove("open");
            }
        });
    });

    // Задание 3: галерея (7 изображений, видно одновременно 3).
    const galleryTrack = document.getElementById("galleryTrack");
    const galleryPrev = document.getElementById("galleryPrev");
    const galleryNext = document.getElementById("galleryNext");
    const allImages = galleryTrack.querySelectorAll("img");

    let startIndex = 0;
    const visibleCount = 3;
    const maxStartIndex = allImages.length - visibleCount;

    function renderGallery() {
        const shiftPercent = (startIndex * 100) / visibleCount;
        galleryTrack.style.transform = `translateX(-${shiftPercent}%)`;
    }

    galleryPrev.addEventListener("click", () => {
        if (startIndex > 0) {
            startIndex -= 1;
            renderGallery();
        }
    });

    galleryNext.addEventListener("click", () => {
        if (startIndex < maxStartIndex) {
            startIndex += 1;
            renderGallery();
        }
    });
});
