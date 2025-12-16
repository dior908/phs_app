const tg = window.Telegram.WebApp;
tg.ready();
tg.expand();

const regions = [
  "г. Ташкент",
  "Ташкентская область",
  "Андижанская область",
  "Ферганская область",
  "Наманганская область",
  "Самаркандская область",
  "Бухарская область",
  "Навоийская область",
  "Джизакская область",
  "Сырдарьинская область",
  "Кашкадарьинская область",
  "Сурхандарьинская область",
  "Хорезмская область",
  "Республика Каракалпакстан"
];

const wheel = document.getElementById("wheel");
let startY = 0;
let offsetY = 0;
let currentIndex = 0;
const itemHeight = 40;

regions.forEach(r => {
  const li = document.createElement("li");
  li.textContent = r;
  wheel.appendChild(li);
});

updateActive();

wheel.addEventListener("touchstart", e => {
  startY = e.touches[0].clientY;
});

wheel.addEventListener("touchmove", e => {
  const delta = e.touches[0].clientY - startY;
  wheel.style.transform = `translateY(${offsetY + delta}px)`;
});

wheel.addEventListener("touchend", e => {
  const delta = e.changedTouches[0].clientY - startY;
  offsetY += delta;

  currentIndex = Math.round(-offsetY / itemHeight);
  currentIndex = Math.max(0, Math.min(regions.length - 1, currentIndex));

  offsetY = -currentIndex * itemHeight;
  wheel.style.transform = `translateY(${offsetY}px)`;

  updateActive();
});

function updateActive() {
  [...wheel.children].forEach((li, i) => {
    li.classList.toggle("active", i === currentIndex);
  });
}

function sendRegistration() {
  const name = document.getElementById("full_name").value.trim();
  const phone = document.getElementById("phone").value.trim();
  const region = regions[currentIndex];

  if (!name || !phone) {
    alert("Заполните все поля");
    return;
  }

  tg.sendData(`REG|${name}|${phone}|${region}`);
}
