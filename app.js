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

const picker = document.getElementById("picker");
const itemHeight = 40;

// padding сверху и снизу, чтобы первый/последний попадали в центр
picker.appendChild(document.createElement("div")).style.height = "60px";

regions.forEach(r => {
  const div = document.createElement("div");
  div.className = "picker-item";
  div.textContent = r;
  picker.appendChild(div);
});

picker.appendChild(document.createElement("div")).style.height = "60px";

document.getElementById("submit").onclick = () => {
  const name = document.getElementById("full_name").value.trim();
  const phone = document.getElementById("phone").value.trim();

  if (!name || !phone) {
    alert("Заполните все поля");
    return;
  }

  const index = Math.round(picker.scrollTop / itemHeight);
  const region = regions[index];

  tg.sendData(`REG|${name}|${phone}|${region}`);
};
