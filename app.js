function sendRegistration() {
  const name = document.getElementById("full_name").value.trim();
  const phone = document.getElementById("phone").value.trim();
  const region = document.getElementById("region").value;

  if (!name || !phone || !region) {
    alert("Пожалуйста, заполните все поля");
    return;
  }

  tg.sendData(`REG|${name}|${phone}|${region}`);
}
