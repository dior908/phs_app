function send() {
  Telegram.WebApp.sendData(
    "REG|" +
    document.getElementById("name").value + "|" +
    document.getElementById("phone").value + "|" +
    document.getElementById("region").value
  )
}

function verify() {
  Telegram.WebApp.sendData(
    "CODE|" + document.getElementById("code").value
  )
}

function setPin() {
  Telegram.WebApp.sendData(
    "PIN|" + document.getElementById("pin").value
  )
}
