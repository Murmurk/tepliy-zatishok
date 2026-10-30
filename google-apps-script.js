/**
 * Google Apps Script — приём заявок из сайта «Теплий Затишок»
 * Пишет в таблицу + шлёт уведомление в Telegram.
 *
 * Как подключить:
 * 1. Google Таблицы → Расширения → Apps Script
 * 2. Вставить этот код в Code.gs
 * 3. Деплой → Веб-приложение → Доступ: Все (Anyone)
 * 4. URL деплоя вписать в CONFIG.ORDER_URL в index.html
 */

var TG_BOT = "ВСТАВЬ_ТОКЕН_БОТА";
var TG_CHAT = "ВСТАВЬ_CHAT_ID";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    if (sheet.getLastRow() === 0) {
      var headers = ["Дата", "Ім'я", "Телефон", "Колір", "Розмір", "Ціна", "Місто", "Відділення НП", "Без дзвінка", "Оброблено"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f0f0f0");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 140);
      sheet.setColumnWidth(7, 160);
      sheet.setColumnWidth(8, 220);
    }

    var ts = data.createdAt ? new Date(data.createdAt) : new Date();
    var row = sheet.getLastRow() + 1;

    sheet.appendRow([
      ts,
      data.name || "",
      data.phone || "",
      data.color || "",
      data.size || "",
      data.price || "",
      data.city || "",
      data.warehouse || "",
      data.noCall ? "Так" : "",
      false
    ]);

    sheet.getRange(row, 1).setNumberFormat("dd.MM.yyyy HH:mm");
    sheet.getRange(row, 10).insertCheckboxes();

    var text = "\u{1F6D2} *Нове замовлення*\n\n"
      + "\u{1F464} " + (data.name || "") + "\n"
      + "\u{1F4DE} " + (data.phone || "") + "\n"
      + "\u{1F3A8} " + (data.color || "") + " · " + (data.size || "") + "\n"
      + "\u{1F4B0} " + (data.price || "") + " грн\n"
      + "\u{1F3D9} " + (data.city || "") + "\n"
      + "\u{1F4E6} " + (data.warehouse || "")
      + (data.noCall ? "\n\n⚡ Без дзвінка менеджера" : "");

    UrlFetchApp.fetch("https://api.telegram.org/bot" + TG_BOT + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: TG_CHAT, text: text, parse_mode: "Markdown" })
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
