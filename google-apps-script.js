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
    }

    var ts = data.createdAt ? new Date(data.createdAt) : new Date();
    var phone = String(data.phone || "").split("+").join("");
    var row = sheet.getLastRow() + 1;

    sheet.getRange(row, 1, 1, 10).setNumberFormat("@");
    sheet.getRange(row, 1, 1, 10).setValues([[
      Utilities.formatDate(ts, "Europe/Kiev", "dd.MM.yyyy HH:mm"),
      data.name || "",
      phone,
      data.color || "",
      data.size || "",
      String(data.price || ""),
      data.city || "",
      data.warehouse || "",
      data.noCall ? "Так" : "",
      ""
    ]]);
    sheet.getRange(row, 10).insertCheckboxes();

    var TG_BOT = "9149023389:AAH97uU4OKd1-XpUkRwQLcNqesLCxMDZuFc";
    var TG_CHAT = "8609163495";
    var text = "Нове замовлення\n\n"
      + "Ім'я: " + (data.name || "") + "\n"
      + "Тел: " + (data.phone || "") + "\n"
      + "Колір: " + (data.color || "") + " / " + (data.size || "") + "\n"
      + "Ціна: " + (data.price || "") + " грн\n"
      + "Місто: " + (data.city || "") + "\n"
      + "НП: " + (data.warehouse || "")
      + (data.noCall ? "\n\nБез дзвінка менеджера" : "");

    UrlFetchApp.fetch("https://api.telegram.org/bot" + TG_BOT + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: TG_CHAT, text: text })
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
