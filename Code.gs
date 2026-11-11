function doGet() {
  return HtmlService.createHtmlOutput(
    "Pet Hotel Shift Manager is running!"
  );
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Shift Requests");
  const staffSheet = ss.getSheetByName("Staff");

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return HtmlService.createHtmlOutput("OK");
    }

    if (!sheet || !staffSheet) {
      throw new Error("Missing Shift Requests or Staff sheet");
    }

    const data = JSON.parse(e.postData.contents);
    const events = data.events || [];

    const token = PropertiesService.getScriptProperties()
      .getProperty("LINE_CHANNEL_ACCESS_TOKEN");

    if (!token) {
      throw new Error("Missing LINE token");
    }

    for (const event of events) {
      if (!event.source || !event.source.userId) continue;

      const userId = event.source.userId;

      if (event.type === "follow") {
        let row = findStaffRow(staffSheet, userId);

        if (row === 0) {
          staffSheet.appendRow([userId, ""]);
          row = staffSheet.getLastRow();
        }

        const name = staffSheet.getRange(row, 2).getValue();

        sendReply(
          token,
          event.replyToken,
          name
            ? "Welcome back, " + name + "! 🐾\nSend your availability like this:\n10/30 8:00-15:00"
            : "Welcome to the Pet Hotel Shift Manager! 🐾\nWhat's your name?",
          sheet
        );
        continue;
      }

      if (
        event.type !== "message" ||
        !event.message ||
        event.message.type !== "text"
      ) {
        continue;
      }

      const message = event.message.text.trim();
      let staffRow = findStaffRow(staffSheet, userId);

      if (staffRow === 0) {
        staffSheet.appendRow([userId, ""]);
        sendReply(
          token,
          event.replyToken,
          "Welcome! 🐾 What's your name?",
          sheet
        );
        continue;
      }

      let staffName = String(
        staffSheet.getRange(staffRow, 2).getValue()
      ).trim();

      if (!staffName) {
        if (!message || message.length > 60 ||
            /^\d{1,2}\/\d{1,2}\s/.test(message)) {
          sendReply(
            token,
            event.replyToken,
            "Please enter your name first (up to 60 characters).",
            sheet
          );
          continue;
        }

        staffSheet.getRange(staffRow, 2).setValue(message);

        sendReply(
          token,
          event.replyToken,
          "Nice to meet you, " + message + "! 🐾\n\n" +
          "Your name has been saved.\n" +
          "Send your availability like this:\n10/30 8:00-15:00",
          sheet
        );
        continue;
      }

      
const lines = message.split(/\r?\n/).filter(line => line.trim() !== "");

const requests = [];
const errors = [];

for (const line of lines) {
  const match = line.trim().match(
    /^(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$/
  );

  if (!match) {
    errors.push(line + " — invalid format");
    continue;
  }

  const month = Number(match[1]);
  const day = Number(match[2]);
  const sh = Number(match[3]);
  const sm = Number(match[4]);
  const eh = Number(match[5]);
  const em = Number(match[6]);

  const daysInMonth = month >= 1 && month <= 12
    ? new Date(2024, month, 0).getDate()
    : 0;

  if (
    day < 1 || day > daysInMonth ||
    sh > 23 || sm > 59 || eh > 23 || em > 59 ||
    sh * 60 + sm >= eh * 60 + em
  ) {
    errors.push(line + " — invalid date or time");
    continue;
  }

  requests.push([
    userId,
    staffName,
    String(month).padStart(2, "0") + "/" +
      String(day).padStart(2, "0"),
    String(sh).padStart(2, "0") + ":" +
      String(sm).padStart(2, "0"),
    String(eh).padStart(2, "0") + ":" +
      String(em).padStart(2, "0"),
    "Pending"
  ]);
}

if (requests.length === 0) {
  sendReply(
    token,
    event.replyToken,
    "I couldn't find any valid shift requests.\n\n" +
    "Please use this format, one shift per line:\n" +
    "10/10 8:00-15:00\n10/11 8:00-15:00",
    sheet
  );
  continue;
}

sheet.getRange(
  sheet.getLastRow() + 1,
  1,
  requests.length,
  6
).setValues(requests);

sortShiftRequestsByDate();

let reply =
  "✅ " + requests.length + " shift request(s) recorded!\n\n";

for (const request of requests) {
  reply += request[2] + "  " + request[3] +
    "–" + request[4] + "\n";
}

reply += "\nStatus: Pending\nYour manager will confirm your shifts.";

if (errors.length > 0) {
  reply += "\n\n⚠️ These lines were skipped:\n" +
    errors.join("\n");
}

sendReply(token, event.replyToken, reply, sheet);
    }
  } catch (err) {
    if (sheet) {
      sheet.getRange("G2").setValue("ERROR: " + err.message);
    }
  }

  return HtmlService.createHtmlOutput("OK");
}

function findStaffRow(staffSheet, userId) {
  const lastRow = staffSheet.getLastRow();

  if (lastRow < 2) return 0;

  const ids = staffSheet.getRange(2, 1, lastRow - 1, 1)
    .getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(userId)) {
      return i + 2;
    }
  }

  return 0;
}

function sendReply(token, replyToken, message, sheet) {
  try {
    const response = UrlFetchApp.fetch(
      "https://api.line.me/v2/bot/message/reply",
      {
        method: "post",
        contentType: "application/json",
        headers: {
          Authorization: "Bearer " + token
        },
        payload: JSON.stringify({
          replyToken: replyToken,
          messages: [{ type: "text", text: message }]
        }),
        muteHttpExceptions: true
      }
    );

    sheet.getRange("G2").setValue(
      "LINE reply HTTP " + response.getResponseCode() +
      ": " + response.getContentText()
    );
  } catch (err) {
    sheet.getRange("G2").setValue(
      "REPLY ERROR: " + err.message
    );
  }
}


function sortShiftRequestsByDate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Shift Requests");

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 6);
  const rows = range.getValues();

  function dateKey(value) {
    // Handle real Google Sheets date values
    if (value instanceof Date && !isNaN(value.getTime())) {
      return (value.getMonth() + 1) * 100 + value.getDate();
    }

    // Handle text dates such as 11/10
    const match = String(value).trim()
      .match(/^(\d{1,2})\/(\d{1,2})(?:\/\d{4})?$/);

    if (!match) return Infinity;

    const month = Number(match[1]);
    const day = Number(match[2]);

    if (month < 1 || month > 12 || day < 1 || day > 31) {
      return Infinity;
    }

    return month * 100 + day;
  }

  rows.sort((a, b) => dateKey(a[2]) - dateKey(b[2]));
  range.setValues(rows);
}


function syncConfirmedSchedule() {
  // Original spreadsheet
  const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
  const source = sourceSS.getSheetByName("Shift Requests");
  const staffSheet = sourceSS.getSheetByName("Staff");

  // Destination spreadsheet
  const destinationId = "1r4bzsaPRTIsdJBtkyqjxaFugvkq7yoh-RiFHfvy437k";
  const destinationSS = SpreadsheetApp.openById(destinationId);

  const destination = destinationSS.getSheetByName("シフト確認表");
  const staffList = destinationSS.getSheetByName("Staff List");

  // Check all required tabs before making changes
  if (!source || !staffSheet || !destination || !staffList) {
    throw new Error("Required spreadsheet tab not found.");
  }

  // Collect approved shifts
  const output = [];
  const lastRow = source.getLastRow();

  if (lastRow >= 2) {
    const rows = source.getRange(2, 1, lastRow - 1, 6).getValues();

    for (const row of rows) {
      if (String(row[5]).trim() === "Approved") {
        output.push([row[2], row[3], row[4], row[1]]);
      }
    }
  }

  // Update シフト確認表, preserving row 1
  const oldLastRow = destination.getLastRow();

  if (oldLastRow >= 2) {
    destination.getRange(2, 1, oldLastRow - 1, 4)
      .clearContent();
  }

  if (output.length > 0) {
    destination.getRange(2, 1, output.length, 4)
      .setValues(output);

    destination.getRange(2, 1, output.length, 1)
      .setNumberFormat("MM/dd/yyyy");

    destination.getRange(2, 2, output.length, 2)
      .setNumberFormat("HH:mm");
  }

  // Collect registered staff names from original Staff tab
  const staffLastRow = staffSheet.getLastRow();

  const staffNames = staffLastRow >= 2
    ? staffSheet.getRange(2, 2, staffLastRow - 1, 1)
        .getValues()
        .flat()
        .map(name => String(name).trim())
        .filter(name => name !== "")
    : [];

  const uniqueNames = [...new Set(staffNames)];

  // Update Staff List column A only
  const oldStaffLastRow = staffList.getLastRow();

  if (oldStaffLastRow >= 2) {
    staffList.getRange(2, 1, oldStaffLastRow - 1, 1)
      .clearContent();
  }

  staffList.getRange("A1").setValue("Staff Name");

  if (uniqueNames.length > 0) {
    staffList.getRange(2, 1, uniqueNames.length, 1)
      .setValues(uniqueNames.map(name => [name]));
  }
}


