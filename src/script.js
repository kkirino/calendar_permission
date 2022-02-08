function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("サイドバーを開く")
    .addItem("開く", "openSidebar")
    .addToUi();
}

function openSidebar() {
  const htmlOutput = HtmlService.createTemplateFromFile("sidebar").evaluate();
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

const ss = SpreadsheetApp.getActiveSpreadsheet();

function getSettings() {
  const ws = ss.getSheetByName("settings");
  const titles = ws
    .getRange(2, 1, ws.getLastRow() - 1, 1)
    .getValues()
    .map(function (row) {
      return row[0];
    });
  const values = ws
    .getRange(2, 2, ws.getLastRow() - 1, 1)
    .getValues()
    .map(function (row) {
      return row[0];
    });
  const settings = {};
  settings.manualPageUrl = values[titles.indexOf("マニュアルページのURL")];
  return settings;
}

function getCalendarList() {
  const ws = ss.getSheetByName("calendar_list");
  const calendarList = ws
    .getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn())
    .getValues()
    .map(function (row) {
      return {
        title: row[0],
        id: row[1],
      };
    });
  return calendarList;
}

function getInfo(calendar) {
  const calendarInfo = Calendar.Acl.list(calendar.id).items.map(function (e) {
    return {
      title: calendar.title,
      id: calendar.id,
      email: e.id.replace("user:", ""),
      role: e.role,
    };
  });
  const ws = ss.getSheetByName("input");
  const sheetValues = ws
    .getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn())
    .getValues();
  const calendarTitles = ws
    .getRange(1, 1, 1, ws.getLastColumn())
    .getValues()[0];
  const calendarIndex = calendarTitles.indexOf(calendar.title);
  const sheetInfo = sheetValues.map(function (row) {
    return {
      title: calendar.title,
      id: calendar.id,
      email: row[1],
      role: row[calendarIndex],
    };
  });
  return {
    calendarInfo: calendarInfo,
    sheetInfo: sheetInfo,
  };
}

function reduceSum(previousValue, currentValue) {
  return previousValue + currentValue;
}

function removeOldCalendarPermission(info) {
  const calendarInfo = info.calendarInfo;
  const sheetInfo = info.sheetInfo;
  const str = "@group.calendar.google.com";
  const ws = ss.getSheetByName("log");
  calendarInfo.forEach(function (calendar) {
    const isCalendarEmailInSheet = sheetInfo
      .map(function (sheet) {
        return calendar.email === sheet.email;
      })
      .reduce(reduceSum);
    if (
      !calendar.email.includes(str) &&
      calendar.email !== "default" &&
      isCalendarEmailInSheet === 0
    ) {
      Calendar.Acl.insert(
        {
          role: "none",
          scope: {
            type: "user",
            value: calendar.email,
          },
        },
        calendar.id
      );
      ws.appendRow([
        new Date(),
        calendar.title + "で" + calendar.email + "の許可を削除",
      ]);
    }
  });
}

function updatePermissionBySheetInfo(info) {
  const calendarInfo = info.calendarInfo;
  const sheetInfo = info.sheetInfo;
  const ws = ss.getSheetByName("log");
  sheetInfo.forEach(function (sheet) {
    const sheetEmailBools = calendarInfo.map(function (calendar) {
      return calendar.email === sheet.email;
    });

    const isSheetEmailInCalendar = sheetEmailBools.reduce(reduceSum);

    if (
      isSheetEmailInCalendar == 0 ||
      calendarInfo[sheetEmailBools.indexOf(true)].role !== sheet.role
    ) {
      Calendar.Acl.insert(
        {
          role: sheet.role,
          scope: {
            type: "user",
            value: sheet.email,
          },
        },
        sheet.id
      );
      ws.appendRow([
        new Date(),
        sheet.title + "で" + sheet.email + "を" + sheet.role + "として登録",
      ]);
    }
  });
}

function runUpdatePermission(isChecked) {
  if (!isChecked) {
    return "チェックボックスにチェックを入れてください";
  } else {
    try {
      const calendarList = getCalendarList();
      calendarList.forEach(function (calendar) {
        removeOldCalendarPermission(getInfo(calendar));
        updatePermissionBySheetInfo(getInfo(calendar));
      });
      const now = new Date();
      const ws = ss.getSheetByName("input");
      ws.setName(now.toLocaleDateString() + " " + now.toLocaleTimeString());
      return "SUCCESS: カレンダー権限の更新が終わりました。";
    } catch {
      return "ERROR!!: スクリプトは正しく動作しませんでした。";
    }
  }
}
