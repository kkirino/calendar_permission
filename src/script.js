function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("サイドバーを開く")
    .addItem("開く", "openSidebar")
    .addToUi();
}

function openSidebar() {
  const htmlOutput =
    HtmlService.createHtmlOutputFromFile("sidebar").setTitle(
      "実行用コンソール"
    );
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

const ss = SpreadsheetApp.getActiveSpreadsheet();

function getCalendarList() {
  const sheet = ss.getSheetByName("calendar_list");
  const calendarList = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
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

  const sheet = ss.getSheetByName("input");
  const sheetValues = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getValues();
  const calendarTitles = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
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
      Logger.log(calendar.title + "で" + calendar.email + "の許可を消しました");
    }
  });
}

function updatePermissionBySheetInfo(info) {
  const calendarInfo = info.calendarInfo;
  const sheetInfo = info.sheetInfo;

  sheetInfo.forEach(function (sheet) {
    const sheetEmailBools = calendarInfo.map(function (calendar) {
      return calendar.email === sheet.email;
    });

    const isSheetEmailInCalendar = sheetEmailBools.reduce(reduceSum);

    if (
      isSheetEmailInCalendar == 0 || // 新規許可
      calendarInfo[sheetEmailBools.indexOf(true)].role !== sheet.role // 役割変更
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
      Logger.log(
        sheet.title +
          "で" +
          sheet.email +
          "を" +
          sheet.role +
          "として登録しました"
      );
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
