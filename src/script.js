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
const workingSheetName = "【作業中】新しい共有定義";

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

function createWorkingSheet() {
  const sheetNames = ss.getSheets().map(function (sheet) {
    return sheet.getName();
  });
  const sheetNamesWithDate = sheetNames.filter(function (string) {
    const regex = new RegExp(
      "^20[0-9]{2}/[0-9]{1,}/[0-9]{1,} [0-9]{1,}:[0-9]{1,}:[0-9]{1,}$"
    );
    return string.search(regex) > -1;
  });
  function compareDate(dateFirst, dateSecond) {
    return new Date(dateFirst).valueOf() - new Date(dateSecond).valueOf();
  }
  sheetNamesWithDate.sort(compareDate);
  const numSheet = sheetNamesWithDate.length;
  for (let i = 0; i < numSheet - 3; i++) {
    let sheetForDelete = ss.getSheetByName(sheetNamesWithDate[i]);
    ss.deleteSheet(sheetForDelete);
  }
  if (sheetNames.indexOf(workingSheetName) > -1) {
    return (
      "すでに「" +
      workingSheetName +
      "」シートが存在するため、新たなシートは作成しませんでした"
    );
  } else {
    const sheetForCopy = ss.getSheetByName(sheetNamesWithDate[numSheet - 1]);
    const copiedSheet = sheetForCopy.copyTo(ss);
    copiedSheet.setName(workingSheetName);
    return "「" + workingSheetName + "」シートを作成しました";
  }
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
  const ws = ss.getSheetByName(workingSheetName);
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
      const ws = ss.getSheetByName(workingSheetName);
      ws.setName(now.toLocaleDateString() + " " + now.toLocaleTimeString());
      const protection = ws.protect();
      protection.setWarningOnly(true);
      return "カレンダー権限の更新が終わりました。";
    } catch (e) {
      Logger.log(e);
      return "スクリプトは正しく動作しませんでした。";
    }
  }
}
