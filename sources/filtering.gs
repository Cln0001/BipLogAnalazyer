function generateRoleSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var instructionsSheet = ss.getSheetByName("Instructions");
  var baseSheetName = sheet.getName();

  var settings = ss.getSheetByName("settings");
  var webHook = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^5.$").useRegularExpression(true).findNext(), 4).getValue();

  var values = []; values[0] = []; values[0].push(""); values[1] = []; values[1].push("");
  sheet.getRange(3, 5, 2, 1).setValues(values);

  var confSpreadSheet = SpreadsheetApp.openById('1pIbbPkn9i5jxyQ60Xt86fLthtbdCAmFriIpPSvmXiu0');

  try { var lang = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^1.$").useRegularExpression(true).findNext(), 4).getValue(); } catch { }
  var langSheet = confSpreadSheet.getSheetByName("langTexts");
  var offset;
  if (lang != null && lang == "English") {
    lang = "EN";
    offset = 1;
  } else if (lang != null && lang == "Deutsch") {
    lang = "DE";
    offset = 2;
  } else if (lang != null && lang == "简体中文") {
    lang = "CN";
    offset = 3;
  } else if (lang != null && lang == "русский") {
    lang = "RU";
    offset = 4;
  } else if (lang != null && lang == "français") {
    lang = "FR";
    offset = 5;
  } else {
    lang = "EN";
    offset = 1;
  }
  var langKeys = langSheet.getRange(1, 1, 1000, 1).getValues().reduce(function (ar, e) { ar.push(e[0]); return ar; }, []);
  var langTrans = langSheet.getRange(1, 1 + offset, 1000, 1).getValues().reduce(function (ar, e) { ar.push(e[0]); return ar; }, []);

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  if ((baseSheetName.indexOf("All") > -1 || baseSheetName.indexOf(getStringForLang("all", langKeys, langTrans, "", "", "", "")) > -1) && sheet.getRange(4, 63).getValue().toString() == "done") {
    var sheets = ss.getSheets();
    for (var c = sheets.length - 1; c >= 0; c--) {
      var sheetNameSearch = sheets[c].getName();
      if (sheetNameSearch.indexOf("Caster") > -1 || sheetNameSearch.indexOf("Healer") > -1 || sheetNameSearch.indexOf("Physical") > -1 || sheetNameSearch.indexOf("Tank") > -1 || sheetNameSearch.indexOf("Filtered") > -1 || sheetNameSearch.indexOf(getStringForLang("Caster", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Healer", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Physical", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Tank", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("filtered", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("General", langKeys, langTrans, "", "", "", "")) > -1) {
        ss.deleteSheet(sheets[c]);
      }
    }

    var classHeaders = [];
    classHeaders.push("DeathKnights"); classHeaders.push("Druids"); classHeaders.push("Hunters"); classHeaders.push("Mages"); classHeaders.push("Priests"); classHeaders.push("Paladins"); classHeaders.push("Rogues"); classHeaders.push("Shamans"); classHeaders.push("Warlocks"); classHeaders.push("Warriors"); classHeaders.push(getStringForLang("DeathKnightPlur", langKeys, langTrans, "", "", "", "")); classHeaders.push(getStringForLang("DruidPlur", langKeys, langTrans, "", "", "", "")); classHeaders.push(getStringForLang("HunterPlur", langKeys, langTrans, "", "", "", "")); classHeaders.push(getStringForLang("MagePlur", langKeys, langTrans, "", "", "", "")); classHeaders.push(getStringForLang("PriestPlur", langKeys, langTrans, "", "", "", "")); classHeaders.push(getStringForLang("PaladinPlur", langKeys, langTrans, "", "", "", "")); classHeaders.push((getStringForLang("RoguePlur", langKeys, langTrans, "", "", "", ""))); classHeaders.push((getStringForLang("PriestPlur", langKeys, langTrans, "", "", "", ""))); classHeaders.push(getStringForLang("ShamanPlur", langKeys, langTrans, "", "", "", "")); classHeaders.push(getStringForLang("WarlockPlur", langKeys, langTrans, "", "", "", "")); classHeaders.push(getStringForLang("WarriorPlur", langKeys, langTrans, "", "", "", ""));

    var lineToSeparate = sheet.getRange(4, 64).getValue();
    var absorbStart = sheet.getRange(3, 64).getValue();

    var casterSheetname = baseSheetName.replace(getStringForLang("all", langKeys, langTrans, "", "", "", ""), getStringForLang("Caster", langKeys, langTrans, "", "", "", ""));
    var casterSheetnameCasts = casterSheetname + " - " + getStringForLang("casts", langKeys, langTrans, "", "", "", "");
    var healerSheetname = baseSheetName.replace(getStringForLang("all", langKeys, langTrans, "", "", "", ""), getStringForLang("Healer", langKeys, langTrans, "", "", "", ""));
    var healerSheetnameCasts = healerSheetname + " - " + getStringForLang("casts", langKeys, langTrans, "", "", "", "");
    var physicalSheetname = baseSheetName.replace(getStringForLang("all", langKeys, langTrans, "", "", "", ""), getStringForLang("Physical", langKeys, langTrans, "", "", "", ""));
    var physicalSheetnameCasts = physicalSheetname + " - " + getStringForLang("casts", langKeys, langTrans, "", "", "", "");
    var tankSheetname = baseSheetName.replace(getStringForLang("all", langKeys, langTrans, "", "", "", ""), getStringForLang("Tank", langKeys, langTrans, "", "", "", ""));
    var tankSheetnameCasts = tankSheetname + " - " + getStringForLang("casts", langKeys, langTrans, "", "", "", "");
    var filteredSheetname = nameSheetSafely(sheet.copyTo(ss), getStringForLang("all", langKeys, langTrans, "", "", "", "") + " " + getStringForLang("filtered", langKeys, langTrans, "", "", "", ""), langKeys, langTrans);
    var filteredSheet = ss.getSheetByName(filteredSheetname);
    filter(filteredSheetname, false, settings, ss, false, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    var filteredCastsSheetname = nameSheetSafely(sheet.copyTo(ss), getStringForLang("all", langKeys, langTrans, "", "", "", "") + " " + getStringForLang("casts", langKeys, langTrans, "", "", "", "") + " " + getStringForLang("filtered", langKeys, langTrans, "", "", "", ""), langKeys, langTrans);
    var filteredCastsSheet = ss.getSheetByName(filteredCastsSheetname);
    var generalSheetname = nameSheetSafely(sheet.copyTo(ss), getStringForLang("General", langKeys, langTrans, "", "", "", ""), langKeys, langTrans);
    makeGeneralSheet(generalSheetname, settings, ss, langKeys, langTrans, classHeaders);
    filter(filteredCastsSheetname, true, settings, ss, false, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    casterSheetname = nameSheetSafely(filteredSheet.copyTo(ss), casterSheetname, langKeys, langTrans);
    filter(casterSheetname, false, settings, ss, true, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    casterSheetnameCasts = nameSheetSafely(filteredCastsSheet.copyTo(ss), casterSheetnameCasts, langKeys, langTrans);
    filter(casterSheetnameCasts, true, settings, ss, true, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    healerSheetname = nameSheetSafely(filteredSheet.copyTo(ss), healerSheetname, langKeys, langTrans);
    filter(healerSheetname, false, settings, ss, true, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    healerSheetnameCasts = nameSheetSafely(filteredCastsSheet.copyTo(ss), healerSheetnameCasts, langKeys, langTrans);
    filter(healerSheetnameCasts, true, settings, ss, true, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    physicalSheetname = nameSheetSafely(filteredSheet.copyTo(ss), physicalSheetname, langKeys, langTrans);
    filter(physicalSheetname, false, settings, ss, true, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    physicalSheetnameCasts = nameSheetSafely(filteredCastsSheet.copyTo(ss), physicalSheetnameCasts, langKeys, langTrans);
    filter(physicalSheetnameCasts, true, settings, ss, true, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    tankSheetname = nameSheetSafely(filteredSheet.copyTo(ss), tankSheetname, langKeys, langTrans);
    filter(tankSheetname, false, settings, ss, true, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);
    tankSheetnameCasts = nameSheetSafely(filteredCastsSheet.copyTo(ss), tankSheetnameCasts, langKeys, langTrans);
    filter(tankSheetnameCasts, true, settings, ss, true, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart);

    var title = shiftRangeByColumns(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
    var zone = shiftRangeByColumns(sheet, sheet.createTextFinder("^" + getStringForLang("zone", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
    var date = shiftRangeByColumns(sheet, sheet.createTextFinder("^" + getStringForLang("date", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
    var type = sheet.getRange(4, 65).getValue();
    var newSpreadSheet = SpreadsheetApp.create(getStringForLang("RPBforParams", langKeys, langTrans, title, date, zone, type));

    var defaultSheetName = "";
    try { defaultSheetName = newSpreadSheet.getSheets()[0].getName(); } catch (e) { }

    try { ss.getSheetByName(generalSheetname).copyTo(newSpreadSheet).setName(generalSheetname); } catch (e) { }
    try { ss.getSheetByName(casterSheetname).copyTo(newSpreadSheet).setName(casterSheetname); } catch (e) { }
    try { ss.getSheetByName(casterSheetnameCasts).copyTo(newSpreadSheet).setName(casterSheetnameCasts); } catch (e) { }
    try { ss.getSheetByName(healerSheetname).copyTo(newSpreadSheet).setName(healerSheetname); } catch (e) { }
    try { ss.getSheetByName(healerSheetnameCasts).copyTo(newSpreadSheet).setName(healerSheetnameCasts); } catch (e) { }
    try { ss.getSheetByName(physicalSheetname).copyTo(newSpreadSheet).setName(physicalSheetname); } catch (e) { }
    try { ss.getSheetByName(physicalSheetnameCasts).copyTo(newSpreadSheet).setName(physicalSheetnameCasts); } catch (e) { }
    try { ss.getSheetByName(tankSheetname).copyTo(newSpreadSheet).setName(tankSheetname); } catch (e) { }
    try { ss.getSheetByName(tankSheetnameCasts).copyTo(newSpreadSheet).setName(tankSheetnameCasts); } catch (e) { }

    //Thanks to 0nimpulse#7741 for the help on the Discord integration!
    var sheet1 = "";
    if (defaultSheetName == "")
      sheet1 = "Sheet1";
    else
      sheet1 = defaultSheetName;
    try { newSpreadSheet.deleteSheet(newSpreadSheet.getSheetByName(sheet1)); } catch (e) { }
    try { DriveApp.getFileById(newSpreadSheet.getId()).moveTo(DriveApp.getFolderById(DriveApp.getFileById(ss.getId()).getParents().next().getId())); }
    catch (e) { DriveApp.getFileById(newSpreadSheet.getId()).moveTo(DriveApp.getRootFolder()); }

    var url = getPublicURLForSheet(newSpreadSheet);
    if (webHook != null && webHook.toString().length > 0)
      postMessageToDiscord(url, webHook, date, zone, title, type, langKeys, langTrans);

    sheets = ss.getSheets();
    for (var c = sheets.length - 1; c >= 0; c--) {
      var sheetNameSearch = sheets[c].getName();
      if (sheetNameSearch.indexOf("Caster") > -1 || sheetNameSearch.indexOf("Healer") > -1 || sheetNameSearch.indexOf("Physical") > -1 || sheetNameSearch.indexOf("Tank") > -1 || sheetNameSearch.indexOf("Filtered") > -1 || sheetNameSearch.indexOf(getStringForLang("Caster", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Healer", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Physical", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Tank", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("filtered", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("General", langKeys, langTrans, "", "", "", "")) > -1) {
        ss.deleteSheet(sheets[c]);
      }
    }

    var values = []; values[0] = []; values[0].push(getStringForLang("spreadsheetDone", langKeys, langTrans, "", "", "", "")); values[1] = []; values[1].push(url);
    sheet.getRange(3, 5, 2, 1).setValues(values);
  } else
    SpreadsheetApp.getUi().alert(getStringForLang("followInstructions", langKeys, langTrans, "", "", "", ""));
}

function makeGeneralSheet(generalSheetName, settings, ss, langKeys, langTrans, classHeaders) {
  var sheet = ss.getSheetByName(generalSheetName);

  var firstNameRow = 7;
  var firstNameColumn = 3;

  var drawings = sheet.getDrawings();
  drawings.forEach(function (drawingg, drawinggCount) {
    if (drawingg.getWidth() != 1) {
      drawingg.setWidth(1).setHeight(1);
    }
  })

  var lineToSeparate = sheet.getRange(3, 64).getValue();
  sheet.deleteRows(8, lineToSeparate - 8);

  var sheetColumns = sheet.getMaxColumns();

  var names = sheet.getRange(firstNameRow, firstNameColumn, 1, sheetColumns - firstNameColumn + 1).getValues();
  var numberOfColumnsHidden = 0;

  for (o = names[0].length - 1; o >= 0; o--) {
    var isClassColumn = false;
    var head = names[0][o];
    if (classHeaders.includes(head))
      isClassColumn = true;
    if (isClassColumn || sheet.isColumnHiddenByUser(firstNameColumn + o)) {
      numberOfColumnsHidden += 1;
    } else {
      if (numberOfColumnsHidden > 0) {
        sheet.deleteColumns(o + firstNameColumn + 1, numberOfColumnsHidden);
        numberOfColumnsHidden = 0;
      }
    }
  }
  if (numberOfColumnsHidden > 0) {
    try {
      sheet.deleteColumns(o + firstNameColumn + 1, numberOfColumnsHidden);
      numberOfColumnsHidden = 0;
    } catch (e) {
      ss.deleteSheet(sheet);
      return;
    }
  }

  var numberOfRowsHidden = 0;
  var sheetRows = sheet.getMaxRows();
  var values = sheet.getRange(firstNameRow + 2, firstNameColumn - 1, sheetRows - firstNameRow - 2, sheetColumns - firstNameColumn + 1).getValues();
  var rowHeadersWeight = sheet.getRange(firstNameRow + 2, firstNameColumn - 1, sheetRows - firstNameRow - 2, 1).getFontWeights();
  var overwrittens = sheet.getRange(firstNameRow + 2, 1, sheetRows - firstNameRow - 2, 1).getValues();
  var interruptedMergeLineBegins = 0;

  var dontHideRowsHeaderRange = settings.createTextFinder("^not hidden$").useRegularExpression(true).findNext();
  var dontHideRows = settings.getRange(dontHideRowsHeaderRange.getRow(), dontHideRowsHeaderRange.getColumn(), 1000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, [])

  for (var k = values.length - 1; k >= 0; k--) {
    if (values[k][0] == getStringForLang("namesAndSourcesInterruptedSpells", langKeys, langTrans, "", "", "", ""))
      interruptedMergeLineBegins = k;
  }
  for (var k = values.length - 1; k >= 0; k--) {
    var rowIsFilled = false;
    var hideOverwritten = false;
    if (overwrittens[k][0] == "yes")
      hideOverwritten = false;
    else if (overwrittens[k][0] == "no")
      hideOverwritten = true;
    var trimmedHeaderValue = values[k][0];
    if (trimmedHeaderValue.indexOf(" (" + getStringForLang("rank", langKeys, langTrans, "", "", "", "") + "") < 0)
      trimmedHeaderValue = trimmedHeaderValue.split(" (")[0].split(" [")[0];
    if (hideOverwritten) {
      if (!(dontHideRows.indexOf(trimmedHeaderValue) > -1))
        dontHideRows.push(trimmedHeaderValue);
    } else {
      for (t = 0, u = dontHideRows.length; t < u; t++) {
        if (dontHideRows[t] == trimmedHeaderValue) {
          dontHideRows.splice(t, 1);
          hiddenRemoved += 1;
        }
      }
    }
    if (rowHeadersWeight[k][0] != "bold" && !hideOverwritten && !(k >= interruptedMergeLineBegins && k <= (interruptedMergeLineBegins + 13))) {
      for (m = 1, n = values[k].length; m < n; m++) {
        if (values[k][m].toString() != "") {
          rowIsFilled = true;
          break;
        }
      }
      if (!rowIsFilled && trimmedHeaderValue == "") {
        if (k > 0) {
          for (w = 0, x = values[k - 1].length; w < x; w++) {
            if (values[k - 1][w].toString() != "") {
              rowIsFilled = true;
              break;
            }
          }
        }
        else
          rowIsFilled = true;
      }
    } else
      rowIsFilled = true;
    if (!rowIsFilled)
      numberOfRowsHidden += 1;
    else {
      if (numberOfRowsHidden > 0)
        sheet.deleteRows(k + firstNameRow + 3, numberOfRowsHidden);
      numberOfRowsHidden = 0;
    }
  }
  if (numberOfRowsHidden > 0)
    sheet.deleteRows(k + firstNameRow + 3, numberOfRowsHidden);
  sheet.deleteColumns(1, 1);
  sheet.deleteRows(1, 6);
}

function filter(sheetName, isCastsSheet, settings, ss, deleteColumnsAsWell, darkMode, langKeys, langTrans, classHeaders, lineToSeparate, absorbStart) {
  var firstNameRow = 7;
  var firstNameColumn = 3;

  var sheet = ss.getSheetByName(sheetName);

  if (!deleteColumnsAsWell) {
    if (isCastsSheet) {
      sheet.deleteRows(lineToSeparate, sheet.getMaxRows() - lineToSeparate);
    } else {
      sheet.deleteRows(absorbStart, sheet.getMaxRows() - absorbStart);
      sheet.deleteRows(8, lineToSeparate - 8);
    }
  }

  var dontHideRowsHeaderRange = settings.createTextFinder("^not hidden$").useRegularExpression(true).findNext();
  var dontHideRows = settings.getRange(dontHideRowsHeaderRange.getRow(), dontHideRowsHeaderRange.getColumn(), 1000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);

  var role = "";
  if (sheetName.indexOf("Caster") > -1 || sheetName.indexOf(getStringForLang("Caster", langKeys, langTrans, "", "", "", "")) > -1)
    role = "Caster";
  else if (sheetName.indexOf("Healer") > -1 || sheetName.indexOf(getStringForLang("Healer", langKeys, langTrans, "", "", "", "")) > -1)
    role = "Healer";
  else if (sheetName.indexOf("Physical") > -1 || sheetName.indexOf(getStringForLang("Physical", langKeys, langTrans, "", "", "", "")) > -1)
    role = "Physical";
  else if (sheetName.indexOf("Tank") > -1 || sheetName.indexOf(getStringForLang("Tank", langKeys, langTrans, "", "", "", "")) > -1)
    role = "Tank";
  else
    role = "All";

  var sheetColumns = sheet.getMaxColumns();

  if (!deleteColumnsAsWell) {
    var drawings = sheet.getDrawings();
    drawings.forEach(function (drawingg, drawinggCount) {
      if (drawingg.getWidth() != 1) {
        drawingg.setWidth(1).setHeight(1);
      }
    })
  }

  var rolesAndNames = sheet.getRange(firstNameRow - 2, firstNameColumn, 3, sheetColumns - firstNameColumn + 1).getValues();
  var headers1 = rolesAndNames[0];
  var headers2 = rolesAndNames[1];
  var names = rolesAndNames[2];
  var numberOfColumnsHidden = 0;
  var atLeastOnePlayerForThisRole = false;
  if (role == "All")
    atLeastOnePlayerForThisRole = true;
  else {
    for (o = headers1.length - 1; o >= 0; o--) {
      if ((names[o] != "") && ((headers1[o] != "" && headers1[o].indexOf(role) > -1) || (headers2[o] != "" && headers2[o].indexOf(role) > -1) || (headers1[o] != "" && headers1[o].indexOf(getStringForLang(role, langKeys, langTrans, "", "", "", "")) > -1) || (headers2[o] != "" && headers2[o].indexOf(getStringForLang(role, langKeys, langTrans, "", "", "", "")) > -1))) {
        atLeastOnePlayerForThisRole = true;
        break;
      }
    }
    if (!atLeastOnePlayerForThisRole) {
      ss.deleteSheet(sheet);
      return;
    }
  }
  for (o = headers1.length - 1; o >= 0; o--) {
    var isClassColumn = false;
    var header = names[o];
    if (classHeaders.includes(header))
      isClassColumn = true;
    if (isClassColumn && o > 2) {
      if (isCastsSheet)
        sheet.getRange(firstNameRow, firstNameColumn + o, sheet.getMaxRows(), 1).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
    if (((role != "All" && !(headers1[o].indexOf(role) > -1) && !(headers1[o].indexOf(getStringForLang(role, langKeys, langTrans, "", "", "", "")) > -1) && headers1[o].toString() != "") && (!(headers2[o].indexOf(role) > -1) && !(headers2[o].indexOf(getStringForLang(role, langKeys, langTrans, "", "", "", "")) > -1) && headers2[o].toString() != "")) || (!isCastsSheet && isClassColumn) || sheet.isColumnHiddenByUser(firstNameColumn + o) || names[o] == "") {
      numberOfColumnsHidden += 1;
    } else {
      if (numberOfColumnsHidden > 0) {
        sheet.deleteColumns(o + firstNameColumn + 1, numberOfColumnsHidden);
        numberOfColumnsHidden = 0;
      }
    }
  }
  if (numberOfColumnsHidden > 0) {
    try {
      sheet.deleteColumns(o + firstNameColumn + 1, numberOfColumnsHidden);
      numberOfColumnsHidden = 0;
    } catch (e) {
      ss.deleteSheet(sheet);
      return;
    }
  }

  if (isCastsSheet) {
    var rolesAndNamesNew = sheet.getRange(firstNameRow - 2, firstNameColumn, 3, sheet.getMaxColumns() - firstNameColumn + 1).getValues();
    var namesNew = rolesAndNamesNew[2];
    for (p = namesNew.length - 1; p >= 0; p--) {
      if ((p < namesNew.length - 1 && classHeaders.indexOf(namesNew[p]) > -1 && classHeaders.indexOf(namesNew[p + 1]) > -1) || (p == namesNew.length - 1 && classHeaders.indexOf(namesNew[p]) > -1)) {
        numberOfColumnsHidden += 1;
      } else {
        if (numberOfColumnsHidden > 0) {
          sheet.deleteColumns(p + firstNameColumn + 1, numberOfColumnsHidden);
          numberOfColumnsHidden = 0;
        }
      }
    }
    if (numberOfColumnsHidden > 0) {
      try {
        sheet.deleteColumns(p + firstNameColumn + 1, numberOfColumnsHidden);
      } catch (e) {
        ss.deleteSheet(sheet);
        return;
      }
    }
  }

  if (!isCastsSheet) {
    var numberOfRowsHidden = 0;
    var hiddenRemoved = 0;
    var sheetRows = sheet.getMaxRows();
    var values = sheet.getRange(firstNameRow + 2, firstNameColumn - 1, sheetRows - firstNameRow - 2, sheetColumns - firstNameColumn + 1).getValues();
    var rowHeadersWeight = sheet.getRange(firstNameRow + 2, firstNameColumn - 1, sheetRows - firstNameRow - 2, 1).getFontWeights();
    var overwrittens = sheet.getRange(firstNameRow + 2, 1, sheetRows - firstNameRow - 2, 1).getValues();
    var interruptedMergeLineBegins = 0;
    var damageTakenStartRow = sheet.createTextFinder(getStringForLang("damageTakenByTrackedAbilities", langKeys, langTrans, "", "", "", "")).useRegularExpression(true).findNext().getRow() - firstNameRow - 1;
    var damageTakenTotalRow = sheet.createTextFinder(getStringForLang("avoidableDamageTaken", langKeys, langTrans, "", "", "", "")).findNext().getRow() - firstNameRow - 1;

    for (var k = values.length - 1; k >= 0; k--) {
      if (values[k][0] == getStringForLang("namesAndSourcesInterruptedSpells", langKeys, langTrans, "", "", "", ""))
        interruptedMergeLineBegins = k;
    }
    for (var k = values.length - 1; k >= 0; k--) {
      var rowIsFilled = false;
      var hideOverwritten = false;
      if (overwrittens[k][0] == "yes")
        hideOverwritten = false;
      else if (overwrittens[k][0] == "no")
        hideOverwritten = true;
      var trimmedHeaderValue = values[k][0];
      if (trimmedHeaderValue.indexOf(" (" + getStringForLang("rank", langKeys, langTrans, "", "", "", "") + "") < 0)
        trimmedHeaderValue = trimmedHeaderValue.split(" (")[0].split(" [")[0];
      if (hideOverwritten) {
        if (!(dontHideRows.indexOf(trimmedHeaderValue) > -1))
          dontHideRows.push(trimmedHeaderValue);
      } else {
        for (t = 0, u = dontHideRows.length; t < u; t++) {
          if (dontHideRows[t] == trimmedHeaderValue) {
            dontHideRows.splice(t, 1);
            hiddenRemoved += 1;
          }
        }
      }
      if ((role == "Caster" || role == "Healer") && values[k][0].indexOf(getStringForLang("shoutUptimeOnYou", langKeys, langTrans, "", "", "", "")) > -1) {
        rowIsFilled = false;
      } else if (rowHeadersWeight[k][0] != "bold" && !hideOverwritten && !(k >= interruptedMergeLineBegins && k <= (interruptedMergeLineBegins + 13))) {
        for (m = 1, n = values[k].length; m < n; m++) {
          if (values[k][m].toString() != "") {
            rowIsFilled = true;
            break;
          }
        }
        if (!rowIsFilled && trimmedHeaderValue == "" && !(k > damageTakenStartRow && k < damageTakenTotalRow - 1)) {
          if (k > 0) {
            for (w = 0, x = values[k - 1].length; w < x; w++) {
              if (values[k - 1][w].toString() != "") {
                rowIsFilled = true;
                break;
              }
            }
          }
          else
            rowIsFilled = true;
        }
      } else
        rowIsFilled = true;
      if (!rowIsFilled)
        numberOfRowsHidden += 1;
      else {
        if (numberOfRowsHidden > 0)
          sheet.deleteRows(k + firstNameRow + 3, numberOfRowsHidden);
        numberOfRowsHidden = 0;
      }
    }
    if (numberOfRowsHidden > 0)
      sheet.deleteRows(k + firstNameRow + 3, numberOfRowsHidden);
  } else {
    var numberOfRowsHidden = 0;
    var sheetRows = sheet.getMaxRows();
    var values = sheet.getRange(firstNameRow + 2, firstNameColumn - 1, sheetRows - firstNameRow - 2, sheetColumns - firstNameColumn + 2).getValues();
    for (var k = values.length - 1; k >= 0; k--) {
      var rowIsFilled = false;
      for (m = 0, n = values[k].length; m < n; m++) {
        if (values[k][m].toString().length > 0) {
          rowIsFilled = true;
          break;
        }
        if (!rowIsFilled) {
          if (k > 0) {
            for (w = 0, x = values[k - 1].length; w < x; w++) {
              if (values[k - 1][w].toString().length > 0) {
                rowIsFilled = true;
                break;
              }
            }
          }
          else
            rowIsFilled = true;
        }
      }
      if (!rowIsFilled)
        numberOfRowsHidden += 1;
      else {
        if (numberOfRowsHidden > 0)
          sheet.deleteRows(k + firstNameRow + 3, numberOfRowsHidden);
        numberOfRowsHidden = 0;
      }
    }
  }

  if (deleteColumnsAsWell) {
    sheet.deleteColumns(1, 1);
    sheet.deleteRows(1, 6);
  }

  for (var q = 0; q < hiddenRemoved; q++) {
    dontHideRows.push("");
  }
  settings.getRange(dontHideRowsHeaderRange.getRow(), dontHideRowsHeaderRange.getColumn(), dontHideRows.length, 1).setValues(convertMultiRowSingleColumnArraytoMultidimensionalArray(dontHideRows));

  if (!isCastsSheet && deleteColumnsAsWell) {
    var namesNewBackColours = sheet.getRange(1, 2, 1, sheet.getMaxColumns() - 1).getBackgrounds()[0];
    var lastColour;
    for (s = namesNewBackColours.length - 1; s >= 0; s--) {
      if (lastColour != namesNewBackColours[s]) {
        sheet.getRange(1, 3 + s, sheet.getMaxRows(), 1).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
      }
      lastColour = namesNewBackColours[s];
    }
  }
  if (darkMode)
    sheet.getRange(1, 2, sheet.getMaxRows(), 1).setBorder(null, true, null, null, null, null, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
  else
    sheet.getRange(1, 2, sheet.getMaxRows(), 1).setBorder(null, true, null, null, null, null, "white", SpreadsheetApp.BorderStyle.SOLID);
}

function convertMultiRowSingleColumnArraytoMultidimensionalArray(array) {
  var returnArray = [];
  for (i = 0, j = array.length; i < j; i++) {
    addSingleEntryToMultiDimArray(returnArray, array[i]);
  }
  return returnArray;
}

function getPublicURLForSheet(sheet) {
  var file = DriveApp.getFileById(sheet.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function postMessageToDiscord(url, webHook, date, zone, title, type, langKeys, langTrans) {
  if (type == "")
    type = getStringForLang("trashAndBosses", langKeys, langTrans, "", "", "", "");
  else
    type = type.replace(")", "").replace(" (", "");
  if (type == "no wipes")
    type = getStringForLang("trashAndBosses", langKeys, langTrans, "", "", "", "") + " " + getStringForLang("noWipes", langKeys, langTrans, "", "", "", "");
  var payload = JSON.stringify({
    "username": getStringForLang("RolePerformanceBreakdownLong", langKeys, langTrans, "", "", "", ""),
    "avatar_url": "https://i.imgur.com/gLRG4ci.png",
    "embeds": [{
      "title": "\"" + title + "\"",
      "url": url,
      "color": 10783477,
      "fields": [
        {
          "name": getStringForLang("zone", langKeys, langTrans, "", "", "", ""),
          "value": zone,
          "inline": true
        },
        {
          "name": getStringForLang("type", langKeys, langTrans, "", "", "", ""),
          "value": type,
          "inline": true
        },
        {
          "name": getStringForLang("dateAndTime", langKeys, langTrans, "", "", "", ""),
          "value": date,
          "inline": true
        }
      ],
      "footer": {
        "text": getStringForLang("spreadsheetsBy", langKeys, langTrans, "", "", "", "") + " - https://discord.gg/nGvt5zH",
        "icon_url": "https://i.imgur.com/xopArYu.png"
      }
    }]
  });

  var params = {
    headers: {
      'Content-Type': 'application/json'
    },
    method: "POST",
    payload: payload,
    muteHttpExceptions: false
  };
  if (webHook.indexOf("$$$$$") > -1) {
    UrlFetchApp.fetch(webHook.split("$$$$$")[0], params);
    UrlFetchApp.fetch(webHook.split("$$$$$")[1], params);
  } else
    UrlFetchApp.fetch(webHook, params);
}