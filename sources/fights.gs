function populateFights() {
  var firstNameRow = 7;
  var firstNameColumn = 2;
  var firstNameColumnSecond = 10;
  var confSpreadSheet = SpreadsheetApp.openById('1pIbbPkn9i5jxyQ60Xt86fLthtbdCAmFriIpPSvmXiu0');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var instructionsSheet = ss.getSheetByName("Instructions");

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

  instructionsSheet.getRange(26, 2).setValue("");
  instructionsSheet.getRange(27, 2).setValue("");

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  sheet.getRange(firstNameRow, firstNameColumn, 145, 14).clearContent().setFontWeight("normal").setFontStyle("normal").setBorder(false, false, false, false, false, false, false, SpreadsheetApp.BorderStyle.SOLID);
  if (darkMode) {
    sheet.getRange(1, 1, 150, 15).setBackground("#d9d9d9");
  } else {
    sheet.getRange(1, 1, 150, 15).setBackground("white");
  }
  sheet.getRange(firstNameRow - 1, 2, 1, 5).setBackground("#cccccc");
  sheet.getRange(firstNameRow - 2, 7, 2, 2).setBackground("#cccccc");
  sheet.getRange(firstNameRow - 1, 10, 1, 5).setBackground("#cccccc");
  sheet.getRange(firstNameRow - 5, 3, 1, 4).setBackground("#cccccc");

  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue().replace(/\s/g, "");
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var reportPathOrIdFirst = shiftRangeByColumns(sheet, sheet.createTextFinder("^" + getStringForLang("pathFirstLog", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var reportPathOrIdSecond = shiftRangeByColumns(sheet, sheet.createTextFinder("^" + getStringForLang("pathSecondLog", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var speedrunZone = shiftRangeByColumns(sheet, sheet.createTextFinder("^" + getStringForLang("speedrunZone", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var speedrunZoneAbr = "";
  if (speedrunZone == "Karazhan" || speedrunZone == getStringForLang("Karazhan", langKeys, langTrans, "", "", "", ""))
    speedrunZoneAbr = "Kara";
  else if (speedrunZone == "Serpentshrine Cavern" || speedrunZone == getStringForLang("SerpentshrineCavern", langKeys, langTrans, "", "", "", ""))
    speedrunZoneAbr = "SSC";
  else if (speedrunZone == "Tempest Keep" || speedrunZone == getStringForLang("TempestKeep", langKeys, langTrans, "", "", "", ""))
    speedrunZoneAbr = "TK";
  else if (speedrunZone == "Mount Hyjal" || speedrunZone == getStringForLang("MountHyjal", langKeys, langTrans, "", "", "", ""))
    speedrunZoneAbr = "MH";
  else if (speedrunZone == "Black Temple" || speedrunZone == getStringForLang("BlackTemple", langKeys, langTrans, "", "", "", ""))
    speedrunZoneAbr = "BT";
  else if (speedrunZone == "Sunwell" || speedrunZone == getStringForLang("Sunwell", langKeys, langTrans, "", "", "", ""))
    speedrunZoneAbr = "SW";
  var information = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " 1$").useRegularExpression(true).findNext(), 2), 1);
  shiftRangeByColumns(sheet, information, 1).clearContent();
  var informationSecond = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " 2$").useRegularExpression(true).findNext(), 2), 1);
  shiftRangeByColumns(sheet, informationSecond, 1).clearContent();

  if (reportPathOrIdFirst.indexOf("<") < 0 && reportPathOrIdFirst.length > 0)
    reportPathOrId = reportPathOrIdFirst;

  var logId = "";
  var logIdSecond = "";
  if (reportPathOrId.indexOf("vanilla.warcraftlogs") > -1)
    if (reportPathOrId.indexOf("vanilla.warcraftlogs") > -1)
      SpreadsheetApp.getUi().alert(getStringForLang("vanillaExecution", langKeys, langTrans, "", "", "", ""));
  reportPathOrId = reportPathOrId.replace(".cn/", ".com/");
  reportPathOrIdSecond = reportPathOrIdSecond.replace(".cn/", ".com/");
  if (reportPathOrId.indexOf("classic.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("classic.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else if (reportPathOrId.indexOf("tbc.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("tbc.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else if (reportPathOrId.indexOf("fresh.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("fresh.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else
    logId = reportPathOrId;
  if (reportPathOrIdSecond.indexOf("<") < 0 && reportPathOrIdSecond.length > 0) {
    if (reportPathOrIdSecond.indexOf("classic.warcraftlogs.com/reports/") > -1)
      logIdSecond = reportPathOrIdSecond.split("classic.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
    else if (reportPathOrIdSecond.indexOf("tbc.warcraftlogs.com/reports/") > -1)
      logIdSecond = reportPathOrIdSecond.split("tbc.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
    else if (reportPathOrIdSecond.indexOf("fresh.warcraftlogs.com/reports/") > -1)
      logIdSecond = reportPathOrIdSecond.split("fresh.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
    else
      logIdSecond = reportPathOrIdSecond;
  }
  var apiKeyString = "?translate=true&api_key=" + api_key;
  var baseUrl = "https://classic.warcraftlogs.com:443/v1/";
  var baseUrlFrontEnd = "https://classic.warcraftlogs.com/reports/"
  if (lang != "EN") {
    baseUrl = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com:443/v1/";
    baseUrlFrontEnd = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com/reports/";
  }
  var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString;
  var urlAllFightsSecond = "";
  if (logIdSecond.length > 0)
    urlAllFightsSecond = baseUrl + "report/fights/" + logIdSecond + apiKeyString;

  var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));
  var allFightsDataSecond = "";
  if (logIdSecond.length > 0)
    allFightsDataSecond = JSON.parse(UrlFetchApp.fetch(urlAllFightsSecond));
  var baseSheetName = getStringForLang("fightsTab", langKeys, langTrans, "", "", "", "");
  if (speedrunZoneAbr != "")
    baseSheetName += speedrunZoneAbr;
  if (includeReportTitleInSheetNames.indexOf("yes") > -1) {
    if (logIdSecond.length > 0)
      baseSheetName += " " + allFightsData.title + " - " + allFightsDataSecond.title;
    else
      baseSheetName += " " + allFightsData.title;
  }
  try {
    sheet.setName(baseSheetName);
  } catch (err) {
    try {
      sheet.setName(baseSheetName + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", ""));
    } catch (err2) {
      try {
        sheet.setName(baseSheetName + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", ""));
      } catch (err3) {
        sheet.setName(baseSheetName + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", ""));
      }
    }
  }

  var returnVal = getRaidStartAndEnd(allFightsData, ss, baseUrl + "report/events/summary/" + logId + apiKeyString);
  var zonesFound = [];
  if (returnVal != null && returnVal.zonesFound != null)
    zonesFound = returnVal.zonesFound;
  var zoneTimesString = " (";
  if (zonesFound != null && zonesFound.length > 0) {
    zonesFound.forEach(function (raidZone, raidZoneCount) {
      zoneTimesString += raidZone[5] + " " + getStringForLang("in", langKeys, langTrans, "", "", "", "") + " ";
      if (raidZone[10] > 0) {
        zoneTimesString += getStringForTimeStamp(raidZone[10], true) + ", ";
      } else {
        zoneTimesString += getStringForTimeStamp(raidZone[2] - raidZone[1], true) + ", ";
      }
    })
    zoneTimesString = zoneTimesString.substr(0, zoneTimesString.length - 2);
    if (zoneTimesString.length > 0)
      sheet.getRange(information.getRow(), information.getColumn() + 1).setValue(allFightsData.title + zoneTimesString + ")");
    else
      sheet.getRange(information.getRow(), information.getColumn() + 1).setValue(allFightsData.title);
  } else
    SpreadsheetApp.getUi().alert(getStringForLang("noRaidZone", langKeys, langTrans, "", "", "", ""));

  if (logIdSecond.length > 0) {
    var returnValSecond = getRaidStartAndEnd(allFightsDataSecond, ss, baseUrl + "report/events/summary/" + logId + apiKeyString);
    var zonesFoundSecond = [];
    if (returnVal != null && returnValSecond.zonesFound != null)
      zonesFoundSecond = returnValSecond.zonesFound;
    zoneTimesString = " (";
    if (zonesFoundSecond != null && zonesFoundSecond.length > 0) {
      zonesFoundSecond.forEach(function (raidZone, raidZoneCount) {
        zoneTimesString += raidZone[5] + " in ";
        if (raidZone[10] > 0) {
          zoneTimesString += getStringForTimeStamp(raidZone[10], true) + ", ";
        } else {
          zoneTimesString += getStringForTimeStamp(raidZone[2] - raidZone[1], true) + ", ";
        }
      })
      zoneTimesString = zoneTimesString.substr(0, zoneTimesString.length - 2);
      if (zoneTimesString.length > 0)
        sheet.getRange(informationSecond.getRow(), informationSecond.getColumn() + 1).setValue(allFightsDataSecond.title + zoneTimesString + ")");
      else
        sheet.getRange(informationSecond.getRow(), informationSecond.getColumn() + 1).setValue(allFightsDataSecond.title);
    } else
      SpreadsheetApp.getUi().alert(getStringForLang("noRaidZone", langKeys, langTrans, "", "", "", ""));
  }

  var nameSet = false;
  allFightsData.fights.forEach(function (fight, fightCount) {
    if (fight.zoneName != null && fight.zoneName.length > 0 && !nameSet) {
      sheet.getRange(information.getRow() + 1, information.getColumn() + 1).setValue(fight.zoneName);
      nameSet = true;
    }
  })
  var nameSetSecond = false;
  if (logIdSecond.length > 0 && allFightsDataSecond.zone != null) {
    allFightsDataSecond.fights.forEach(function (fightSecond, fightSecondCount) {
      if (fightSecond.zoneName != null && fightSecond.zoneName.length > 0 && !nameSetSecond) {
        sheet.getRange(informationSecond.getRow() + 1, informationSecond.getColumn() + 1).setValue(fightSecond.zoneName);
        nameSetSecond = true;
      }
    })
  }
  if (allFightsData.zone != null && allFightsData.zone > 0 && (allFightsData.zone < 1007 || (allFightsData.zone >= 2000 && allFightsData.zone < 2007)))
    SpreadsheetApp.getUi().alert(getStringForLang("vanillaExecution", langKeys, langTrans, "", "", "", ""));
  else if (allFightsData.zone <= 0)
    SpreadsheetApp.getUi().alert(getStringForLang("zoneNotRecognized", langKeys, langTrans, "", "", "", ""));

  var dateString = "";
  if (lang == "DE" || lang == "RU")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "dd.MM.yyyy HH:mm:ss");
  else if (lang == "EN")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "MMMM dd, yyyy HH:mm:ss");
  else if (lang == "CN")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", 'yyyy年M月d日 HH:mm:ss');
  else if (lang == "FR")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "dd/MM/yyyy HH:mm:ss");
  sheet.getRange(information.getRow() + 2, information.getColumn() + 1).setValue(dateString);
  if (logIdSecond.length > 0) {
    var dateStringSecond = "";
    if (lang == "DE" || lang == "RU")
      dateStringSecond = Utilities.formatDate(new Date(allFightsDataSecond.start), "GMT+1", "dd.MM.yyyy HH:mm:ss");
    else if (lang == "EN")
      dateStringSecond = Utilities.formatDate(new Date(allFightsDataSecond.start), "GMT+1", "MMMM dd, yyyy HH:mm:ss");
    else if (lang == "CN")
      dateStringSecond = Utilities.formatDate(new Date(allFightsDataSecond.start), "GMT+1", 'yyyy年M月d日 HH:mm:ss');
    else if (lang == "FR")
      dateStringSecond = Utilities.formatDate(new Date(allFightsDataSecond.start), "GMT+1", "dd/MM/yyyy HH:mm:ss");
    sheet.getRange(informationSecond.getRow() + 2, informationSecond.getColumn() + 1).setValue(dateStringSecond);
  }

  var trashString = getStringForLang("trashAlone", langKeys, langTrans, "", "", "", "");

  var count = 0;
  if (zonesFound != null && zonesFound.length > 0) {
    var previousFightEnd = -1;
    var zoneEnd = 999999999999;
    if (speedrunZoneAbr != "")
      zoneEnd = -1;
    var zoneStart = -1;
    zonesFound.forEach(function (raidZone, raidZoneCount) {
      if (raidZone[5] == speedrunZoneAbr) {
        zoneStart = raidZone[1];
        zoneEnd = raidZone[2];
      }
    })
    var arr = [];
    allFightsData.fights.forEach(function (fight, fightCount) {
      var validFight = false;
      if ((fight.start_time >= zoneStart && fight.end_time <= zoneEnd) && (fight.end_time - fight.start_time > 4000)) {
        allFightsData.enemies.forEach(function (enemy, enemyCount) {
          enemy.fights.forEach(function (enemyFight, enemyFightCount) {
            if (enemyFight.id == fight.id && (enemy.type == "NPC" || enemy.type == "Boss")) {
              validFight = true;
              if (zoneStart == -1)
                zoneStart = fight.start_time;
            }
          })
        })
      }
      if (validFight) {
        var isBossFight = false;
        if (fight.boss != null && Number(fight.boss) > 0) {
          sheet.getRange(firstNameRow + count, firstNameColumn, 1, 7).setFontWeight("bold");
          isBossFight = true;
          if (fight.kill != null && fight.kill == true) {
            if (logIdSecond.length > 0) {
              sheet.getRange(firstNameRow + count, firstNameColumn + 5).setValue('=' + sheet.getRange(firstNameRow + count, firstNameColumn + 3).getA1Notation() + '-VLOOKUP(' + sheet.getRange(firstNameRow + count, firstNameColumn).getA1Notation() + ';SORT($J:$N;3;FALSE);4;FALSE)');
              sheet.getRange(firstNameRow + count, firstNameColumn + 6).setValue('=' + sheet.getRange(firstNameRow + count, firstNameColumn + 4).getA1Notation() + '-VLOOKUP(' + sheet.getRange(firstNameRow + count, firstNameColumn).getA1Notation() + ';SORT($J:$N;3;FALSE);5;FALSE)');
            }
          }
        }
        arr[arr.length] = [];
        if (isBossFight)
          arr[arr.length - 1].push(fight.name);
        else
          arr[arr.length - 1].push(fight.name + " (" + trashString + ")");
        if (previousFightEnd == -1) {
          arr[arr.length - 1].push("---");
        } else {
          arr[arr.length - 1].push(getStringForTimeStamp(fight.start_time - previousFightEnd - zoneStart, true));
        }
        arr[arr.length - 1].push(getStringForTimeStamp(fight.start_time - zoneStart, true));
        arr[arr.length - 1].push(getStringForTimeStamp(fight.end_time - fight.start_time, true));
        arr[arr.length - 1].push(getStringForTimeStamp(fight.end_time - zoneStart, true));
        previousFightEnd = fight.end_time - zoneStart;
        count++;
      }
    })
  }
  sheet.getRange(firstNameRow + count, firstNameColumn, 1, 7).setBorder(true, false, false, false, false, false, false, SpreadsheetApp.BorderStyle.SOLID)
  sheet.getRange(firstNameRow + count, firstNameColumn).setFontStyle("italic").setValue(getStringForLang("totalIdleTime", langKeys, langTrans, "", "", "", ""));
  sheet.getRange(firstNameRow + count, firstNameColumn + 1).setFontStyle("italic").setValue('=SUM(' + sheet.getRange(firstNameRow, firstNameColumn + 1).getA1Notation() + ':' + sheet.getRange(firstNameRow + count - 1, firstNameColumn + 1).getA1Notation() + ')');
  if (count > 0) {
    var rangeAll = sheet.getRange(firstNameRow, firstNameColumn, count, 5);
    rangeAll.setValues(arr);
  }

  if (logIdSecond.length > 0) {
    sheet.getRange(firstNameRow + count, firstNameColumn + 6).setFontStyle("italic").setValue('=' + sheet.getRange(firstNameRow + count, firstNameColumn + 1).getA1Notation() + '-VLOOKUP(' + sheet.getRange(firstNameRow + count, firstNameColumn).getA1Notation() + ';SORT($J:$N;3;FALSE);2;FALSE)');

    var countSecond = 0;
    if (zonesFoundSecond != null && zonesFoundSecond.length > 0) {
      var previousFightEndSecond = -1;
      var zoneEndSecond = 999999999999;
      var zoneStartSecond = -1;
      if (speedrunZoneAbr != "")
        zoneEndSecond = -1;
      zonesFoundSecond.forEach(function (raidZone, raidZoneCount) {
        if (raidZone[5] == speedrunZoneAbr) {
          zoneStartSecond = raidZone[1];
          zoneEndSecond = raidZone[2];
        }
      })
      var arrSecond = [];
      allFightsDataSecond.fights.forEach(function (fight, fightCount) {
        var validFight = false;
        if ((fight.start_time >= zoneStartSecond && fight.end_time <= zoneEndSecond) && (fight.end_time - fight.start_time > 4000)) {
          allFightsDataSecond.enemies.forEach(function (enemy, enemyCount) {
            enemy.fights.forEach(function (enemyFight, enemyFightCount) {
              if (enemyFight.id == fight.id && (enemy.type == "NPC" || enemy.type == "Boss")) {
                validFight = true;
                if (zoneStartSecond == -1)
                  zoneStartSecond = fight.start_time;
              }
            })
          })
        }
        if (validFight) {
          var isBossFight = false;
          if (fight.boss != null && Number(fight.boss) > 0) {
            sheet.getRange(firstNameRow + countSecond, firstNameColumnSecond, 1, 5).setFontWeight("bold");
            isBossFight = true;
          }
          arrSecond[arrSecond.length] = [];
          if (isBossFight)
            arrSecond[arrSecond.length - 1].push(fight.name);
          else
            arrSecond[arrSecond.length - 1].push(fight.name + " (" + trashString + ")");
          if (previousFightEndSecond == -1) {
            arrSecond[arrSecond.length - 1].push("---");
          } else {
            arrSecond[arrSecond.length - 1].push(getStringForTimeStamp(fight.start_time - previousFightEndSecond - zoneStartSecond, true));
          }
          arrSecond[arrSecond.length - 1].push(getStringForTimeStamp(fight.start_time - zoneStartSecond, true));
          arrSecond[arrSecond.length - 1].push(getStringForTimeStamp(fight.end_time - fight.start_time, true));
          arrSecond[arrSecond.length - 1].push(getStringForTimeStamp(fight.end_time - zoneStartSecond, true));
          previousFightEndSecond = fight.end_time - zoneStartSecond;
          countSecond++;
        }
      })
    }
    sheet.getRange(firstNameRow + countSecond, firstNameColumnSecond, 1, 5).setBorder(true, false, false, false, false, false, false, SpreadsheetApp.BorderStyle.SOLID)
    sheet.getRange(firstNameRow + countSecond, firstNameColumnSecond).setFontStyle("italic").setValue(getStringForLang("totalIdleTime", langKeys, langTrans, "", "", "", ""));
    sheet.getRange(firstNameRow + countSecond, firstNameColumnSecond + 1).setFontStyle("italic").setValue('=SUM(' + sheet.getRange(firstNameRow, firstNameColumnSecond + 1).getA1Notation() + ':' + sheet.getRange(firstNameRow + countSecond - 1, firstNameColumnSecond + 1).getA1Notation() + ')');

    if (countSecond > 0) {
      var rangeAllSecond = sheet.getRange(firstNameRow, firstNameColumnSecond, countSecond, 5);
      rangeAllSecond.setValues(arrSecond);
    }
  }
}