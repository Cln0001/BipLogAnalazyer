function populateValidate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var instructionsSheet = ss.getSheetByName("Instructions");
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

  instructionsSheet.getRange(26, 2).setValue("");
  instructionsSheet.getRange(27, 2).setValue("");

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  sheet.getRange(6, 5, 29, 1).clearContent();
  sheet.getRange(8, 10, 7, 2).clearContent();
  sheet.getRange(10, 9, 1, 2).clearContent();
  sheet.getRange(10, 8, 1, 2).copyTo(sheet.getRange(10, 9, 1, 2), SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false);
  sheet.getRange(12, 9, 1, 2).clearContent();
  sheet.getRange(12, 8, 1, 2).clearContent();
  if (darkMode) {
    sheet.getRange(1, 1, 34, 11).setBackground("#d9d9d9");
    sheet.getRange(1, 1, 34, 2).setFontColor("#d9d9d9");
    sheet.getRange(1, 7, 34, 1).setFontColor("#d9d9d9");
    sheet.getRange(1, 11, 34, 1).setFontColor("#d9d9d9");
    sheet.getRange(2, 10, 1, 1).setFontColor("#d9d9d9");
  } else {
    sheet.getRange(1, 1, 34, 11).setBackground("white");
    sheet.getRange(1, 1, 34, 2).setFontColor("white");
    sheet.getRange(1, 7, 34, 1).setFontColor("white");
    sheet.getRange(1, 11, 34, 1).setFontColor("white");
    sheet.getRange(2, 10, 1, 1).setFontColor("white");
  }
  sheet.getRange(2, 4, 1, 1).setBackground("#cccccc");

  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue().replace(/\s/g, "");
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var information = addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 2);
  var stringl = getStringForLang("manualOverwriteZone", langKeys, langTrans, "", "", "", "");
  var validateZone = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("manualOverwriteZone", langKeys, langTrans, "", "", "", "")).useRegularExpression(false).findNext(), 1).getValue();
  var validateZoneAbr = "";
  if (validateZone.indexOf("Karazhan") > -1 || validateZone.indexOf(getStringForLang("Karazhan", langKeys, langTrans, "", "", "", "")) > - 1)
    validateZoneAbr = "Kara";
  if (validateZone.indexOf("TK") > -1 || validateZone.indexOf(getStringForLang("TK", langKeys, langTrans, "", "", "", "")) > - 1)
    validateZoneAbr = "TK";
  if (validateZone.indexOf("BT") > -1 || validateZone.indexOf(getStringForLang("BT", langKeys, langTrans, "", "", "", "")) > - 1)
    validateZoneAbr = "BT";
  if (validateZone.indexOf("Sunwell") > -1 || validateZone.indexOf(getStringForLang("Sunwell", langKeys, langTrans, "", "", "", "")) > - 1)
    validateZoneAbr = "Sunwell";
  shiftRangeByColumns(sheet, information, 1).clearContent();

  reportPathOrId = reportPathOrId.replace(".cn/", ".com/");
  var logId = "";
  if (reportPathOrId.indexOf("vanilla.warcraftlogs") > -1)
    SpreadsheetApp.getUi().alert(getStringForLang("vanillaExecution", langKeys, langTrans, "", "", "", ""));
  if (reportPathOrId.indexOf("classic.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("classic.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else if (reportPathOrId.indexOf("tbc.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("tbc.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else if (reportPathOrId.indexOf("fresh.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("fresh.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else
    logId = reportPathOrId;
  var startEndString = "&start=0&end=999999999999";
  var apiKeyString = "?translate=true&api_key=" + api_key;
  var baseUrl = "https://classic.warcraftlogs.com:443/v1/";
  if (lang != "EN") {
    baseUrl = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com:443/v1/";
    baseUrlFrontEnd = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com/reports/";
  }
  var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString;
  var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));

  var baseSheetName = getStringForLang("validateTab", langKeys, langTrans, "", "", "", "");
  if (includeReportTitleInSheetNames.indexOf("yes") > -1)
    baseSheetName += " " + allFightsData.title;
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

  var zoneId = allFightsData.zone;
  allFightsData.fights.forEach(function (fight, fightCount) {
    if (fight.zoneName != null && fight.zoneName.length > 0) {
      if (validateZoneAbr != "") {
        if (fight.zoneName.indexOf(validateZoneAbr) > -1) {
          sheet.getRange(information.getRow() + 1, information.getColumn() + 1).setValue(fight.zoneName);
          zoneId = fight.zoneID;
        }
      } else
        sheet.getRange(information.getRow() + 1, information.getColumn() + 1).setValue(fight.zoneName);
    }
  })
  if (allFightsData.zone != null && allFightsData.zone > 0 && (allFightsData.zone < 1007 || (allFightsData.zone >= 2000 && allFightsData.zone < 2007)))
    SpreadsheetApp.getUi().alert(getStringForLang("vanillaExecution", langKeys, langTrans, "", "", "", ""));
  else if (allFightsData.zone <= 0)
    SpreadsheetApp.getUi().alert(getStringForLang("zoneNotRecognized", langKeys, langTrans, "", "", "", ""));

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

    var dateString = "";
    if (lang == "DE" || lang == "RU")
      dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "dd.MM.yyyy HH:mm:ss");
    else if (lang == "EN")
      dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "MMMM dd, yyyy HH:mm:ss");
    else if (lang == "CN")
      dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "yyyy年M月d日 HH:mm:ss");
    else if (lang == "FR")
      dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "dd/MM/yyyy HH:mm:ss");
    sheet.getRange(information.getRow() + 2, information.getColumn() + 1).setValue(dateString);

    Utilities.sleep(1500);
    var confMobsToTrack = sheet.createTextFinder("^IDs$").useRegularExpression(true).findNext();
    var mobsToTrack = addRowsToRange(sheet, shiftRangeByRows(sheet, confMobsToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var dungeonsToTrack = shiftRangeByColumns(sheet, addRowsToRange(sheet, shiftRangeByRows(sheet, confMobsToTrack, 1), 200), 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    mobsToTrack.forEach(function (mob, mobCount) {
      var idCell = shiftRangeByRows(sheet, confMobsToTrack, 1 + mobCount);
      var dungeon = dungeonsToTrack[mobCount];
      var zoneEnd = -1;
      var zoneStart = -1;
      if (zonesFound != null && zonesFound.length > 0) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZone[5] == dungeon) {
            zoneStart = raidZone[1];
            zoneEnd = raidZone[2];
          }
        })
      }
      var amountCell = shiftRangeByColumns(sheet, idCell, 4);
      var ids = mob.toString().split(",");
      var idDeathCount = 0;
      ids.forEach(function (id, idCount) {
        allFightsData.enemies.forEach(function (enemyData, enemyDataCount) {
          if (enemyData.id != null && enemyData.id.toString().length > 0 && enemyData.guid != null && enemyData.guid.toString().length > 0 && id.toString() == enemyData.guid.toString()) {
            var urlDeathsTracked = baseUrl + "report/events/deaths/" + logId + apiKeyString + startEndString.replace("&start=0", "&start=" + zoneStart).replace("&end=999999999999", "&end=" + zoneEnd) + "&hostility=1&sourceid=";
            var deathsTrackedData = JSON.parse(UrlFetchApp.fetch(urlDeathsTracked + enemyData.id.toString()));
            Utilities.sleep(30);
            deathsTrackedData.events.forEach(function (enemy, enemyCount) {
              idDeathCount++;
            })
          }
        })
      })
      amountCell.setValue(idDeathCount);
    })

    var urlCharactersTracked = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndString.replace("&start=0", "&start=" + zonesFound[0][1]).replace("&end=999999999999", "&end=" + zonesFound[zonesFound.length - 1][2]) + "&encounter=-2";
    var charactersTrackedData = JSON.parse(UrlFetchApp.fetch(urlCharactersTracked));
    var charactersTracked = 0;
    if (charactersTrackedData != null && charactersTrackedData.entries != null && charactersTrackedData.entries.length > 0) {
      charactersTrackedData.entries.forEach(function (characterTracked, characterTrackedCount) {
        if (characterTracked != null && characterTracked.type != null && characterTracked.type.toString().length > 0) {
          if (characterTracked.type.toString() != "NPC" && characterTracked.type.toString() != "Boss")
            charactersTracked += 1;
        }
      })
    }
    var charactersCell = sheet.getRange(8, 10);
    charactersCell.setValue(charactersTracked);

    var confSpreadSheet = SpreadsheetApp.openById('1pIbbPkn9i5jxyQ60Xt86fLthtbdCAmFriIpPSvmXiu0');
    var validateConfigSheetKara = confSpreadSheet.getSheetByName("validateKaraLog");
    var validateConfigSheetSSCTK = confSpreadSheet.getSheetByName("validateSSCTKLog");
    var validateConfigSheetMHBT = confSpreadSheet.getSheetByName("validateMHBTLog");
    var validateConfigSheetZA = confSpreadSheet.getSheetByName("validateZALog");
    var validateConfigSheetSW = confSpreadSheet.getSheetByName("validateSWLog");

    var karaWCLzoneID = validateConfigSheetKara.getRange(2, validateConfigSheetKara.createTextFinder("Kara WCLzoneID").useRegularExpression(true).findNext().getColumn()).getValue();
    var sscWCLzoneID = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("SSC WCLzoneID").useRegularExpression(true).findNext().getColumn()).getValue();
    var tkWCLzoneID = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("TK WCLzoneID").useRegularExpression(true).findNext().getColumn()).getValue();
    var mhWCLzoneID = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("MH WCLzoneID").useRegularExpression(true).findNext().getColumn()).getValue();
    var btWCLzoneID = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("BT WCLzoneID").useRegularExpression(true).findNext().getColumn()).getValue();
    var zaWCLzoneID = validateConfigSheetZA.getRange(2, validateConfigSheetZA.createTextFinder("ZA WCLzoneID").useRegularExpression(true).findNext().getColumn()).getValue();

    if (zoneId.toString() == karaWCLzoneID) {
      var karaBosses = validateConfigSheetKara.getRange(2, validateConfigSheetKara.createTextFinder("Kara boss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
      var killedBosses = 0;
      allFightsData.fights.forEach(function (fight, fightCount) {
        if (fight.boss != null && Number(fight.boss) > 0 && fight.kill == true && karaBosses.indexOf(fight.boss) > -1)
          killedBosses += 1;
      })
      sheet.getRange(10, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("numberOfBossesKilledSingle", langKeys, langTrans, "10", "", "", ""));
      sheet.getRange(10, 10).setValue("'" + killedBosses);
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')>=10')
        .setBackground("#93c47d")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')>=0,VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')<10)')
        .setBackground("#ea9999")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rules = sheet.getConditionalFormatRules();
      rules.push(rule2);
      rules.push(rule);
      sheet.setConditionalFormatRules(rules);

      sheet.getRange(12, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("containsStartPoint", langKeys, langTrans, "", "", "", ""));

      var karaStartingPointFound = false;
      if (zonesFound != null && zonesFound.length > 0) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZone[5] == "Kara" && raidZone[6] == "true") {
            karaStartingPointFound = true;
          }
        })
      }
      if (karaStartingPointFound)
        sheet.getRange(12, 10).setValue(getStringForLang("yes", langKeys, langTrans, "", "", "", ""));
      else
        sheet.getRange(12, 10).setValue(getStringForLang("no", langKeys, langTrans, "", "", "", ""));
    } else if (zoneId.toString() == sscWCLzoneID || zoneId.toString() == tkWCLzoneID) {
      var sscBosses = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("SSC boss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
      var tkBosses = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("TK boss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
      var killedSSCBosses = 0;
      var killedTKBosses = 0;
      allFightsData.fights.forEach(function (fight, fightCount) {
        if (fight.boss != null && Number(fight.boss) > 0 && fight.kill == true) {
          if (sscBosses.indexOf(fight.boss) > -1)
            killedSSCBosses += 1;
          if (tkBosses.indexOf(fight.boss) > -1)
            killedTKBosses += 1;
        }
      })
      sheet.getRange(10, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("numberOfBossesKilledDouble", langKeys, langTrans, "6", "SSC", "4", "TK"));
      sheet.getRange(10, 10).setValue("SSC: " + killedSSCBosses + " - TK: " + killedTKBosses);

      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(REGEXMATCH(J10, "TK: 4"), REGEXMATCH(J10, "SSC: 6"))')
        .setBackground("#93c47d")
        .setRanges([sheet.getRange(10, 10)])
        .build();


      var rule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('SSC: 6')
        .setBackground("#fff2cc")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rule3 = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('TK: 4')
        .setBackground("#fff2cc")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rule4 = SpreadsheetApp.newConditionalFormatRule()
        .whenCellNotEmpty()
        .setBackground("#ea9999")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rules = sheet.getConditionalFormatRules();
      rules.push(rule);
      rules.push(rule2);
      rules.push(rule3);
      rules.push(rule4);
      sheet.setConditionalFormatRules(rules);

      var stringToPrint = "SSC: ";
      sheet.getRange(12, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("containsStartPoint", langKeys, langTrans, "", "", "", ""));

      var sscStartingPointFound = false;
      if (zonesFound != null && zonesFound.length > 0) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZone[5] == "SSC" && raidZone[6] == "true") {
            sscStartingPointFound = true;
          }
        })
      }
      var tkStartingPointFound = false;
      if (zonesFound != null && zonesFound.length > 0) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZone[5] == "TK" && raidZone[6] == "true") {
            tkStartingPointFound = true;
          }
        })
      }

      if (sscStartingPointFound)
        stringToPrint += getStringForLang("yes", langKeys, langTrans, "", "", "", "") + " - TK: ";
      else
        stringToPrint += getStringForLang("no", langKeys, langTrans, "", "", "", "") + " - TK: ";
      if (tkStartingPointFound)
        stringToPrint += getStringForLang("yes", langKeys, langTrans, "", "", "", "");
      else
        stringToPrint += getStringForLang("no", langKeys, langTrans, "", "", "", "");

      sheet.getRange(12, 10).setValue(stringToPrint);
    } else if (zoneId.toString() == mhWCLzoneID || zoneId.toString() == btWCLzoneID) {
      var mhBosses = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("MH boss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
      var btBosses = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("BT boss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
      var killedMHBosses = 0;
      var killedBTBosses = 0;
      allFightsData.fights.forEach(function (fight, fightCount) {
        if (fight.boss != null && Number(fight.boss) > 0 && fight.kill == true) {
          if (mhBosses.indexOf(fight.boss) > -1)
            killedMHBosses += 1;
          if (btBosses.indexOf(fight.boss) > -1)
            killedBTBosses += 1;
        }
      })
      sheet.getRange(10, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("numberOfBossesKilledDouble", langKeys, langTrans, "5", "MH", "9", "BT"));
      sheet.getRange(10, 10).setValue("MH: " + killedMHBosses + " - BT: " + killedBTBosses);

      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(REGEXMATCH(J10, "MH: 5"), REGEXMATCH(J10, "BT: 9"))')
        .setBackground("#93c47d")
        .setRanges([sheet.getRange(10, 10)])
        .build();


      var rule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('MH: 5')
        .setBackground("#fff2cc")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rule3 = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('BT: 9')
        .setBackground("#fff2cc")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rule4 = SpreadsheetApp.newConditionalFormatRule()
        .whenCellNotEmpty()
        .setBackground("#ea9999")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rules = sheet.getConditionalFormatRules();
      rules.push(rule);
      rules.push(rule2);
      rules.push(rule3);
      rules.push(rule4);
      sheet.setConditionalFormatRules(rules);

      var stringToPrint = "MH: ";
      sheet.getRange(12, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("containsStartPoint", langKeys, langTrans, "", "", "", ""));

      var mhStartingPointFound = false;
      if (zonesFound != null && zonesFound.length > 0) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZone[5] == "MH" && raidZone[6] == "true") {
            mhStartingPointFound = true;
          }
        })
      }
      var btStartingPointFound = false;
      if (zonesFound != null && zonesFound.length > 0) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZone[5] == "BT" && raidZone[6] == "true") {
            btStartingPointFound = true;
          }
        })
      }

      if (mhStartingPointFound)
        stringToPrint += getStringForLang("yes", langKeys, langTrans, "", "", "", "") + " - BT: ";
      else
        stringToPrint += getStringForLang("no", langKeys, langTrans, "", "", "", "") + " - BT: ";
      if (btStartingPointFound)
        stringToPrint += getStringForLang("yes", langKeys, langTrans, "", "", "", "");
      else
        stringToPrint += getStringForLang("no", langKeys, langTrans, "", "", "", "");

      sheet.getRange(12, 10).setValue(stringToPrint);
    } else if (zoneId.toString() == zaWCLzoneID) {
      var zaBosses = validateConfigSheetZA.getRange(2, validateConfigSheetZA.createTextFinder("ZA boss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
      var killedBosses = 0;
      allFightsData.fights.forEach(function (fight, fightCount) {
        if (fight.boss != null && Number(fight.boss) > 0 && fight.kill == true && zaBosses.indexOf(fight.boss) > -1)
          killedBosses += 1;
      })
      sheet.getRange(10, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("numberOfBossesKilledSingle", langKeys, langTrans, "6", "", "", ""));
      sheet.getRange(10, 10).setValue("'" + killedBosses);
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')>=6')
        .setBackground("#93c47d")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')>=0,VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')<6)')
        .setBackground("#ea9999")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rules = sheet.getConditionalFormatRules();
      rules.push(rule2);
      rules.push(rule);
      sheet.setConditionalFormatRules(rules);

      sheet.getRange(12, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("containsStartPoint", langKeys, langTrans, "", "", "", ""));

      var zaStartingPointFound = false;
      if (zonesFound != null && zonesFound.length > 0) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZone[5] == "ZA" && raidZone[6] == "true") {
            zaStartingPointFound = true;
          }
        })
      }
      if (zaStartingPointFound)
        sheet.getRange(12, 10).setValue(getStringForLang("yes", langKeys, langTrans, "", "", "", ""));
      else
        sheet.getRange(12, 10).setValue(getStringForLang("no", langKeys, langTrans, "", "", "", ""));
    } else if (zoneId.toString() == sscWCLzoneID) {
      var swBosses = validateConfigSheetSW.getRange(2, validateConfigSheetSW.createTextFinder("SW boss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
      var killedBosses = 0;
      allFightsData.fights.forEach(function (fight, fightCount) {
        if (fight.boss != null && Number(fight.boss) > 0 && fight.kill == true && swBosses.indexOf(fight.boss) > -1)
          killedBosses += 1;
      })
      sheet.getRange(10, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("numberOfBossesKilledSingle", langKeys, langTrans, "6", "", "", ""));
      sheet.getRange(10, 10).setValue("'" + killedBosses);
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')>=6')
        .setBackground("#93c47d")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')>=0,VALUE(' + sheet.getRange(10, 10).getA1Notation() + ')<6)')
        .setBackground("#ea9999")
        .setRanges([sheet.getRange(10, 10)])
        .build();

      var rules = sheet.getConditionalFormatRules();
      rules.push(rule2);
      rules.push(rule);
      sheet.setConditionalFormatRules(rules);

      sheet.getRange(12, 9).setFontWeight("bold").setHorizontalAlignment("right").setValue(getStringForLang("containsStartPoint", langKeys, langTrans, "", "", "", ""));

      var swStartingPointFound = false;
      if (zonesFound != null && zonesFound.length > 0) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZone[5] == "SW" && raidZone[6] == "true") {
            swStartingPointFound = true;
          }
        })
      }
      if (swStartingPointFound)
        sheet.getRange(12, 10).setValue(getStringForLang("yes", langKeys, langTrans, "", "", "", ""));
      else
        sheet.getRange(12, 10).setValue(getStringForLang("no", langKeys, langTrans, "", "", "", ""));
    }
  } else
    SpreadsheetApp.getUi().alert(getStringForLang("noRaidZone", langKeys, langTrans, "", "", "", ""));
}