function exportSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var instructionsSheet = ss.getSheetByName("Instructions");

  instructionsSheet.getRange(26, 2).setValue("");
  instructionsSheet.getRange(27, 2).setValue("");

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

  var webHook = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^5.$").useRegularExpression(true).findNext(), 4).getValue();
  var exportGearIssues = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("gearIssuesTab", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var exportGearListing = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("gearListingTab", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var exportIgnites = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("drumsTab", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var exportValidateLog = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("validateTabWithInfo", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var exportShadowResi = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("shadowResiTab", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var exportFights = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("fightsTab", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();
  var exportBuffConsumables = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("buffConsumablesTab", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 1).getValue();

  var sheetsToConsider = [];
  if (exportGearIssues.indexOf("yes") > -1)
    sheetsToConsider.push(getStringForLang("gearIssuesTab", langKeys, langTrans, "", "", "", ""));
  if (exportGearListing.indexOf("yes") > -1)
    sheetsToConsider.push(getStringForLang("gearListingTab", langKeys, langTrans, "", "", "", ""));
  if (exportIgnites.indexOf("yes") > -1)
    sheetsToConsider.push(getStringForLang("drumsTab", langKeys, langTrans, "", "", "", ""));
  if (exportValidateLog.indexOf("yes") > -1)
    sheetsToConsider.push(getStringForLang("validateTab", langKeys, langTrans, "", "", "", ""));
  if (exportShadowResi.indexOf("yes") > -1)
    sheetsToConsider.push(getStringForLang("shadowResiTab", langKeys, langTrans, "", "", "", ""));
  if (exportFights.indexOf("yes") > -1)
    sheetsToConsider.push(getStringForLang("fightsTab", langKeys, langTrans, "", "", "", ""));
  if (exportBuffConsumables.indexOf("yes") > -1)
    sheetsToConsider.push(getStringForLang("buffConsumablesTab", langKeys, langTrans, "", "", "", ""));

  var defaultSheetName = "";
  var newSpreadSheet = null;
  var newSpreadSheetExists = false;
  var title = "";
  var zone = "";
  var date = "";
  var errorMessageShown = false;
  ss.getSheets().forEach(function (sheetToCheck, sheetToCheckCount) {
    for (var i = 0, j = sheetsToConsider.length; i < j; i++) {
      var sheetToCheckName = sheetToCheck.getName();
      if (sheetToCheckName.indexOf(sheetsToConsider[i]) > -1) {
        if (sheetToCheck.createTextFinder("^Player1").useRegularExpression(true).findNext() != null && sheetToCheck.createTextFinder("^Player1").useRegularExpression(true).findNext().getValue() == "Player1") {
          if (!errorMessageShown) {
            SpreadsheetApp.getUi().alert(getStringForLang("individualSheetsInfo", langKeys, langTrans, "", "", "", ""));
            errorMessageShown = true;
          }
        } else {
          if (!newSpreadSheetExists) {
            title = shiftRangeByColumns(sheetToCheck, sheetToCheck.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 1).getValue();
            if (title.length < 2)
              title = shiftRangeByColumns(sheetToCheck, sheetToCheck.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " 1$").useRegularExpression(true).findNext(), 1).getValue();
            zone = shiftRangeByColumns(sheetToCheck, sheetToCheck.createTextFinder("^" + getStringForLang("zone", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 1).getValue();
            if (zone.length < 2)
              zone = shiftRangeByColumns(sheetToCheck, sheetToCheck.createTextFinder("^" + getStringForLang("zone", langKeys, langTrans, "", "", "", "") + " 1$").useRegularExpression(true).findNext(), 1).getValue();
            date = shiftRangeByColumns(sheetToCheck, sheetToCheck.createTextFinder("^" + getStringForLang("date", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 1).getValue();
            if (date.length < 2)
              date = shiftRangeByColumns(sheetToCheck, sheetToCheck.createTextFinder("^" + getStringForLang("date", langKeys, langTrans, "", "", "", "") + " 1$").useRegularExpression(true).findNext(), 1).getValue();
            newSpreadSheet = SpreadsheetApp.create(getStringForLang("CLAforParams", langKeys, langTrans, title, date, zone, ""));
            try { defaultSheetName = newSpreadSheet.getSheets()[0].getName(); } catch (e) { }
            ss.getSheetByName("Instructions").copyTo(ss).setName("export Instructions");
            var newSheet = ss.getSheetByName("export Instructions");
            newSheet.getRange(9, 5).setValue("");
            ss.getSheetByName("export Instructions").copyTo(newSpreadSheet).setName("Instructions").hideSheet();
            ss.deleteSheet(ss.getSheetByName("export Instructions"));
            ss.getSheetByName("trans").copyTo(newSpreadSheet).setName("trans").hideSheet();
            newSpreadSheetExists = true;
          }
          try {
            if (!errorMessageShown) {
              ss.getSheetByName(sheetToCheckName).copyTo(ss).setName("export " + sheetToCheckName);
              var newSheet = ss.getSheetByName("export " + sheetToCheckName);
              newSheet.getDrawings().forEach(function (drawing, drawingCount) {
                if (drawing.getHeight() < 50 || drawing.getHeight() > 200)
                  drawing.remove();
              })
              if (sheetToCheckName.indexOf(getStringForLang("gearIssuesTab", langKeys, langTrans, "", "", "", "")) > -1) {
                newSheet.deleteColumns(1, 3);
              }
              ss.getSheetByName("export " + sheetToCheckName).copyTo(newSpreadSheet).setName(sheetToCheckName);
              ss.deleteSheet(ss.getSheetByName("export " + sheetToCheckName));
            }
          } catch (e) { }
        }
      }
    }
  })

  //Thanks to 0nimpulse#7741 for the help on the Discord integration!
  var sheet1 = "";
  if (defaultSheetName == "")
    sheet1 = "Sheet1";
  else
    sheet1 = defaultSheetName;
  try { newSpreadSheet.deleteSheet(newSpreadSheet.getSheetByName(sheet1)); } catch (e) { }
  try {
    DriveApp.getFileById(newSpreadSheet.getId()).moveTo(DriveApp.getFolderById(DriveApp.getFileById(ss.getId()).getParents().next().getId()));

    var url = getPublicURLForSheet(newSpreadSheet);
    if (webHook != null && webHook.toString().length > 0)
      postMessageToDiscord(url, webHook, date, zone, title, langKeys, langTrans);

    sheet.getRange(26, 2).setValue(getStringForLang("spreadsheetDone", langKeys, langTrans, "", "", "", ""));
    sheet.getRange(27, 2).setValue(url);
  } catch (e) { SpreadsheetApp.getUi().alert(getStringForLang("noSingleSheetsStarted", langKeys, langTrans, "", "", "", "") + ": " + e.message); }
}

function getPublicURLForSheet(sheet) {
  var file = DriveApp.getFileById(sheet.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function postMessageToDiscord(url, webHook, date, zone, title, langKeys, langTrans) {
  var payload = JSON.stringify({
    "username": getStringForLang("CombatLogAnalyticsLong", langKeys, langTrans, "", "", "", ""),
    "avatar_url": "https://i.imgur.com/0i6culm.png",
    "embeds": [{
      "title": "\"" + title + "\"",
      "url": url,
      "color": 10544871,
      "fields": [
        {
          "name": getStringForLang("zone", langKeys, langTrans, "", "", "", ""),
          "value": zone,
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

function getStringForTimeStamp(timeStamp, includeHours) {
  var delta = Math.abs(timeStamp) / 1000;
  var days = Math.floor(delta / 86400);
  delta -= days * 86400;
  var hours = Math.floor(delta / 3600) % 24;
  delta -= hours * 3600;
  var minutes = Math.floor(delta / 60) % 60;
  delta -= minutes * 60;
  var seconds = Math.floor(delta % 60);

  var secondsString = '';
  if (seconds < 10)
    secondsString = '0' + seconds.toString();
  else
    secondsString = seconds.toString();

  var minutesString = '';
  if (minutes < 10)
    minutesString = '0' + minutes.toString();
  else
    minutesString = minutes.toString();

  if (includeHours)
    return hours + ":" + minutesString + ":" + secondsString;
  else
    return minutesString + ":" + secondsString;
}

function getRaidStartAndEnd(allFightsData, ss, queryEnemy) {
  var confSpreadSheet = SpreadsheetApp.openById('1pIbbPkn9i5jxyQ60Xt86fLthtbdCAmFriIpPSvmXiu0');
  var validateConfigSheetKara = confSpreadSheet.getSheetByName("validateKaraLog");
  var validateConfigSheetSSCTK = confSpreadSheet.getSheetByName("validateSSCTKLog");
  var validateConfigSheetMHBT = confSpreadSheet.getSheetByName("validateMHBTLog");
  var validateConfigSheetZA = confSpreadSheet.getSheetByName("validateZALog");
  var validateConfigSheetSW = confSpreadSheet.getSheetByName("validateSWLog");
  var otherSheet = confSpreadSheet.getSheetByName("other");

  var queryEnemyFilled = false;
  if (queryEnemy != null && queryEnemy.length > 0) {
    queryEnemy = queryEnemy + "&hostility=1&sourceid=";
    queryEnemyFilled = true;
  }

  var zonesFound = [];

  var validZones = [];
  validZones.push(532); validZones.push(249); validZones.push(309); validZones.push(409); validZones.push(469); validZones.push(509); validZones.push(531); validZones.push(544); validZones.push(548); validZones.push(550); validZones.push(564); validZones.push(565); validZones.push(568); validZones.push(580); validZones.push(534); validZones.push(533);

  var karaZoneID = validateConfigSheetKara.getRange(2, validateConfigSheetKara.createTextFinder("Kara zoneID").useRegularExpression(true).findNext().getColumn()).getValue();
  var karaStartPoint = validateConfigSheetKara.getRange(2, validateConfigSheetKara.createTextFinder("Kara start point").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var karaEndbosses = validateConfigSheetKara.getRange(2, validateConfigSheetKara.createTextFinder("Kara endboss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var karaMobs = validateConfigSheetKara.getRange(2, validateConfigSheetKara.createTextFinder("Kara mobs").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var sscZoneID = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("SSC zoneID").useRegularExpression(true).findNext().getColumn()).getValue();
  var sscStartPoint = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("SSC start point").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var sscEndbosses = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("SSC endboss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var sscMobs = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("SSC mobs").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var tkZoneID = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("TK zoneID").useRegularExpression(true).findNext().getColumn()).getValue();
  var tkStartPoint = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("TK start point").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var tkEndbosses = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("TK endboss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var tkMobs = validateConfigSheetSSCTK.getRange(2, validateConfigSheetSSCTK.createTextFinder("TK mobs").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var mhZoneID = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("MH zoneID").useRegularExpression(true).findNext().getColumn()).getValue();
  var mhStartPoint = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("MH start point").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var mhEndbosses = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("MH endboss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var mhMobs = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("MH mobs").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var btZoneID = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("BT zoneID").useRegularExpression(true).findNext().getColumn()).getValue();
  var btStartPoint = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("BT start point").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var btEndbosses = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("BT endboss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var btMobs = validateConfigSheetMHBT.getRange(2, validateConfigSheetMHBT.createTextFinder("BT mobs").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var zaZoneID = validateConfigSheetZA.getRange(2, validateConfigSheetZA.createTextFinder("ZA zoneID").useRegularExpression(true).findNext().getColumn()).getValue();
  var zaStartPoint = validateConfigSheetZA.getRange(2, validateConfigSheetZA.createTextFinder("ZA start point").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var zaEndbosses = validateConfigSheetZA.getRange(2, validateConfigSheetZA.createTextFinder("ZA endboss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var zaMobs = validateConfigSheetZA.getRange(2, validateConfigSheetZA.createTextFinder("ZA mobs").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var swZoneID = validateConfigSheetSW.getRange(2, validateConfigSheetSW.createTextFinder("SW zoneID").useRegularExpression(true).findNext().getColumn()).getValue();
  var swStartPoint = validateConfigSheetSW.getRange(2, validateConfigSheetSW.createTextFinder("SW start point").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var swEndbosses = validateConfigSheetSW.getRange(2, validateConfigSheetSW.createTextFinder("SW endboss").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
  var swMobs = validateConfigSheetSW.getRange(2, validateConfigSheetSW.createTextFinder("SW mobs").useRegularExpression(true).findNext().getColumn(), 2000, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);

  var maxMillisecondsInfight = Number(otherSheet.getRange(1, 1).getValue());

  var atLeastOneStartPointFoundAfterXSecondsInfight = false;

  allFightsData.fights.forEach(function (fight, fightCount) {
    var raidZoneFound = -1;
    var zoneStart = -1;
    var zoneEnd = -1;
    var zoneStartRaw = -1;
    var zoneEndRaw = -1;
    zonesFound.forEach(function (raidZone, raidZoneCount) {
      if (fight.zoneID == raidZone[0]) {
        raidZoneFound = fight.zoneID;
        zoneStart = raidZone[1];
        zoneEnd = raidZone[2];
        zoneStartRaw = raidZone[3];
        zoneEndRaw = raidZone[4];
      }
    })
    if (raidZoneFound == -1) {
      zonesFound.forEach(function (raidZone, raidZoneCount) {
        allFightsData.enemies.forEach(function (enemy, enemyCount) {
          enemy.fights.forEach(function (enemyFight, enemyFightCount) {
            if (fight.id == enemyFight.id && (karaMobs.indexOf(enemy.guid) > -1 || sscMobs.indexOf(enemy.guid) > -1 || tkMobs.indexOf(enemy.guid) > -1 || mhMobs.indexOf(enemy.guid) > -1 || btMobs.indexOf(enemy.guid) > -1 || zaMobs.indexOf(enemy.guid) > -1 || swMobs.indexOf(enemy.guid) > -1)) {
              if ((karaMobs.indexOf(enemy.guid) > -1 && karaZoneID == raidZone[0]) || (sscMobs.indexOf(enemy.guid) > -1 && sscZoneID == raidZone[0]) || (tkMobs.indexOf(enemy.guid) > -1 && tkZoneID == raidZone[0]) || (mhMobs.indexOf(enemy.guid) > -1 && mhZoneID == raidZone[0]) || (btMobs.indexOf(enemy.guid) > -1 && btZoneID == raidZone[0]) || (zaMobs.indexOf(enemy.guid) > -1 && zaZoneID == raidZone[0]) || (swMobs.indexOf(enemy.guid) > -1 && swZoneID == raidZone[0])) {
                raidZoneFound = raidZone[0];
                zoneStart = raidZone[1];
                zoneEnd = raidZone[2];
                zoneStartRaw = raidZone[3];
                zoneEndRaw = raidZone[4];
              }
            }
          })
        })
      })
    }
    if (raidZoneFound == -1) {
      if (validZones.indexOf(fight.zoneID) > -1)
        raidZoneFound = fight.zoneID;
      else {
        allFightsData.enemies.forEach(function (enemy, enemyCount) {
          enemy.fights.forEach(function (enemyFight, enemyFightCount) {
            if (raidZoneFound == -1 && fight.id == enemyFight.id && (karaMobs.indexOf(enemy.guid) > -1 || sscMobs.indexOf(enemy.guid) > -1 || tkMobs.indexOf(enemy.guid) > -1 || mhMobs.indexOf(enemy.guid) > -1 || btMobs.indexOf(enemy.guid) > -1 || zaMobs.indexOf(enemy.guid) > -1 || swMobs.indexOf(enemy.guid) > -1)) {
              if (karaMobs.indexOf(enemy.guid) > -1)
                raidZoneFound = karaZoneID;
              else if (sscMobs.indexOf(enemy.guid) > -1)
                raidZoneFound = sscZoneID;
              else if (tkMobs.indexOf(enemy.guid) > -1)
                raidZoneFound = tkZoneID;
              else if (mhMobs.indexOf(enemy.guid) > -1)
                raidZoneFound = mhZoneID;
              else if (btMobs.indexOf(enemy.guid) > -1)
                raidZoneFound = btZoneID;
              else if (zaMobs.indexOf(enemy.guid) > -1)
                raidZoneFound = zaZoneID;
              else if (swMobs.indexOf(enemy.guid) > -1)
                raidZoneFound = swZoneID;
            }
          })
        })
      }
      if (raidZoneFound != -1) {
        zonesFound[zonesFound.length] = [];
        zonesFound[zonesFound.length - 1].push(raidZoneFound);
        zonesFound[zonesFound.length - 1].push(zoneStart);
        zonesFound[zonesFound.length - 1].push(zoneEnd);
        zonesFound[zonesFound.length - 1].push(zoneStartRaw);
        zonesFound[zonesFound.length - 1].push(zoneEndRaw);
        if (karaZoneID == raidZoneFound)
          zonesFound[zonesFound.length - 1].push("Kara");
        else if (sscZoneID == raidZoneFound)
          zonesFound[zonesFound.length - 1].push("SSC");
        else if (tkZoneID == raidZoneFound)
          zonesFound[zonesFound.length - 1].push("TK");
        else if (mhZoneID == raidZoneFound)
          zonesFound[zonesFound.length - 1].push("MH");
        else if (btZoneID == raidZoneFound)
          zonesFound[zonesFound.length - 1].push("BT");
        else if (zaZoneID == raidZoneFound)
          zonesFound[zonesFound.length - 1].push("ZA");
        else if (swZoneID == raidZoneFound)
          zonesFound[zonesFound.length - 1].push("SW");
        else {
          if (fight.zoneName != null && fight.zoneName.toString().length > 0)
            zonesFound[zonesFound.length - 1].push(fight.zoneName);
        }
        zonesFound[zonesFound.length - 1].push("false"); //startPointFound
        zonesFound[zonesFound.length - 1].push("false"); //endbossFound
        zonesFound[zonesFound.length - 1].push("false"); //firstBossFound
        zonesFound[zonesFound.length - 1].push("false"); //atLeastOneStartPointFoundAfterXSecondsInfight
        zonesFound[zonesFound.length - 1].push(0); //WCLTotalTime
        zonesFound[zonesFound.length - 1].push(0); //WCLPenaltyTime
      }
    }
    var startPointFoundStart = false;
    var startPointFoundEnd = false;
    var endbossFound = false;
    allFightsData.enemies.forEach(function (enemy, enemyCount) {
      enemy.fights.forEach(function (enemyFight, enemyFightCount) {
        if (enemyFight.id == fight.id && (enemy.type == "NPC" || enemy.type == "Boss")) {
          if ((raidZoneFound == karaZoneID && karaStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == sscZoneID && sscStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == tkZoneID && tkStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == mhZoneID && mhStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == btZoneID && btStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == zaZoneID && zaStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == swZoneID && swStartPoint.indexOf(enemy.guid) > -1)) {
            if (((enemy.guid == "21216") && fight.boss != null && fight.boss > 0) || enemy.guid != "21216") {
              startPointFoundStart = true;
            }
          } else if ((raidZoneFound == karaZoneID && karaStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == sscStartPoint && sscStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == tkZoneID && tkStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == mhZoneID && mhStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == btZoneID && btStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == zaZoneID && zaStartPoint.indexOf(enemy.guid) > -1) || (raidZoneFound == swZoneID && swStartPoint.indexOf(enemy.guid) > -1)) {
            if (queryEnemyFilled) {
              var queryEnemyData = JSON.parse(UrlFetchApp.fetch(queryEnemy + enemy.id.toString() + "&start=" + fight.start_time.toString() + "&end=" + (fight.start_time + maxMillisecondsInfight).toString()));
              if (queryEnemyData != null && queryEnemyData.events != null && queryEnemyData.events.length > 0)
                startPointFoundStart = true;
              else
                atLeastOneStartPointFoundAfterXSecondsInfight = true;
              Utilities.sleep(50);
            } else
              startPointFoundStart = true;
          }
        }
        if (fight.boss != null && Number(fight.boss) > 0 && fight.kill == true && (raidZoneFound == karaZoneID && karaEndbosses.indexOf(fight.boss) > -1) || (raidZoneFound == sscZoneID && sscEndbosses.indexOf(fight.boss) > -1) || (raidZoneFound == tkZoneID && tkEndbosses.indexOf(fight.boss) > -1) || (raidZoneFound == mhZoneID && mhEndbosses.indexOf(fight.boss) > -1) || (raidZoneFound == btZoneID && btEndbosses.indexOf(fight.boss) > -1) || (raidZoneFound == zaZoneID && zaEndbosses.indexOf(fight.boss) > -1) || (raidZoneFound == swZoneID && swEndbosses.indexOf(fight.boss) > -1))
          endbossFound = true;
      })
    })
    if (startPointFoundStart) {
      if (zoneStart == -1 || fight.start_time < zoneStart) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZoneFound == raidZone[0] && raidZone[8] == "false") {
            raidZone[1] = fight.start_time;
            raidZone[6] = "true";
          }
        })
      }
    } else if (startPointFoundEnd) {
      if (zoneStart == -1 || fight.end_time < zoneStart) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZoneFound == raidZone[0] && raidZone[8] == "false") {
            raidZone[1] = fight.end_time;
            raidZone[6] = "true";
          }
        })
      }
    } else {
      zonesFound.forEach(function (raidZone, raidZoneCount) {
        if (atLeastOneStartPointFoundAfterXSecondsInfight)
          raidZone[9] = "true";
      })
    }
    if (fight.boss != null && Number(fight.boss) > 0 && fight.kill != null && fight.kill.toString() == "true") {
      zonesFound.forEach(function (raidZone, raidZoneCount) {
        if (raidZoneFound == raidZone[0] && raidZone[8] == "false") {
          raidZone[8] = "true";
        }
      })
    }
    if (endbossFound) {
      if (zoneEnd == -1 || fight.end_time > zoneEnd) {
        zonesFound.forEach(function (raidZone, raidZoneCount) {
          if (raidZoneFound == raidZone[0]) {
            raidZone[2] = fight.end_time;
            raidZone[7] = "true";
          }
        })
      }
    }
  })
  zonesFound.forEach(function (raidZone, raidZoneCount) {
    allFightsData.fights.forEach(function (fight, fightCount) {
      if (validZones.indexOf(fight.zoneID) > -1) {
        if (fight.zoneID == raidZone[0] && (raidZone[3] == -1 || fight.start_time < raidZone[3]))
          raidZone[3] = fight.start_time;
      } else {
        allFightsData.enemies.forEach(function (enemy, enemyCount) {
          enemy.fights.forEach(function (enemyFight, enemyFightCount) {
            if (fight.id == enemyFight.id && (karaMobs.indexOf(enemy.guid) > -1 || sscMobs.indexOf(enemy.guid) > -1 || tkMobs.indexOf(enemy.guid) > -1 || mhMobs.indexOf(enemy.guid) > -1 || btMobs.indexOf(enemy.guid) > -1 || zaMobs.indexOf(enemy.guid) > -1 || swMobs.indexOf(enemy.guid) > -1)) {
              if (karaMobs.indexOf(enemy.guid) > -1 && (karaZoneID == raidZone[0] && (raidZone[3] == -1 || fight.start_time < raidZone[3])))
                raidZone[3] = fight.start_time;
              else if (sscMobs.indexOf(enemy.guid) > -1 && (sscZoneID == raidZone[0] && (raidZone[3] == -1 || fight.start_time < raidZone[3])))
                raidZone[3] = fight.start_time;
              else if (tkMobs.indexOf(enemy.guid) > -1 && (tkZoneID == raidZone[0] && (raidZone[3] == -1 || fight.start_time < raidZone[3])))
                raidZone[3] = fight.start_time;
              else if (mhMobs.indexOf(enemy.guid) > -1 && (mhZoneID == raidZone[0] && (raidZone[3] == -1 || fight.start_time < raidZone[3])))
                raidZone[3] = fight.start_time;
              else if (btMobs.indexOf(enemy.guid) > -1 && (btZoneID == raidZone[0] && (raidZone[3] == -1 || fight.start_time < raidZone[3])))
                raidZone[3] = fight.start_time;
              else if (zaMobs.indexOf(enemy.guid) > -1 && (zaZoneID == raidZone[0] && (raidZone[3] == -1 || fight.start_time < raidZone[3])))
                raidZone[3] = fight.start_time;
              else if (swMobs.indexOf(enemy.guid) > -1 && (swZoneID == raidZone[0] && (raidZone[3] == -1 || fight.start_time < raidZone[3])))
                raidZone[3] = fight.start_time;
            }
          })
        })
      }
    })
    if (raidZone[1] == -1) {
      raidZone[1] = raidZone[3];
    }

    allFightsData.fights.forEach(function (fight, fightCount) {
      if (validZones.indexOf(fight.zoneID) > -1) {
        if (fight.zoneID == raidZone[0] && (raidZone[4] == -1 || fight.end_time > raidZone[4]))
          raidZone[4] = fight.end_time;
      } else {
        allFightsData.enemies.forEach(function (enemy, enemyCount) {
          enemy.fights.forEach(function (enemyFight, enemyFightCount) {
            if (fight.id == enemyFight.id && (karaMobs.indexOf(enemy.guid) > -1 || sscMobs.indexOf(enemy.guid) > -1 || tkMobs.indexOf(enemy.guid) > -1 || mhMobs.indexOf(enemy.guid) > -1 || btMobs.indexOf(enemy.guid) > -1 || zaMobs.indexOf(enemy.guid) > -1 || swMobs.indexOf(enemy.guid) > -1)) {
              if (karaMobs.indexOf(enemy.guid) > -1 && (karaZoneID == raidZone[0] && (raidZone[4] == -1 || fight.end_time > raidZone[4])))
                raidZone[4] = fight.end_time;
              else if (sscMobs.indexOf(enemy.guid) > -1 && (sscZoneID == raidZone[0] && (raidZone[4] == -1 || fight.end_time > raidZone[4])))
                raidZone[4] = fight.end_time;
              else if (tkMobs.indexOf(enemy.guid) > -1 && (tkZoneID == raidZone[0] && (raidZone[4] == -1 || fight.end_time > raidZone[4])))
                raidZone[4] = fight.end_time;
              else if (mhMobs.indexOf(enemy.guid) > -1 && (mhZoneID == raidZone[0] && (raidZone[4] == -1 || fight.end_time > raidZone[4])))
                raidZone[4] = fight.end_time;
              else if (btMobs.indexOf(enemy.guid) > -1 && (btZoneID == raidZone[0] && (raidZone[4] == -1 || fight.end_time > raidZone[4])))
                raidZone[4] = fight.end_time;
              else if (zaMobs.indexOf(enemy.guid) > -1 && (zaZoneID == raidZone[0] && (raidZone[4] == -1 || fight.end_time > raidZone[4])))
                raidZone[4] = fight.end_time;
              else if (swMobs.indexOf(enemy.guid) > -1 && (swZoneID == raidZone[0] && (raidZone[4] == -1 || fight.end_time > raidZone[4])))
                raidZone[4] = fight.end_time;
            }
          })
        })
      }
    })
    if (raidZone[2] == -1) {
      raidZone[2] = raidZone[4];
    }
  })
  zonesFound.forEach(function (raidZone, raidZoneCount) {
    if (allFightsData.completeRaids != null) {
      allFightsData.completeRaids.forEach(function (completeRaid, completeRaidCount) {
        if (completeRaid.start_time == raidZone[1]) {
          raidZone[10] = completeRaid.end_time - completeRaid.start_time;
          var timePenalty = 0;
          if (completeRaid.missedTrashDetails != null) {
            completeRaid.missedTrashDetails.forEach(function (missedTrashDetail, missedTrashDetailCount) {
              if (missedTrashDetail.timePenalty != null && missedTrashDetail.timePenalty > 0)
                timePenalty += missedTrashDetail.timePenalty;
            })
          }
          raidZone[11] = timePenalty;
          if (raidZone[2] - raidZone[1] > raidZone[10])
            raidZone[2] = raidZone[1] + raidZone[10];
        }
      })
    }
  })
  return { zonesFound };
}

function toggleDarkMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var instructionsSheet = ss.getSheetByName("Instructions");

  var confSpreadSheet = SpreadsheetApp.openById('1pIbbPkn9i5jxyQ60Xt86fLthtbdCAmFriIpPSvmXiu0');

  var lang = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^1.$").useRegularExpression(true).findNext(), 4).getValue();
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
    var infoShownCellRange = shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 5);
    if (infoShownCellRange.getValue().indexOf("no") > -1) {
      infoShownCellRange.setValue("yes");
      SpreadsheetApp.getUi().alert(getStringForLang("toggleModeFirstInfo", langKeys, langTrans, "", "", "", ""));
    }
    var darkModeCellRange = shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4);
    var darkModeValue = darkModeCellRange.getValue();
    if (darkModeValue.indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  if (!darkMode) {
    darkModeCellRange.setValue("yes");
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
    darkModeCellRange.setFontColor("#d9d9d9");
    infoShownCellRange.setFontColor("#d9d9d9");
  } else {
    darkModeCellRange.setValue("no");
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground("white").setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
    darkModeCellRange.setFontColor("white");
    infoShownCellRange.setFontColor("white");
  }
  sheet.getRange(5, 5, 11, 1).setBackground("#fce5cd").setBorder(true, true, true, true, true, true, "#fce5cd", SpreadsheetApp.BorderStyle.SOLID).setFontColor("black");
  sheet.getRange(7, 5, 1, 1).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(9, 5, 1, 1).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(11, 5, 1, 1).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(13, 5, 1, 1).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(15, 5, 1, 1).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(25, 6, 7, 1).setBackground("#fce5cd").setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(25, 5, 7, 2).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function getStringForLang(key, langkeys, langTrans, param1, param2, param3, param4) {
  if (langkeys.indexOf(key) > -1)
    return langTrans[langkeys.indexOf(key)].replace("<param1>", param1).replace("<param2>", param2).replace("<param3>", param3).replace("<param4>", param4);
  else {
    return "missing/fehlend/manquant/失踪/отсутствует";
  }
}

function getColourForPlayerClass(playerClass) {
  if (playerClass == "Druid")
    return "#f6b26b";
  else if (playerClass == "Hunter")
    return "#b6d7a8";
  else if (playerClass == "Mage")
    return "#a4c2f4";
  else if (playerClass == "Paladin")
    return "#d5a6bd";
  else if (playerClass == "Priest")
    return "#efefef";
  else if (playerClass == "Rogue")
    return "#fff2cc";
  else if (playerClass == "Shaman")
    return "#6d9eeb";
  else if (playerClass == "Warlock")
    return "#b4a7d6";
  else if (playerClass == "Warrior")
    return "#e2d3c9";
}

function sortByProperty(objArray, prop) {
  if (arguments.length < 2) throw new Error("ARRAY, AND OBJECT PROPERTY MINIMUM ARGUMENTS, OPTIONAL DIRECTION");
  if (!Array.isArray(objArray)) throw new Error("FIRST ARGUMENT NOT AN ARRAY");
  const clone = objArray.slice(0);
  const direct = arguments.length > 2 ? arguments[2] : 1; //Default to ascending
  const propPath = (prop.constructor === Array) ? prop : prop.split(".");
  clone.sort(function (a, b) {
    for (let p in propPath) {
      if (a[propPath[p]] && b[propPath[p]]) {
        a = a[propPath[p]];
        b = b[propPath[p]];
      }
    }
    // convert numeric strings to integers
    a = a.toString().match(/^\d+$/) ? +a : a;
    b = b.toString().match(/^\d+$/) ? +b : b;
    return ((a < b) ? -1 * direct : ((a > b) ? 1 * direct : 0));
  });
  return clone;
}

function searchEntryForId(idArray, dataArray, index) {
  var count = 0;
  var returnvalue = "";
  idArray.forEach(function (id, idCount) {
    if (id.toString() == index.toString())
      returnvalue = dataArray[count];
    count++;
  })
  return returnvalue;
}

function addSingleEntryToMultiDimArray(multiArray, value) {
  multiArray[multiArray.length] = [];
  multiArray[multiArray.length - 1].push(value);
}

function addColumnsToRange(sheet, range, columnsToAdd) {
  return sheet.getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns() + columnsToAdd);
}

function addRowsToRange(sheet, range, rowsToAdd) {
  return sheet.getRange(range.getRow(), range.getColumn(), range.getNumRows() + rowsToAdd, range.getNumColumns());
}

function shiftRangeByColumns(sheet, range, columnsToShift) {
  return sheet.getRange(range.getRow(), range.getColumn() + columnsToShift, range.getNumRows(), range.getNumColumns());
}

function shiftRangeByRows(sheet, range, rowsToShift) {
  return sheet.getRange(range.getRow() + rowsToShift, range.getColumn(), range.getNumRows(), range.getNumColumns());
}