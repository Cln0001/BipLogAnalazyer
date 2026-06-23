function populateGearBreakdown() {
  var firstPlayerNameRow = 5;
  var firstPlayerNameColumn = 2;
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

  instructionsSheet.getRange(27, 2).setValue("");
  instructionsSheet.getRange(28, 2).setValue("");

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, 34, 18).clearContent();
  if (darkMode)
    sheet.getRange(1, 1, 38, 19).setBackground("#d9d9d9");
  else
    sheet.getRange(1, 1, 38, 19).setBackground("white");
  sheet.getRange(2, 7, 1, 1).setBackground("#cccccc");

  var baseUrl = "https://classic.warcraftlogs.com:443/v1/";
  var baseUrlFrontEnd = "https://classic.warcraftlogs.com/reports/"
  if (lang != "EN") {
    baseUrl = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com:443/v1/";
    baseUrlFrontEnd = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com/reports/";
  }
  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue().replace(/\s/g, "");
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var onlyFightNr = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("onlyFightId", langKeys, langTrans, "", "", "", "")).findNext(), 1).getValue();
  var information = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 2), 1);
  shiftRangeByColumns(sheet, information, 1).clearContent();
  sheet.getRange(firstPlayerNameRow - 1, firstPlayerNameColumn).clearContent();

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
  var apiKeyString = "?translate=true&api_key=" + api_key;

  var allPlayers = UrlFetchApp.fetch(baseUrl + "report/tables/casts/" + logId + apiKeyString + "&start=0&end=999999999999");
  var allPlayersData = JSON.parse(allPlayers);

  var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString;
  var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));
  var baseSheetName = getStringForLang("gearListingTab", langKeys, langTrans, "", "", "", "")
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

  var returnVal = getRaidStartAndEnd(allFightsData, ss, baseUrl + "report/events/summary/" + logId + apiKeyString);
  var zonesFound = [];
  if (returnVal != null && returnVal.zonesFound != null)
    zonesFound = returnVal.zonesFound;
  var zoneTimesString = " (";
  if (zonesFound != null && zonesFound.length > 0) {
    zonesFound.forEach(function (raidZone, raidZoneCount) {
      zoneTimesString += raidZone[5] + " in ";
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

  var nameSet = false;
  allFightsData.fights.forEach(function (fight, fightCount) {
    if (fight.zoneName != null && fight.zoneName.length > 0 && !nameSet) {
      sheet.getRange(information.getRow() + 1, information.getColumn() + 1).setValue(fight.zoneName);
      nameSet = true;
    }
  })
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
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "yyyy年M月d日 HH:mm:ss");
  else if (lang == "FR")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "dd/MM/yyyy HH:mm:ss");
  sheet.getRange(information.getRow() + 2, information.getColumn() + 1).setValue(dateString);

  var fightDataArr = [];
  var fightDataIndexArr = [];
  var playersDone = 0;
  var lastBossFightStart = 0;
  var lastBossFightId = "";
  allFightsData.fights.forEach(function (fight, fightRawCount) {
    if (onlyFightNr == null || onlyFightNr.toString().length == 0 || (onlyFightNr != null && onlyFightNr.toString().length > 0 && fight.id.toString() == onlyFightNr.toString())) {
      if (fight.boss != null && Number(fight.boss) > 0) {
        if (fight.start_time >= lastBossFightStart) {
          var fightData = searchEntryForId(fightDataIndexArr, fightDataArr, fight.id.toString());
          if (fightData == "") {
            Utilities.sleep(30);
            var urlSummaryPerFight = baseUrl + "report/tables/casts/" + logId + apiKeyString + "&start=" + fight.start_time + "&end=" + fight.end_time;
            fightData = JSON.parse(UrlFetchApp.fetch(urlSummaryPerFight));
            fightDataArr.push(fightData);
            fightDataIndexArr.push(fight.id.toString());
            if (fightData != null && fightData.entries != null) {
              fightData.entries.forEach(function (player, playerCount) {
                if (lastBossFightId != fight.id.toString() && player.gear != null && player.gear.length > 0) {
                  player.gear.forEach(function (item, itemCount) {
                    if (lastBossFightId != fight.id.toString() && item.id != null && item.id.toString().length > 0 && item.id.toString() != "0" && item.name != null && item.name.toString().length > 0) {
                      lastBossFightStart = fight.end_time - fight.start_time;
                      lastBossFightId = fight.id.toString();
                    }
                  })
                }
              })
            }
          }
        }
      }
    }
  })
  var rangeBoss = sheet.getRange(firstPlayerNameRow - 2, firstPlayerNameColumn);
  const allPlayersByNameAsc = sortByProperty(sortByProperty(allPlayersData.entries, "name"), "type");
  allPlayersByNameAsc.forEach(function (playerByNameAsc, playerCountByNameAsc) {
    var fightCount = 0;
    allFightsData.fights.forEach(function (fight, fightRawCount) {
      if (fight.id.toString() == lastBossFightId) {
        if (fight.kill == true)
          rangeBoss.setValue(fight.name + " (" + getStringForLang("killInTime", langKeys, langTrans, Math.round((fight.end_time - fight.start_time) / 1000), "", "", "") + "s)");
        else
          rangeBoss.setValue(fight.name + " (" + Math.round(Number(fight.fightPercentage) / 100) + "% " + getStringForLang("wipeAfterTime", langKeys, langTrans, Math.round((fight.end_time - fight.start_time) / 1000), "", "", "") + "s)");
        rangeBoss.setFontWeight("bold").setHorizontalAlignment("center");

        var fightData = searchEntryForId(fightDataIndexArr, fightDataArr, fight.id.toString());
        if (fightData == "") {
          var urlSummaryPerFight = baseUrl + "report/tables/casts/" + logId + apiKeyString + "&start=" + fight.start_time + "&end=" + fight.end_time;
          fightData = JSON.parse(UrlFetchApp.fetch(urlSummaryPerFight));
          fightDataArr.push(fightData);
          fightDataIndexArr.push(fight.id.toString());
          Utilities.sleep(30);
        }
        fightData.entries.forEach(function (player, playerCount) {
          if (playerByNameAsc.name == player.name) {
            if (player.gear != null && player.gear.length > 0) {
              player.gear.forEach(function (item, itemCount) {
                if (item.id != null && item.id.toString().length > 0 && item.id.toString() != "0" && item.slot != 3 && item.slot != 18) {
                  var itemPos = 0;
                  if (item.slot == 0 || item.slot == 1 || item.slot == 2 || item.slot == 4 || item.slot == 10 || item.slot == 11 || item.slot == 12 || item.slot == 13)
                    itemPos = item.slot;
                  else if (item.slot == 5)
                    itemPos = 7;
                  else if (item.slot == 6)
                    itemPos = 8;
                  else if (item.slot == 7)
                    itemPos = 9;
                  else if (item.slot == 8)
                    itemPos = 5;
                  else if (item.slot == 9)
                    itemPos = 6;
                  else if (item.slot == 14)
                    itemPos = 3;
                  else if (item.slot == 15)
                    itemPos = 14;
                  else if (item.slot == 16)
                    itemPos = 15;
                  else if (item.slot == 17)
                    itemPos = 16;
                  sheet.getRange(playersDone + firstPlayerNameRow, firstPlayerNameColumn + 1 + fightCount + itemPos).setValue(item.name);
                  /*if (item.permanentEnchantName != null && item.permanentEnchantName.length > 0)
                    sheet.getRange(playersDone + firstPlayerNameRow, firstPlayerNameColumn + 2 + fightCount + itemPos*3).setValue(item.permanentEnchantName);
                  if (item.gems != null && item.gems.length > 0) {
                    var gemString = "";
                    item.gems.forEach(function (gem, gemCount) {
                      if (gem.itemLevel != null) {
                        gemString += gem.id + ", ";
                      }
                    })
                    sheet.getRange(playersDone + firstPlayerNameRow, firstPlayerNameColumn + 3 + fightCount + itemPos*3).setValue(gemString.substr(0, gemString.length-2));
                  }*/
                }
              })
            }
            if (fightCount == 0) {
              var range = sheet.getRange(playersDone + firstPlayerNameRow, firstPlayerNameColumn);
              range.setValue(player.name);
              range.setBackground(getColourForPlayerClass(player.type));
              playersDone++;
            }
            fightCount++;
          }
        })
      }
    })
  })
}