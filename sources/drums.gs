function populateDrumsEffectiveness() {
  var firstPlayerNameRow = 7;
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

  instructionsSheet.getRange(26, 2).setValue("");
  instructionsSheet.getRange(27, 2).setValue("");

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, 33, 7).clearContent().clearNote();
  if (darkMode)
    sheet.getRange(1, 1, 38, 10).setBackground("#d9d9d9");
  else
    sheet.getRange(1, 1, 38, 10).setBackground("white");
  sheet.getRange(2, 4, 1, 1).setBackground("#cccccc");

  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue().replace(/\s/g, "");
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var information = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 2), 1);
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
  var apiKeyString = "?translate=true&api_key=" + api_key;
  var baseUrl = "https://classic.warcraftlogs.com:443/v1/";
  var baseUrlFrontEnd = "https://classic.warcraftlogs.com/reports/"
  if (lang != "EN") {
    baseUrl = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com:443/v1/";
    baseUrlFrontEnd = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com/reports/";
  }

  var sectionToLookAtStart = 0;
  var sectionToLookAtEnd = 999999999999;
  var startEndString = "&start=" + sectionToLookAtStart.toString() + "&end=" + sectionToLookAtEnd.toString();

  var manualStartAndEnd = "";
  try {
    var manualStartAndEnd = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("startEndOptional", langKeys, langTrans, "", "", "", "")).findNext(), 1).getValue();
    if (manualStartAndEnd != null && manualStartAndEnd.toString().length > 0) {
      manualStartAndEnd = manualStartAndEnd.replace(" ", "");
      var startEndParts = manualStartAndEnd.split("-");
      sectionToLookAtStart = startEndParts[0];
      sectionToLookAtEnd = startEndParts[1];
      startEndString = "&start=" + sectionToLookAtStart.toString() + "&end=" + sectionToLookAtEnd.toString();
    }
  } catch {
    var sectionToLookAtStart = 0;
    var sectionToLookAtEnd = 999999999999;
    var startEndString = "&start=" + sectionToLookAtStart.toString() + "&end=" + sectionToLookAtEnd.toString();
  }

  var urlAllDrums = baseUrl + "report/events/buffs/" + logId + apiKeyString + startEndString + "&by=source";
  var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString;
  var allPlayersUrl = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndString;

  var baseSheetName = getStringForLang("drumsTab", langKeys, langTrans, "", "", "", "")
  var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));
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

  var allPlayersData = JSON.parse(UrlFetchApp.fetch(allPlayersUrl));
  const allPlayersByNameAsc = sortByProperty(sortByProperty(allPlayersData.entries, 'name'), "type");

  var playersFound = 0;
  allPlayersByNameAsc.forEach(function (playerByNameAsc, playerCountByNameAsc) {
    if ((playerByNameAsc.type == "Druid" || playerByNameAsc.type == "Hunter" || playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Priest" || playerByNameAsc.type == "Paladin" || playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Shaman" || playerByNameAsc.type == "Warlock" || playerByNameAsc.type == "Warrior") && playerByNameAsc.total > 20) {
      Utilities.sleep(30);
      var allDrumsBuffData = JSON.parse(UrlFetchApp.fetch(urlAllDrums + "&targetid=" + playerByNameAsc.id + "&filter=ability.id%20IN%20%2835478%2C35476%2C35475%2C351355%2C351358%2C351360%29"));
      var allEventsString = JSON.stringify(allDrumsBuffData.events);
      while (allDrumsBuffData.nextPageTimestamp != null && allDrumsBuffData.nextPageTimestamp > 0) {
        allDrumsBuffData = JSON.parse(UrlFetchApp.fetch(urlAllDrums.replace("&start=0", "&start=" + allDrumsBuffData.nextPageTimestamp) + "&targetid=" + playerByNameAsc.id + "&filter=ability.id%20IN%20%2835478%2C35476%2C35475%2C351355%2C351358%2C351360%29"));
        var additionalDrumsBuffDataString = JSON.stringify(allDrumsBuffData.events);
        allEventsString = allEventsString.substr(0, allEventsString.length - 1);
        allEventsString += "," + additionalDrumsBuffDataString.substr(1, additionalDrumsBuffDataString.length - 1);
      }
      var allEvents = JSON.parse(allEventsString);
      var allDrumsCastData = JSON.parse(UrlFetchApp.fetch(urlAllDrums.replace("/buffs/", "/casts/") + "&sourceid=" + playerByNameAsc.id + "&filter=ability.id%20IN%20%2835478%2C35476%2C35475%2C351355%2C351358%2C351360%29"));
      var noBuffsFoundWar = 0;
      var noBuffsFoundBattle = 0;
      var noBuffsFoundRestoration = 0;
      var nonGreaterWarDrumsUsed = 0;
      var nonGreaterBattleDrumsUsed = 0;
      var nonGreaterRestorationDrumsUsed = 0;
      allDrumsCastData.events.forEach(function (drumCast, drumCastCount) {
        if (drumCast.type == "cast") {
          if (drumCast.ability.guid == "35475")
            nonGreaterWarDrumsUsed++;
          else if (drumCast.ability.guid == "35476")
            nonGreaterBattleDrumsUsed++;
          else if (drumCast.ability.guid == "35478")
            nonGreaterRestorationDrumsUsed++;
          var buffFound = false;
          allEvents.forEach(function (drumsEvent, drumsEventCount) {
            if ((drumCast.timestamp >= drumsEvent.timestamp && drumCast.timestamp < drumsEvent.timestamp + 100) || (drumCast.timestamp < drumsEvent.timestamp && drumCast.timestamp > drumsEvent.timestamp - 100))
              buffFound = true;
          })
          if (!buffFound) {
            if (drumCast.ability.guid == "35475" || drumCast.ability.guid == "351360")
              noBuffsFoundWar++;
            else if (drumCast.ability.guid == "35476" || drumCast.ability.guid == "351355")
              noBuffsFoundBattle++;
            else if (drumCast.ability.guid == "35478" || drumCast.ability.guid == "351358")
              noBuffsFoundRestoration++;
          }
        }
      })
      var timesTotalDrummed = noBuffsFoundWar + noBuffsFoundRestoration + noBuffsFoundBattle;
      var timesTotalPlayersReceived = 0;
      var timesWarDrummed = 0;
      var timesWarPlayersReceived = 0;
      var timesBattleDrummed = 0;
      var timesBattlePlayersReceived = 0;
      var timesRestorationDrummed = 0;
      var timesRestorationPlayersReceived = 0;
      var timestampsInfoDone = [];
      var timestampsInfoDoneRemoved = [];
      allEvents.forEach(function (drumsEvent, drumsEventCount) {
        if (drumsEvent.type == "applybuff") {
          if (!timestampsInfoDone.indexOf(drumsEvent.timestamp) > -1 && !checkIfTimestampIsClose(drumsEvent.timestamp, drumsEvent.sourceID, timestampsInfoDone)) {
            timestampsInfoDone.push(drumsEvent.timestamp + '-' + drumsEvent.sourceID);
            timesTotalDrummed++;
            timesTotalPlayersReceived++;
            if (drumsEvent.ability.guid == "35475" || drumsEvent.ability.guid == "351360") {
              timesWarDrummed++;
              timesWarPlayersReceived++;
            } else if (drumsEvent.ability.guid == "35476" || drumsEvent.ability.guid == "351355") {
              timesBattleDrummed++;
              timesBattlePlayersReceived++;
            } else if (drumsEvent.ability.guid == "35478" || drumsEvent.ability.guid == "351358") {
              timesRestorationDrummed++;
              timesRestorationPlayersReceived++;
            }
          } else {
            timesTotalPlayersReceived++;
            if (drumsEvent.ability.guid == "35475" || drumsEvent.ability.guid == "351360")
              timesWarPlayersReceived++;
            else if (drumsEvent.ability.guid == "35476" || drumsEvent.ability.guid == "351355")
              timesBattlePlayersReceived++;
            else if (drumsEvent.ability.guid == "35478" || drumsEvent.ability.guid == "351358")
              timesRestorationPlayersReceived++;
          }
        }
      })
      allEvents.forEach(function (drumsEvent, drumsEventCount) {
        if (drumsEvent.type == "removebuff") {
          if (!checkIfTimestampIsClose(drumsEvent.timestamp, drumsEvent.sourceID, timestampsInfoDone)) {
            if (!timestampsInfoDone.indexOf(drumsEvent.timestamp) > -1 && !checkIfTimestampIsClose(drumsEvent.timestamp, drumsEvent.sourceID, timestampsInfoDoneRemoved)) {
              timestampsInfoDoneRemoved.push(drumsEvent.timestamp + '-' + drumsEvent.sourceID);
              timesTotalDrummed++;
              timesTotalPlayersReceived++;
              if (drumsEvent.ability.guid == "35475" || drumsEvent.ability.guid == "351360") {
                timesWarDrummed++;
                timesWarPlayersReceived++;
              } else if (drumsEvent.ability.guid == "35476" || drumsEvent.ability.guid == "351355") {
                timesBattleDrummed++;
                timesBattlePlayersReceived++;
              } else if (drumsEvent.ability.guid == "35478" || drumsEvent.ability.guid == "351358") {
                timesRestorationDrummed++;
                timesRestorationPlayersReceived++;
              }
            } else {
              timesTotalPlayersReceived++;
              if (drumsEvent.ability.guid == "35475" || drumsEvent.ability.guid == "351360")
                timesWarPlayersReceived++;
              else if (drumsEvent.ability.guid == "35476" || drumsEvent.ability.guid == "351355")
                timesBattlePlayersReceived++;
              else if (drumsEvent.ability.guid == "35478" || drumsEvent.ability.guid == "351358")
                timesRestorationPlayersReceived++;
            }
          }
        }
      })
      if (timesTotalDrummed > 0) {
        var range = sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn);
        range.setValue(playerByNameAsc.name);
        range.setBackground(getColourForPlayerClass(playerByNameAsc.type));
        if (timesBattleDrummed > 0) {
          shiftRangeByColumns(sheet, range, 1).setValue((timesBattleDrummed + noBuffsFoundBattle) + " (⌀ " + Math.round(timesBattlePlayersReceived * 100 / (timesBattleDrummed + noBuffsFoundBattle)) / 100 + ")");
          if (nonGreaterBattleDrumsUsed > 0)
            shiftRangeByColumns(sheet, range, 1).setNote(getStringForLang("notGreaterDrumsUsed", langKeys, langTrans, nonGreaterBattleDrumsUsed.toString(), "", "", "")).setBackground("#cccccc");
        }
        if (timesWarDrummed > 0) {
          shiftRangeByColumns(sheet, range, 2).setValue((timesWarDrummed + noBuffsFoundWar) + " (⌀ " + Math.round(timesWarPlayersReceived * 100 / (timesWarDrummed + noBuffsFoundWar)) / 100 + ")");
          if (nonGreaterWarDrumsUsed > 0)
            shiftRangeByColumns(sheet, range, 2).setNote(getStringForLang("notGreaterDrumsUsed", langKeys, langTrans, nonGreaterWarDrumsUsed.toString(), "", "", "")).setBackground("#cccccc");
        }
        if (timesRestorationDrummed > 0) {
          shiftRangeByColumns(sheet, range, 3).setValue((timesRestorationDrummed + noBuffsFoundRestoration) + " (⌀ " + Math.round(timesRestorationPlayersReceived * 100 / (timesRestorationDrummed + noBuffsFoundRestoration)) / 100 + ")");
          if (nonGreaterRestorationDrumsUsed > 0)
            shiftRangeByColumns(sheet, range, 3).setNote(getStringForLang("notGreaterDrumsUsed", langKeys, langTrans, nonGreaterRestorationDrumsUsed.toString(), "", "", "")).setBackground("#cccccc");
        }
        if ((noBuffsFoundBattle + noBuffsFoundWar + noBuffsFoundRestoration) > 0)
          shiftRangeByColumns(sheet, range, 4).setValue(noBuffsFoundBattle + noBuffsFoundWar + noBuffsFoundRestoration);
        shiftRangeByColumns(sheet, range, 5).setValue(timesTotalDrummed);
        shiftRangeByColumns(sheet, range, 6).setValue(Math.round(timesTotalPlayersReceived * 100 / timesTotalDrummed) / 100);
        playersFound++;
      }
    }
  })
}

function checkIfTimestampIsClose(timestamp, sourceID, timestampsInfoDone) {
  for (var i = 0, j = timestampsInfoDone.length; i < j; i++) {
    var timestampDone = timestampsInfoDone[i].split('-')[0];
    var sourceIDDone = timestampsInfoDone[i].split('-')[1];
    if (sourceIDDone == sourceID && timestamp < (Number(timestampDone) + 30100) && timestampDone <= timestamp)
      return true;
  }
  return false;
}