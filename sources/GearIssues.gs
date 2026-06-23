function populateGearIssues() {
  var codeVersion = '1.6.0';
  var confSpreadSheet = SpreadsheetApp.openById('1pIbbPkn9i5jxyQ60Xt86fLthtbdCAmFriIpPSvmXiu0');
  var currentVersion = confSpreadSheet.getSheetByName("currentVersion").getRange(1, 1).getValue();

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

  var suboptimalSheet = confSpreadSheet.getSheetByName("suboptimalEnchants");
  var suboptimalItems = suboptimalSheet.getRange(2, 48, 1000, 1).getValues().reduce(function (ar, e) { ar.push(e[0]); return ar; }, []);

  if (currentVersion.indexOf(codeVersion) < 0) {
    SpreadsheetApp.getUi().alert(getStringForLang("sheetOutdated", langKeys, langTrans, "", "", "", ""));
  }

  var firstPlayerNameRow = 5;
  var firstPlayerNameColumn = 8;

  instructionsSheet.getRange(26, 2).setValue("");
  instructionsSheet.getRange(27, 2).setValue("");

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, 145, 30).clearContent().clearNote();
  if (darkMode) {
    sheet.getRange(1, 1, 150, 38).setBackground("#d9d9d9");
  } else
    sheet.getRange(1, 1, 150, 38).setBackground("white");

  var badIdsRaw = sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn - 6, 200, 1).getValues();
  var badIdsRawStyle = sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn - 6, 200, 1).getFontStyles();
  var counter = 0;
  var badIds = badIdsRaw.reduce(function (ar, e) {
    if (e[0]) {
      if (badIdsRawStyle[counter][0] != "italic")
        ar.push(e[0])
    }
    counter++;
    return ar;
  }, []);

  var badIdNamesRaw = sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn - 5, 200, 1).getValues();
  var badIdNamesStyle = sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn - 5, 200, 1).getFontStyles();
  counter = 0;
  var italicMessageShown = false;
  var badIdNames = badIdNamesRaw.reduce(function (ar, e) {
    if (e[0]) {
      if (badIdsRawStyle[counter][0] != "italic")
        ar.push(e[0])
      if (badIdsRawStyle[counter][0] != badIdNamesStyle[counter][0] && !italicMessageShown) {
        SpreadsheetApp.getUi().alert(getStringForLang("wrongItalic", langKeys, langTrans, "", "", "", ""));
        italicMessageShown = true;
      }
    }
    counter++;
    return ar;
  }, []);

  var gearToIgnoreRaw = sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn - 3, 200, 1).getValues();
  var gearToIgnoreRawStyles = sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn - 3, 200, 1).getFontStyles();
  counter = 0
  var gearToIgnore = gearToIgnoreRaw.reduce(function (ar, e) {
    if (e[0]) {
      if (gearToIgnoreRawStyles[counter][0] != "italic")
        ar.push(e[0])
    }
    counter++;
    return ar;
  }, []);

  var gemSheet = ss.getSheetByName("sockets");
  var gemItemIdsRaw = gemSheet.getRange(2, 1, 2000, 1).getValues();
  var gemItemIds = gemItemIdsRaw.reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
  }, []);

  var gemSocketsRaw = gemSheet.getRange(2, 2, 2000, 1).getValues();
  var gemSockets = gemSocketsRaw.reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
  }, []);

  var baseUrl = "https://classic.warcraftlogs.com:443/v1/";
  if (lang != "EN") {
    baseUrl = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com:443/v1/";
    baseUrlFrontEnd = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com/reports/";
  }

  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue().replace(/\s/g, "");
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var information = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 2), 1);
  shiftRangeByColumns(sheet, information, 1).clearContent();

  var logId = "";
  reportPathOrId = reportPathOrId.replace(".cn/", ".com/");
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

  var allPlayersUrl = baseUrl + "report/tables/casts/" + logId + apiKeyString + "&start=0&end=999999999999";
  var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString;
  var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));

  var baseSheetName = getStringForLang("gearIssuesTab", langKeys, langTrans, "", "", "", "");
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

  var excludeMotherShahraz = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("excludeMotherShahraz", langKeys, langTrans, "", "", "", "")).useRegularExpression(true).findNext(), -1).getValue();
  var listPlayersWithNoIssues = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("listNoIssues", langKeys, langTrans, "", "", "", "")).useRegularExpression(true).findNext(), -1).getValue();

  var gemsToConsider = 3;
  var onlyGems = false;
  var gemsToConsiderSelection = sheet.createTextFinder(getStringForLang("minimumGemQuality", langKeys, langTrans, "", "", "", "")).useRegularExpression(true).findNext();
  if (gemsToConsiderSelection != null) {
    var gemsToConsiderSelectionValue = shiftRangeByRows(sheet, gemsToConsiderSelection, 1);
    if (gemsToConsiderSelectionValue != null) {
      var value = gemsToConsiderSelectionValue.getValue();
      if (value != null) {
        if (value.toString().indexOf(getStringForLang("ignoreGems", langKeys, langTrans, "", "", "", "")) > -1) {
          gemsToConsider = 0;
        } else if (value.toString().indexOf(getStringForLang("uncommon", langKeys, langTrans, "", "", "", "")) > -1) {
          gemsToConsider = 2;
        } else if (value.toString().indexOf(getStringForLang("common", langKeys, langTrans, "", "", "", "")) > -1 && value.toString().indexOf("uncommon") < 0) {
          gemsToConsider = 1;
        } else if (value.toString().indexOf(getStringForLang("rare", langKeys, langTrans, "", "", "", "")) > -1) {
          gemsToConsider = 3;
        } else if (value.toString().indexOf(getStringForLang("epic", langKeys, langTrans, "", "", "", "")) > -1) {
          gemsToConsider = 4;
        }
        if (value.toString().indexOf(getStringForLang("onlyGems", langKeys, langTrans, "", "", "", "")) > -1) {
          onlyGems = true;
        }
      }
    }
  }

  var allPlayersData;
  var allPlayersString = "";
  var startSection = 0;
  var endLastFight = 0;
  allFightsData.fights.forEach(function (fight, fightCount) {
    if (((fight.originalBoss != null && Number(fight.originalBoss) > 0) && (fight.end_time - fight.start_time < 10000) && (fight.end_time - fight.start_time > 0)) && endLastFight > startSection) {
      allPlayersData = JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace("&start=0", "&start=" + startSection.toString()).replace("&end=999999999999", "&end=" + endLastFight.toString())));
      if (allPlayersData.entries != null && allPlayersData.entries.length > 0) {
        var additionalAllPlayersString = JSON.stringify(allPlayersData.entries);
        if (allPlayersString.length > 0) {
          allPlayersString = allPlayersString.substr(0, allPlayersString.length - 1);
          allPlayersString += "," + additionalAllPlayersString.substr(1, additionalAllPlayersString.length - 1);
        } else {
          allPlayersString = JSON.stringify(allPlayersData.entries);
        }
      }
      startSection = fight.end_time + 1;
      if (fightCount % 10 == 0)
        Utilities.sleep(300);
    }
    if ((fight.boss != null && fight.boss > 0) || (fight.originalBoss != null && Number(fight.originalBoss) > 0))
      endLastFight = fight.end_time;
  })
  allPlayersData = JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace("&start=0", "&start=" + startSection.toString()).replace("&end=999999999999", "&end=" + endLastFight.toString())));
  if (allPlayersData.entries != null && allPlayersData.entries.length > 0) {
    var additionalAllPlayersString = JSON.stringify(allPlayersData.entries);
    if (allPlayersString.length > 0) {
      allPlayersString = allPlayersString.substr(0, allPlayersString.length - 1);
      allPlayersString += "," + additionalAllPlayersString.substr(1, additionalAllPlayersString.length - 1);
    } else {
      allPlayersString = JSON.stringify(allPlayersData.entries);
    }
  }
  if (allPlayersString.length > 0)
    allPlayersData = JSON.parse(allPlayersString);

  Utilities.sleep(100);

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

  var allPlayersDataUnfiltered = JSON.parse(UrlFetchApp.fetch(allPlayersUrl));
  const allPlayersByNameAsc = sortByProperty(sortByProperty(allPlayersDataUnfiltered.entries, 'name'), "type");

  var playersFound = 0;
  var bossSummaryDataAll = [];
  allFightsData.fights.forEach(function (fight, fightCount) {
    if (fight.boss != null && fight.boss > 0) {
      bossSummaryDataAll[fightCount] = [];
      var bossUrl = allPlayersUrl.replace("&start=0&end=999999999999", "&start=" + fight.start_time + "&end=" + fight.end_time);
      bossSummaryDataAll[fightCount].push(JSON.parse(UrlFetchApp.fetch(bossUrl)));
      bossSummaryDataAll[fightCount].push(fight.boss.toString());
      bossSummaryDataAll[fightCount].push(fight.name.toString());
      if (fightCount % 10 == 0)
        Utilities.sleep(300);
    }
  })
  var gemIdInformedAbout = [];
  allPlayersByNameAsc.forEach(function (playerByNameAsc, playerCountByNameAsc) {
    if ((playerByNameAsc.type == "Druid" || playerByNameAsc.type == "Hunter" || playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Priest" || playerByNameAsc.type == "Paladin" || playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Shaman" || playerByNameAsc.type == "Warlock" || playerByNameAsc.type == "Warrior") && playerByNameAsc.total > 20) {
      var itemsFound = 0;
      var alreadyFilled = [];
      var playerName = "";
      var playerType = "";

      if (gemsToConsider > 0) {
        bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
          bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
            if (playerBoss.name == playerByNameAsc.name) {
              var metaGemfound = 0;
              var metaGemStringToPrint = "";
              var blueGemsFound = 0;
              var redGemsFound = 0;
              var yellowGemsFound = 0;
              if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                playerBoss.gear.forEach(function (item, itemCount) {
                  if (item.id != null && item.id.toString().length > 0) {
                    if (item.gems != null) {
                      item.gems.forEach(function (gem, gemCount) {
                        if (gem.itemLevel != null) {
                          var gemIdentified = false;
                          if (gem.id.toString() == "25897" || gem.id.toString() == "25899" || gem.id.toString() == "34220" || gem.id.toString() == "25890" || gem.id.toString() == "35503" || gem.id.toString() == "25895" || gem.id.toString() == "35501" || gem.id.toString() == "32641" || gem.id.toString() == "25901" || gem.id.toString() == "25893" || gem.id.toString() == "25896" || gem.id.toString() == "28557" || gem.id.toString() == "32409" || gem.id.toString() == "25894" || gem.id.toString() == "28556" || gem.id.toString() == "25898" || gem.id.toString() == "32640" || gem.id.toString() == "32410") {
                            metaGemfound = gem.id;
                            metaGemStringToPrint = item.name + " [" + getStringForLang("metaGemInactive", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]";
                            gemIdentified = true;
                          }
                          if (gem.id.toString() == "28466" || gem.id.toString() == "23113" || gem.id.toString() == "24047" || gem.id.toString() == "32204" || gem.id.toString() == "33139" || gem.id.toString() == "24053" || gem.id.toString() == "27679" || gem.id.toString() == "32209" || gem.id.toString() == "33138" || gem.id.toString() == "35315" || gem.id.toString() == "35761" || gem.id.toString() == "28467" || gem.id.toString() == "23114" || gem.id.toString() == "24048" || gem.id.toString() == "32205" || gem.id.toString() == "33143" || gem.id.toString() == "28469" || gem.id.toString() == "23114" || gem.id.toString() == "24050" || gem.id.toString() == "28120" || gem.id.toString() == "32207" || gem.id.toString() == "33140" || gem.id.toString() == "28470" || gem.id.toString() == "23115" || gem.id.toString() == "24052" || gem.id.toString() == "32208" || gem.id.toString() == "33144" || gem.id.toString() == "31860" || gem.id.toString() == "31861" || gem.id.toString() == "32210" || gem.id.toString() == "28468" || gem.id.toString() == "23116" || gem.id.toString() == "24051" || gem.id.toString() == "32206" || gem.id.toString() == "33142" || gem.id.toString() == "31869" || gem.id.toString() == "32637" || gem.id.toString() == "31868" || gem.id.toString() == "30582" || gem.id.toString() == "32222" || gem.id.toString() == "23098" || gem.id.toString() == "24058" || gem.id.toString() == "30584" || gem.id.toString() == "32217" || gem.id.toString() == "23101" || gem.id.toString() == "24059" || gem.id.toString() == "32218" || gem.id.toString() == "30593" || gem.id.toString() == "23099" || gem.id.toString() == "32638" || gem.id.toString() == "24060" || gem.id.toString() == "35760" || gem.id.toString() == "30547" || gem.id.toString() == "23100" || gem.id.toString() == "24061" || gem.id.toString() == "30556" || gem.id.toString() == "32220" || gem.id.toString() == "31866" || gem.id.toString() == "31867" || gem.id.toString() == "23099" || gem.id.toString() == "32221" || gem.id.toString() == "24060" || gem.id.toString() == "30547" || gem.id.toString() == "32219" || gem.id.toString() == "30587" || gem.id.toString() == "30604" || gem.id.toString() == "30585" || gem.id.toString() == "30554" || gem.id.toString() == "30607" || gem.id.toString() == "30551" || gem.id.toString() == "30564" || gem.id.toString() == "30573" || gem.id.toString() == "30581" || gem.id.toString() == "30593" || gem.id.toString() == "30601" || gem.id.toString() == "30565" || gem.id.toString() == "30575" || gem.id.toString() == "30553" || gem.id.toString() == "30591" || gem.id.toString() == "30558" || gem.id.toString() == "30559" || gem.id.toString() == "35318" || gem.id.toString() == "35759" || gem.id.toString() == "23104" || gem.id.toString() == "27809" || gem.id.toString() == "32639" || gem.id.toString() == "24067" || gem.id.toString() == "30602" || gem.id.toString() == "32226" || gem.id.toString() == "23105" || gem.id.toString() == "24062" || gem.id.toString() == "30590" || gem.id.toString() == "32223" || gem.id.toString() == "23106" || gem.id.toString() == "24065" || gem.id.toString() == "32225" || gem.id.toString() == "23103" || gem.id.toString() == "24066" || gem.id.toString() == "30608" || gem.id.toString() == "32224" || gem.id.toString() == "33782" || gem.id.toString() == "30592" || gem.id.toString() == "35758" || gem.id.toString() == "27785" || gem.id.toString() == "30548" || gem.id.toString() == "32635" || gem.id.toString() == "32639" || gem.id.toString() == "30583" || gem.id.toString() == "30586" || gem.id.toString() == "30605" || gem.id.toString() == "30606" || gem.id.toString() == "30560" || gem.id.toString() == "30594" || gem.id.toString() == "30550" || gem.id.toString() == "22460" || gem.id.toString() == "22459" || gem.id.toString() == "28119" || gem.id.toString() == "28120" || gem.id.toString() == "28123" || gem.id.toString() == "28363" || gem.id.toString() == "30588" || gem.id.toString() == "28290" || gem.id.toString() == "35316" || gem.id.toString() == "30589" || gem.id.toString() == "38550" || gem.id.toString() == "33633" || gem.id.toString() == "32735" || gem.id.toString() == "27786" || gem.id.toString() == "27820" || gem.id.toString() == "38546" || gem.id.toString() == "38548" || gem.id.toString() == "38547") {
                            yellowGemsFound++;
                            gemIdentified = true;
                          }
                          if (gem.id.toString() == "28458" || gem.id.toString() == "23095" || gem.id.toString() == "24027" || gem.id.toString() == "32193" || gem.id.toString() == "30598" || gem.id.toString() == "30571" || gem.id.toString() == "28459" || gem.id.toString() == "28595" || gem.id.toString() == "24028" || gem.id.toString() == "32194" || gem.id.toString() == "33131" || gem.id.toString() == "24036" || gem.id.toString() == "32199" || gem.id.toString() == "24032" || gem.id.toString() == "32198" || gem.id.toString() == "28460" || gem.id.toString() == "23094" || gem.id.toString() == "24029" || gem.id.toString() == "35489" || gem.id.toString() == "33134" || gem.id.toString() == "28461" || gem.id.toString() == "23096" || gem.id.toString() == "27812" || gem.id.toString() == "24030" || gem.id.toString() == "38549" || gem.id.toString() == "32196" || gem.id.toString() == "33133" || gem.id.toString() == "28462" || gem.id.toString() == "28595" || gem.id.toString() == "24031" || gem.id.toString() == "32197" || gem.id.toString() == "28361" || gem.id.toString() == "22460" || gem.id.toString() == "22459" || gem.id.toString() == "23108" || gem.id.toString() == "24056" || gem.id.toString() == "30555" || gem.id.toString() == "32215" || gem.id.toString() == "23109" || gem.id.toString() == "24057" || gem.id.toString() == "30603" || gem.id.toString() == "32216" || gem.id.toString() == "32833" || gem.id.toString() == "32836" || gem.id.toString() == "37503" || gem.id.toString() == "23110" || gem.id.toString() == "24055" || gem.id.toString() == "32212" || gem.id.toString() == "30549" || gem.id.toString() == "31862" || gem.id.toString() == "31863" || gem.id.toString() == "32213" || gem.id.toString() == "32634" || gem.id.toString() == "30574" || gem.id.toString() == "31118" || gem.id.toString() == "31864" || gem.id.toString() == "31865" || gem.id.toString() == "31116" || gem.id.toString() == "32214" || gem.id.toString() == "35707" || gem.id.toString() == "30563" || gem.id.toString() == "23111" || gem.id.toString() == "24054" || gem.id.toString() == "32211" || gem.id.toString() == "30546" || gem.id.toString() == "30566" || gem.id.toString() == "30600" || gem.id.toString() == "30552" || gem.id.toString() == "31117" || gem.id.toString() == "30572" || gem.id.toString() == "31869" || gem.id.toString() == "32637" || gem.id.toString() == "31868" || gem.id.toString() == "30582" || gem.id.toString() == "32222" || gem.id.toString() == "23098" || gem.id.toString() == "24058" || gem.id.toString() == "30584" || gem.id.toString() == "32217" || gem.id.toString() == "23101" || gem.id.toString() == "24059" || gem.id.toString() == "32218" || gem.id.toString() == "30593" || gem.id.toString() == "23099" || gem.id.toString() == "32638" || gem.id.toString() == "24060" || gem.id.toString() == "35760" || gem.id.toString() == "30547" || gem.id.toString() == "23100" || gem.id.toString() == "24061" || gem.id.toString() == "30556" || gem.id.toString() == "32220" || gem.id.toString() == "31866" || gem.id.toString() == "31867" || gem.id.toString() == "23099" || gem.id.toString() == "32221" || gem.id.toString() == "24060" || gem.id.toString() == "30547" || gem.id.toString() == "32219" || gem.id.toString() == "30587" || gem.id.toString() == "30604" || gem.id.toString() == "30585" || gem.id.toString() == "30554" || gem.id.toString() == "30607" || gem.id.toString() == "30551" || gem.id.toString() == "30564" || gem.id.toString() == "30573" || gem.id.toString() == "30581" || gem.id.toString() == "30593" || gem.id.toString() == "30601" || gem.id.toString() == "30565" || gem.id.toString() == "30575" || gem.id.toString() == "30553" || gem.id.toString() == "30591" || gem.id.toString() == "30558" || gem.id.toString() == "30559" || gem.id.toString() == "22460" || gem.id.toString() == "22459" || gem.id.toString() == "28118" || gem.id.toString() == "28123" || gem.id.toString() == "28362" || gem.id.toString() == "28363" || gem.id.toString() == "23097" || gem.id.toString() == "30588" || gem.id.toString() == "35316" || gem.id.toString() == "27777" || gem.id.toString() == "27812" || gem.id.toString() == "28361" || gem.id.toString() == "28360" || gem.id.toString() == "38545" || gem.id.toString() == "33633" || gem.id.toString() == "32636" || gem.id.toString() == "32195" || gem.id.toString() == "35487" || gem.id.toString() == "35488" || gem.id.toString() == "33132" || gem.id.toString() == "38548" || gem.id.toString() == "38547") {
                            redGemsFound++;
                            gemIdentified = true;
                          }
                          if (gem.id.toString() == "28463" || gem.id.toString() == "23118" || gem.id.toString() == "24033" || gem.id.toString() == "32200" || gem.id.toString() == "34831" || gem.id.toString() == "33135" || gem.id.toString() == "34256" || gem.id.toString() == "28464" || gem.id.toString() == "23119" || gem.id.toString() == "24035" || gem.id.toString() == "32201" || gem.id.toString() == "23120" || gem.id.toString() == "24039" || gem.id.toString() == "32203" || gem.id.toString() == "28465" || gem.id.toString() == "23121" || gem.id.toString() == "24037" || gem.id.toString() == "32202" || gem.id.toString() == "23108" || gem.id.toString() == "24056" || gem.id.toString() == "30555" || gem.id.toString() == "32215" || gem.id.toString() == "23109" || gem.id.toString() == "24057" || gem.id.toString() == "30603" || gem.id.toString() == "32216" || gem.id.toString() == "32833" || gem.id.toString() == "32836" || gem.id.toString() == "37503" || gem.id.toString() == "23110" || gem.id.toString() == "24055" || gem.id.toString() == "32212" || gem.id.toString() == "30549" || gem.id.toString() == "31862" || gem.id.toString() == "31863" || gem.id.toString() == "32213" || gem.id.toString() == "32634" || gem.id.toString() == "30574" || gem.id.toString() == "31118" || gem.id.toString() == "31864" || gem.id.toString() == "31865" || gem.id.toString() == "31116" || gem.id.toString() == "32214" || gem.id.toString() == "35707" || gem.id.toString() == "30563" || gem.id.toString() == "23111" || gem.id.toString() == "24054" || gem.id.toString() == "32211" || gem.id.toString() == "30546" || gem.id.toString() == "30566" || gem.id.toString() == "30600" || gem.id.toString() == "30552" || gem.id.toString() == "31117" || gem.id.toString() == "30572" || gem.id.toString() == "35318" || gem.id.toString() == "35759" || gem.id.toString() == "23104" || gem.id.toString() == "27809" || gem.id.toString() == "32639" || gem.id.toString() == "24067" || gem.id.toString() == "30602" || gem.id.toString() == "32226" || gem.id.toString() == "23105" || gem.id.toString() == "24062" || gem.id.toString() == "30590" || gem.id.toString() == "32223" || gem.id.toString() == "23106" || gem.id.toString() == "24065" || gem.id.toString() == "32225" || gem.id.toString() == "23103" || gem.id.toString() == "24066" || gem.id.toString() == "30608" || gem.id.toString() == "32224" || gem.id.toString() == "33782" || gem.id.toString() == "30592" || gem.id.toString() == "35758" || gem.id.toString() == "27785" || gem.id.toString() == "30548" || gem.id.toString() == "32635" || gem.id.toString() == "32639" || gem.id.toString() == "30583" || gem.id.toString() == "30586" || gem.id.toString() == "30605" || gem.id.toString() == "30606" || gem.id.toString() == "30560" || gem.id.toString() == "30594" || gem.id.toString() == "30550" || gem.id.toString() == "22460" || gem.id.toString() == "22459" || gem.id.toString() == "30589" || gem.id.toString() == "33633" || gem.id.toString() == "32735" || gem.id.toString() == "27820" || gem.id.toString() == "27786" || gem.id.toString() == "32636") {
                            blueGemsFound++;
                            gemIdentified = true;
                          }
                          if (!gemIdentified && gem.id != null && gem.id.toString().length > 0 && !gemIdInformedAbout.includes(gem.id.toString())) {
                            SpreadsheetApp.getUi().alert(getStringForLang("gemNotIdentified", langKeys, langTrans, gem.id.toString(), "", "", ""));
                            gemIdInformedAbout.push(gem.id.toString());
                          }
                        }
                      })
                    }
                  }
                })
              }
              if (!alreadyFilled.includes(metaGemStringToPrint) && metaGemfound > 0) {
                var metaGemActive = false;
                if (metaGemfound == 25896 && blueGemsFound > 2)
                  metaGemActive = true;
                if (metaGemfound == 25897 && redGemsFound > blueGemsFound)
                  metaGemActive = true;
                if ((metaGemfound == 32409 || metaGemfound == 25899 || metaGemfound == 25901 || metaGemfound == 25890 || metaGemfound == 32410) && redGemsFound > 1 && blueGemsFound > 1 && yellowGemsFound > 1)
                  metaGemActive = true;
                if (metaGemfound == 25898 && blueGemsFound > 4)
                  metaGemActive = true;
                if ((metaGemfound == 25893 || metaGemfound == 32640) && blueGemsFound > yellowGemsFound)
                  metaGemActive = true;
                if (metaGemfound == 34220 && blueGemsFound > 1)
                  metaGemActive = true;
                if (metaGemfound == 25895 && redGemsFound > yellowGemsFound)
                  metaGemActive = true;
                if ((metaGemfound == 25894 || metaGemfound == 28556 || metaGemfound == 28557) && redGemsFound > 0 && yellowGemsFound > 1)
                  metaGemActive = true;
                if (metaGemfound == 32641 && yellowGemsFound > 2)
                  metaGemActive = true;
                if (metaGemfound == 35503 && redGemsFound > 2)
                  metaGemActive = true;
                if (metaGemfound == 35501 && blueGemsFound > 1 && yellowGemsFound > 0)
                  metaGemActive = true;
                if (!metaGemActive) {
                  sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#f9cb9c").setValue(metaGemStringToPrint);
                  alreadyFilled.push(metaGemStringToPrint);
                  itemsFound++;
                }
              }
            }
          })
        })
      }
      if (gemsToConsider > 0) {
        bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
          bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
            if (playerBoss.name == playerByNameAsc.name) {
              if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                playerBoss.gear.forEach(function (item, itemCount) {
                  if (item.id != null && item.id.toString().length > 0) {
                    if (item.temporaryEnchant != null && item.temporaryEnchant.toString().length > 0 && item.temporaryEnchant.toString() != "0") {
                      if ((playerByNameAsc.type == "Hunter" || playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Warrior") && (item.temporaryEnchant.toString() == "3002" || item.temporaryEnchant.toString() == "2935")) {
                        if (!alreadyFilled.includes(item.name + " [" + getStringForLang("spellHitGearOnNonCaster", langKeys, langTrans, "", "", "", "") + "]")) {
                          sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("spellHitGearOnNonCaster", langKeys, langTrans, "", "", "", "") + "]");
                          alreadyFilled.push(item.name + " [" + getStringForLang("spellHitGearOnNonCaster", langKeys, langTrans, "", "", "", "") + "]");
                          itemsFound++;
                        }
                      }
                      if ((playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Priest" || playerByNameAsc.type == "Warlock") && (item.temporaryEnchant.toString() == "3003" || item.temporaryEnchant.toString() == "2658")) {
                        if (!alreadyFilled.includes(item.name + " [" + getStringForLang("meleeHitGearOnCaster", langKeys, langTrans, "", "", "", "") + "]")) {
                          sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("meleeHitGearOnCaster", langKeys, langTrans, "", "", "", "") + "]");
                          alreadyFilled.push(item.name + " [" + getStringForLang("meleeHitGearOnCaster", langKeys, langTrans, "", "", "", "") + "]");
                          itemsFound++;
                        }
                      }
                    }
                    if (item.gems != null) {
                      item.gems.forEach(function (gem, gemCount) {
                        if (gem.itemLevel != null) {
                          if ((playerByNameAsc.type == "Hunter" || playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Warrior") && (gem.id.toString() == "31860" || gem.id.toString() == "31861" || gem.id.toString() == "39725" || gem.id.toString() == "30606" || gem.id.toString() == "30605" || gem.id.toString() == "30564" || gem.id.toString() == "32221" || gem.id.toString() == "31867" || gem.id.toString() == "31866")) {
                            if (!alreadyFilled.includes(item.name + " [" + getStringForLang("spellHitGearOnNonCaster", langKeys, langTrans, "", "", "", "") + "]")) {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("spellHitGearOnNonCaster", langKeys, langTrans, "", "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("spellHitGearOnNonCaster", langKeys, langTrans, "", "", "", "") + "]");
                              itemsFound++;
                            }
                          }
                          if ((playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Priest" || playerByNameAsc.type == "Warlock") && (gem.id.toString() == "28468" || gem.id.toString() == "23116" || gem.id.toString() == "24051" || gem.id.toString() == "30559" || gem.id.toString() == "30553" || gem.id.toString() == "30575" || gem.id.toString() == "30556" || gem.id.toString() == "32220" || gem.id.toString() == "24061" || gem.id.toString() == "23100" || gem.id.toString() == "33142" || gem.id.toString() == "32206")) {
                            if (!alreadyFilled.includes(item.name + " [" + getStringForLang("meleeHitGearOnCaster", langKeys, langTrans, "", "", "", "") + "]")) {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("meleeHitGearOnCaster", langKeys, langTrans, "", "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("meleeHitGearOnCaster", langKeys, langTrans, "", "", "", "") + "]");
                              itemsFound++;
                            }
                          }
                          if (gem.id.toString() == "23112" || gem.id.toString() == "23436" || gem.id.toString() == "23077" || gem.id.toString() == "23441" || gem.id.toString() == "23440" || gem.id.toString() == "23117" || gem.id.toString() == "23438" || gem.id.toString() == "23437" || gem.id.toString() == "23107" || gem.id.toString() == "23079" || gem.id.toString() == "21929" || gem.id.toString() == "23439" || gem.id.toString() == "32227" || gem.id.toString() == "32229" || gem.id.toString() == "32228" || gem.id.toString() == "32231" || gem.id.toString() == "32249" || gem.id.toString() == "32230") {
                            if (!alreadyFilled.includes(item.name + " [" + getStringForLang("uncutGem", langKeys, langTrans, "", "", "", "") + "]")) {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("uncutGem", langKeys, langTrans, "", "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("uncutGem", langKeys, langTrans, "", "", "", "") + "]");
                              itemsFound++;
                            }
                          }
                        }
                      })
                    }
                  }
                })
              }
            }
          })
        })
      }

      if (!onlyGems) {
        bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
          bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
            if (playerBoss.name == playerByNameAsc.name) {
              if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                  if (itemBoss != null && itemBoss.slot != null) {
                    if (((itemBoss.id == null || itemBoss.id.toString().length == 0 || itemBoss.name == null || itemBoss.name.toString().length == 0)) && (Number(itemBoss.slot.toString()) >= 0 && Number(itemBoss.slot.toString()) <= 17 && Number(itemBoss.slot.toString()) != 3 && Number(itemBoss.slot.toString()) != 16)) {
                      if (!alreadyFilled.includes(itemBoss.name + " [" + getStringForLang("noItem", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id.toString() == "0")) {
                        var slotName = "error";
                        if (Number(itemBoss.slot.toString()) == 0)
                          slotName = getStringForLang("Head", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 1)
                          slotName = getStringForLang("Neck", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 2)
                          slotName = getStringForLang("Shoulders", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 4)
                          slotName = getStringForLang("Chest", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 5)
                          slotName = getStringForLang("Waist", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 6)
                          slotName = getStringForLang("Legs", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 7)
                          slotName = getStringForLang("Feet", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 8)
                          slotName = getStringForLang("Bracers", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 9)
                          slotName = getStringForLang("Hands", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 10)
                          slotName = getStringForLang("Ring1", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 11)
                          slotName = getStringForLang("Ring2", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 12)
                          slotName = getStringForLang("Trinket1", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 13)
                          slotName = getStringForLang("Trinket2", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 14)
                          slotName = getStringForLang("Cloak", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 15)
                          slotName = getStringForLang("Weapon", langKeys, langTrans, "", "", "", "");
                        else if (Number(itemBoss.slot.toString()) == 17)
                          slotName = getStringForLang("WandEtc", langKeys, langTrans, "", "", "", "");
                        sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#ff0707").setValue(slotName + " [" + getStringForLang("noItem", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                        alreadyFilled.push(itemBoss.name + " [" + getStringForLang("noItem", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                        itemsFound++;
                      }
                    }
                  }
                })
              }
            }
          })
        })
      }

      const excludedUndead = ["652", "653", "658", "662", "618", "726", "603", "604"];
      const excludedUndeadDemon = ["652", "653", "658", "662", "618", "726", "603", "604", "651", "657", "661", "619", "620", "621", "622", "725", "727", "729", "602", "607", "609"];

      allPlayersData.forEach(function (player, playerCount) {
        if (playerByNameAsc.name == player.name) {
          playerName = player.name;
          playerType = player.type;
          if (player.gear != null && player.gear.length > 0) {
            player.gear.forEach(function (item, itemCount) {
              if (item.id != null && item.id.toString().length > 0) {
                if (!onlyGems && (item.id.toString() == "13209" || item.id.toString() == "19812" || (item.temporaryEnchant != null && item.temporaryEnchant.toString().length > 0 && item.temporaryEnchant.toString() != "0" && (item.temporaryEnchant.toString() == "2684" || item.temporaryEnchant.toString() == "2685")))) {
                  var wrongTemporaryEnchant = false;
                  if (item.temporaryEnchant != null && item.temporaryEnchant.toString().length > 0 && item.temporaryEnchant.toString() != "0" && (item.temporaryEnchant.toString() == "2684" || item.temporaryEnchant.toString() == "2685"))
                    wrongTemporaryEnchant = true;
                  bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                    if (!excludedUndead.some(x => bossSummaryData[1].toString().endsWith(x))) {
                      bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                        if (playerBoss.name == playerByNameAsc.name) {
                          if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                            playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                              if (!alreadyFilled.includes(item.name + " [" + getStringForLang("vsNonUndeadParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && !alreadyFilled.includes(item.temporaryEnchantName + " [" + getStringForLang("vsNonUndeadParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && ((itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id.toString() == item.id.toString() && !wrongTemporaryEnchant) || wrongTemporaryEnchant)) {
                                if (!wrongTemporaryEnchant) {
                                  sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("vsNonUndeadParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                  alreadyFilled.push(item.name + " [" + getStringForLang("vsNonUndeadParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                  itemsFound++;
                                } else if (itemBoss.temporaryEnchant != null && itemBoss.temporaryEnchant.toString().length > 0 && itemBoss.temporaryEnchant.toString() == item.temporaryEnchant.toString()) {
                                  sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.temporaryEnchantName.replace(getStringForLang("blessedWizardOilRetail", langKeys, langTrans, "", "", "", ""), getStringForLang("blessedWizardOil", langKeys, langTrans, "", "", "", "")).replace(getStringForLang("consecratedSharpeningStoneRetail", langKeys, langTrans, "", "", "", ""), getStringForLang("consecratedSharpeningStone", langKeys, langTrans, "", "", "", "")) + " [" + getStringForLang("vsNonUndeadParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                  alreadyFilled.push(item.temporaryEnchantName + " [" + getStringForLang("vsNonUndeadParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                  itemsFound++;
                                }
                              }
                            })
                          }
                        }
                      })
                    }
                  })
                } else if (!onlyGems && (item.id.toString() == "23206" || item.id.toString() == "23207" || (item.temporaryEnchant != null && item.temporaryEnchant.toString().length > 0 && item.temporaryEnchant.toString() != "0" && item.temporaryEnchant.toString() == "3093"))) {
                  var wrongTemporaryEnchant = false;
                  if (item.temporaryEnchant != null && item.temporaryEnchant.toString().length > 0 && item.temporaryEnchant.toString() != "0" && item.temporaryEnchant.toString() == "3093")
                    wrongTemporaryEnchant = true;
                  bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                    if (!excludedUndeadDemon.some(x => bossSummaryData[1].toString().endsWith(x))) {
                      bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                        if (playerBoss.name == playerByNameAsc.name) {
                          if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                            playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                              if (!alreadyFilled.includes(item.name + " [" + getStringForLang("vsNonUndeadNonDemonParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && !alreadyFilled.includes(item.temporaryEnchantName + " [" + getStringForLang("vsNonUndeadNonDemonParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id.toString() == item.id.toString() && !wrongTemporaryEnchant) || wrongTemporaryEnchant) {
                                if (!wrongTemporaryEnchant) {
                                  sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("vsNonUndeadNonDemonParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                  alreadyFilled.push(item.name + " [" + getStringForLang("vsNonUndeadNonDemonParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                  itemsFound++;
                                } else if (itemBoss.temporaryEnchant != null && itemBoss.temporaryEnchant.toString().length > 0 && itemBoss.temporaryEnchant.toString() == item.temporaryEnchant.toString()) {
                                  sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.temporaryEnchantName.replace(getStringForLang("blessedWizardOilRetail", langKeys, langTrans, "", "", "", ""), getStringForLang("blessedWizardOil", langKeys, langTrans, "", "", "", "")).replace(getStringForLang("consecratedSharpeningStoneRetail", langKeys, langTrans, "", "", "", ""), getStringForLang("consecratedSharpeningStone", langKeys, langTrans, "", "", "", "")) + " [" + getStringForLang("vsNonUndeadNonDemonParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                  alreadyFilled.push(item.temporaryEnchantName + " [" + getStringForLang("vsNonUndeadNonDemonParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                  itemsFound++;
                                }
                              }
                            })
                          }
                        }
                      })
                    }
                  })
                } else if (!onlyGems && (item.id.toString() == "20537" || item.id.toString() == "20538" || item.id.toString() == "20539" || item.id.toString() == "20549" || item.id.toString() == "20550" || item.id.toString() == "20551" || item.id.toString() == "21530" || item.id.toString() == "21627" || item.id.toString() == "21687" || item.id.toString() == "21838" || item.id.toString() == "24097" || item.id.toString() == "31928" || item.id.toString() == "31939" || item.id.toString() == "32389" || item.id.toString() == "32390" || item.id.toString() == "32391" || item.id.toString() == "32392" || item.id.toString() == "32393" || item.id.toString() == "32394" || item.id.toString() == "32395" || item.id.toString() == "32396" || item.id.toString() == "32397" || item.id.toString() == "32398" || item.id.toString() == "32399" || item.id.toString() == "32400" || item.id.toString() == "32401" || item.id.toString() == "32402" || item.id.toString() == "32403" || item.id.toString() == "32404" || item.id.toString() == "32420" || item.id.toString() == "32649" || item.id.toString() == "32757")) {
                  bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                    if (!bossSummaryData[1].toString().endsWith("607") && (!bossSummaryData[1].toString().endsWith("609") || (bossSummaryData[1].toString().endsWith("609") && playerByNameAsc.type != "Mage" && playerByNameAsc.type != "Warlock")) && !bossSummaryData[1].toString().endsWith("620") && !bossSummaryData[1].toString().endsWith("621") && !bossSummaryData[1].toString().endsWith("727")) {
                      bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                        if (playerBoss.name == playerByNameAsc.name) {
                          var totalPlayerOnBossCount = 0;
                          var itemFlaggedPlayerOnBossCount = 0;
                          bossSummaryDataAll.forEach(function (bossSummaryDataAll, bossSummaryAllDataCount) {
                            if (bossSummaryDataAll[1].toString() == bossSummaryData[1].toString()) {
                              bossSummaryDataAll[0].entries.forEach(function (playerBossAll, playerBossAllCount) {
                                if (playerBossAll.gear != null && playerBossAll.gear.length > 0) {
                                  totalPlayerOnBossCount++;
                                  playerBossAll.gear.forEach(function (itemBossAll, itemBossAllCount) {
                                    if (itemBossAll.id.toString() == item.id.toString()) {
                                      itemFlaggedPlayerOnBossCount++;
                                    }
                                  })
                                }
                              })
                            }
                          })
                          if (((itemFlaggedPlayerOnBossCount * 100 / totalPlayerOnBossCount) < 50) && playerBoss.gear != null && playerBoss.gear.length > 0) {
                            playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                              if (!alreadyFilled.includes(item.name + " [" + getStringForLang("wrongSRgear", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id == item.id)) {
                                sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("wrongSRgear", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                alreadyFilled.push(item.name + " [" + getStringForLang("wrongSRgear", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                itemsFound++;
                              }
                            })
                          }
                        }
                      })
                    }
                  })
                } else if (!onlyGems && (item.id.toString() == "18849" || item.id.toString() == "18858" || item.id.toString() == "18862" || item.id.toString() == "18864" || item.id.toString() == "18851" || item.id.toString() == "18845" || item.id.toString() == "18846" || item.id.toString() == "18850" || item.id.toString() == "18834" || item.id.toString() == "18853" || item.id.toString() == "18856" || item.id.toString() == "18854" || item.id.toString() == "18863" || item.id.toString() == "18859" || item.id.toString() == "18857" || item.id.toString() == "18852" || item.id.toString() == "28234" || item.id.toString() == "28235" || item.id.toString() == "28236" || item.id.toString() == "28237" || item.id.toString() == "28238" || item.id.toString() == "28239" || item.id.toString() == "28240" || item.id.toString() == "28241" || item.id.toString() == "28242" || item.id.toString() == "28243" || item.id.toString() == "29592" || item.id.toString() == "29593" || item.id.toString() == "30343" || item.id.toString() == "30344" || item.id.toString() == "30345" || item.id.toString() == "30346" || item.id.toString() == "30348" || item.id.toString() == "30349" || item.id.toString() == "30350" || item.id.toString() == "30351" || item.id.toString() == "33046" || item.id.toString() == "37864" || item.id.toString() == "37865")) {
                  bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                    if (!bossSummaryData[1].toString().endsWith("618") && !bossSummaryData[1].toString().endsWith("619") && !bossSummaryData[1].toString().endsWith("622") && !bossSummaryData[1].toString().endsWith("727")) {
                      bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                        if (playerBoss.name == playerByNameAsc.name) {
                          if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                            playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                              if (!alreadyFilled.includes(item.name + " [" + getStringForLang("wrongPvPgear", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id == item.id)) {
                                sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("wrongPvPgear", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                alreadyFilled.push(item.name + " [" + getStringForLang("wrongPvPgear", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                itemsFound++;
                              }
                            })
                          }
                        }
                      })
                    }
                  })
                } else if (!onlyGems && (item.id.toString() == "25653" || item.id.toString() == "32863" || item.id.toString() == "11122" || item.id.toString() == "37313" || item.id.toString() == "37312" || item.id.toString() == "37311" || item.id.toString() == "32481")) {
                  bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                    bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                      if (playerBoss.name == playerByNameAsc.name) {
                        if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                          playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                            if (!alreadyFilled.includes(item.name + " [" + getStringForLang("uselessRidingGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id == item.id)) {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("uselessRidingGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("uselessRidingGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                              itemsFound++;
                            }
                          })
                        }
                      }
                    })
                  })
                } else if (!onlyGems && (item.id.toString() == "32538" || item.id.toString() == "32539" || item.id.toString() == "10518")) {
                  bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                    bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                      if (playerBoss.name == playerByNameAsc.name) {
                        if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                          playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                            if (!alreadyFilled.includes(item.name + " [" + getStringForLang("uselessSlowfallGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id == item.id)) {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("uselessSlowfallGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("uselessSlowfallGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                              itemsFound++;
                            }
                          })
                        }
                      }
                    })
                  })
                } else if (!onlyGems && (item.id.toString() == "30542" || item.id.toString() == "18984" || item.id.toString() == "30544" || item.id.toString() == "18986" || item.id.toString() == "23824" || item.id.toString() == "35581" || item.id.toString() == "23762" || item.id.toString() == "2789" || item.id.toString() == "10724" || item.id.toString() == "10518" || item.id.toString() == "4397")) {
                  bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                    if (!bossSummaryData[1].toString().endsWith("607") && !(bossSummaryData[1].toString().endsWith("608") && playerByNameAsc.type == "Mage" && item.id.toString() == "35581") && !(bossSummaryData[1].toString().endsWith("729") && (item.id.toString() == "35581" || item.id.toString() == "23824")) && !(bossSummaryData[1].toString().endsWith("727") && item.id.toString() == "4397")) {
                      bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                        if (playerBoss.name == playerByNameAsc.name) {
                          var totalPlayerOnBossCount = 0;
                          var itemFlaggedPlayerOnBossCount = 0;
                          bossSummaryDataAll.forEach(function (bossSummaryDataAll, bossSummaryAllDataCount) {
                            if (bossSummaryDataAll[1].toString() == bossSummaryData[1].toString()) {
                              bossSummaryDataAll[0].entries.forEach(function (playerBossAll, playerBossAllCount) {
                                if (playerBossAll.gear != null && playerBossAll.gear.length > 0) {
                                  totalPlayerOnBossCount++;
                                  playerBossAll.gear.forEach(function (itemBossAll, itemBossAllCount) {
                                    if (itemBossAll.id.toString() == item.id.toString() || (item.id.toString() == "23824" && itemBossAll.id.toString() == "35581") || (item.id.toString() == "35581" && itemBossAll.id.toString() == "23824")) {
                                      itemFlaggedPlayerOnBossCount++;
                                    }
                                  })
                                }
                              })
                            }
                          })
                          if (((itemFlaggedPlayerOnBossCount * 100 / totalPlayerOnBossCount) < 50) && playerBoss.gear != null && playerBoss.gear.length > 0) {
                            playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                              if (!alreadyFilled.includes(item.name + " [" + getStringForLang("uselessEngiGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id == item.id)) {
                                sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("uselessEngiGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                alreadyFilled.push(item.name + " [" + getStringForLang("uselessEngiGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                                itemsFound++;
                              }
                            })
                          }
                        }
                      })
                    }
                  })
                } else if (!onlyGems && (suboptimalItems.indexOf(item.id) > -1)) {
                  bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                    bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                      if (playerBoss.name == playerByNameAsc.name) {
                        if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                          playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                            if (!alreadyFilled.includes(item.name + " [" + getStringForLang("uselessGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]") && (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id == item.id)) {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b9a3ee").setValue(item.name + " [" + getStringForLang("uselessGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("uselessGearParam", langKeys, langTrans, bossSummaryData[2], "", "", "") + "]");
                              itemsFound++;
                            }
                          })
                        }
                      }
                    })
                  })
                }
                if (!isItemToBeIgnored(item, gearToIgnore)) {
                  if ((item.id != null && item.id.toString().length > 0 && item.id.toString() != "0") && item.id.toString() != "21471" && item.id.toString() != "21597" && item.id.toString() != "27471" && item.id.toString() != "1172" && item.id.toString() != "22937" && item.id.toString() != "6803" && item.id.toString() != "23029" && item.id.toString() != "15857" && item.id.toString() != "11928" && item.id.toString() != "23048" && item.id.toString() != "23049" && item.id.toString() != "7344" && item.id.toString() != "22329" && item.id.toString() != "11904" && item.id.toString() != "21666" && item.id.toString() != "8625" && item.id.toString() != "3451" && item.id.toString() != "8624" && item.id.toString() != "6182" && item.id.toString() != "6654" && item.id.toString() != "15986" && item.id.toString() != "12471" && item.id.toString() != "19115" && item.id.toString() != "6774" && item.id.toString() != "2565" && item.id.toString() != "16887" && item.id.toString() != "7297" && item.id.toString() != "15942" && item.id.toString() != "22206" && item.id.toString() != "11522" && item.id.toString() != "3419" && item.id.toString() != "2879" && item.id.toString() != "22994" && item.id.toString() != "4696" && item.id.toString() != "15940" && item.id.toString() != "2410" && item.id.toString() != "8626" && item.id.toString() != "15945" && item.id.toString() != "15973" && item.id.toString() != "15967" && item.id.toString() != "18425" && item.id.toString() != "15993" && item.id.toString() != "4838" && item.id.toString() != "4836" && item.id.toString() != "15984" && item.id.toString() != "13261" && item.id.toString() != "15925" && item.id.toString() != "15935" && item.id.toString() != "15941" && item.id.toString() != "5611" && item.id.toString() != "15972" && item.id.toString() != "5028" && item.id.toString() != "6653" && item.id.toString() != "3424" && item.id.toString() != "15989" && item.id.toString() != "15932" && item.id.toString() != "15974" && item.id.toString() != "15975" && item.id.toString() != "15965" && item.id.toString() != "7609" && item.id.toString() != "15982" && item.id.toString() != "15931" && item.id.toString() != "15970" && item.id.toString() != "15963" && item.id.toString() != "7559" && item.id.toString() != "7611" && item.id.toString() != "9914" && item.id.toString() != "15930" && item.id.toString() != "15985" && item.id.toString() != "15912" && item.id.toString() != "15978" && item.id.toString() != "15971" && item.id.toString() != "4837" && item.id.toString() != "3420" && item.id.toString() != "7555" && item.id.toString() != "15947" && item.id.toString() != "15928" && item.id.toString() != "15962" && item.id.toString() != "15979" && item.id.toString() != "7557" && item.id.toString() != "15976" && item.id.toString() != "15983" && item.id.toString() != "15206" && item.id.toString() != "15934" && item.id.toString() != "3675" && item.id.toString() != "7554" && item.id.toString() != "7608" && item.id.toString() != "7558" && item.id.toString() != "9769" && item.id.toString() != "15939" && item.id.toString() != "29273" && item.id.toString() != "29270" && item.id.toString() != "28412" && item.id.toString() != "29274" && item.id.toString() != "29272" && item.id.toString() != "29170" && item.id.toString() != "29330" && item.id.toString() != "28734" && item.id.toString() != "31493" && item.id.toString() != "28781" && item.id.toString() != "28187" && item.id.toString() != "29271" && item.id.toString() != "28603" && item.id.toString() != "29269" && item.id.toString() != "28525" && item.id.toString() != "28728" && item.id.toString() != "28387" && item.id.toString() != "27714" && item.id.toString() != "27477" && item.id.toString() != "28260" && item.id.toString() != "31732" && item.id.toString() != "32533" && item.id.toString() != "27534" && item.id.toString() != "31699" && item.id.toString() != "31731" && item.id.toString() != "31494" && item.id.toString() != "28213" && item.id.toString() != "25099" && item.id.toString() != "32651" && item.id.toString() != "32452" && item.id.toString() != "34179" && item.id.toString() != "30872" && item.id.toString() != "25095" && item.id.toString() != "28941" && item.id.toString() != "25091" && item.id.toString() != "31823" && item.id.toString() != "25097" && item.id.toString() != "33334" && item.id.toString() != "28938" && item.id.toString() != "34206" && item.id.toString() != "25096" && item.id.toString() != "25092" && item.id.toString() != "25090" && item.id.toString() != "28346" && item.id.toString() != "32520" && item.id.toString() != "32361" && item.id.toString() != "33325" && item.id.toString() != "32350" && item.id.toString() != "25093" && item.id.toString() != "32961" && item.id.toString() != "29923" && item.id.toString() != "34033" && item.id.toString() != "35074" && item.id.toString() != "25098" && item.id.toString() != "33681" && item.id.toString() != "35016" && item.id.toString() != "30911" && item.id.toString() != "33736" && item.id.toString() != "31978" && item.id.toString() != "25094" && item.id.toString() != "35008") {
                    if (!onlyGems && (item.permanentEnchant == null || item.permanentEnchant.length < 1 || isEnchantBad(item.permanentEnchant.toString(), badIds, item.slot.toString()))) {
                      if (item.slot.toString() == "0" || item.slot.toString() == "2" || item.slot.toString() == "4" || item.slot.toString() == "6" || item.slot.toString() == "7" || item.slot.toString() == "8" || item.slot.toString() == "9" || item.slot.toString() == "14" || item.slot.toString() == "15" || (item.slot.toString() == "16" && item.icon.indexOf("_misc_") < 0)) {
                        if (!alreadyFilled.includes(item.name + " [" + getStringForLang("noEnchant", langKeys, langTrans, "", "", "", "") + "]") && !alreadyFilled.includes(item.name + " [" + getStringForLang("badEnchant", langKeys, langTrans, "", "", "", "") + "]")) {
                          var issueFoundOnMother = false;
                          var issueFoundElsewhere = false;
                          var spellPenFoundOnIC = false;
                          var spellPenFoundElsewhere = false;
                          bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                            bossSummaryData[0].entries.forEach(function (playerBoss, playerBossCount) {
                              if (playerBoss.name == playerByNameAsc.name) {
                                if (playerBoss.gear != null && playerBoss.gear.length > 0) {
                                  playerBoss.gear.forEach(function (itemBoss, itemBossCount) {
                                    if (itemBoss.id != null && itemBoss.id.toString().length > 0 && itemBoss.id == item.id) {
                                      if (bossSummaryData[1].toString().endsWith("607")) {
                                        issueFoundOnMother = true;
                                      } else {
                                        issueFoundElsewhere = true;
                                      }
                                      if (bossSummaryData[1].toString().endsWith("608") && item.permanentEnchant != null && item.permanentEnchant.toString() == "2938") {
                                        spellPenFoundOnIC = true;
                                      } else if (item.permanentEnchant != null && item.permanentEnchant.toString() == "2938") {
                                        spellPenFoundElsewhere = true;
                                      }
                                    }
                                  })
                                }
                              }
                            })
                          })
                          if (item.permanentEnchant == null) {
                            if (!(excludeMotherShahraz.indexOf("yes") > -1 && ((issueFoundOnMother && !issueFoundElsewhere))) || issueFoundElsewhere || excludeMotherShahraz.indexOf("yes") < 0) {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#f8aaaa").setValue(item.name + " [" + getStringForLang("noEnchant", langKeys, langTrans, "", "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("noEnchant", langKeys, langTrans, "", "", "", "") + "]");
                              itemsFound++;
                            }
                          }
                          else {
                            if (!(item.permanentEnchant.toString() == "2938" && (playerByNameAsc.type == "Priest" || (spellPenFoundOnIC && !spellPenFoundElsewhere))) && !(item.permanentEnchant.toString() == "2669" && (playerByNameAsc.type == "Paladin" || playerByNameAsc.type == "Shaman"))) {
                              if (((excludeMotherShahraz.indexOf("yes") > -1 && ((issueFoundOnMother && !issueFoundElsewhere) || issueFoundElsewhere)) || excludeMotherShahraz.indexOf("yes") < 0)) {
                                var cellValue = item.name;
                                cellValue += " [" + getEnchantBadName(item.permanentEnchant.toString(), badIds, badIdNames, item.slot.toString()) + "]";
                                sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#fdf2ce").setValue(cellValue);//.setNote(getStringForLang("suboptimalEnchant", langKeys, langTrans, "", "", "", ""));
                                alreadyFilled.push(item.name + " [" + getStringForLang("badEnchant", langKeys, langTrans, "", "", "", "") + "]");
                                itemsFound++;
                              }
                            }
                          }
                        }
                      }
                    }
                    if (gemItemIds.indexOf(item.id) > -1) {
                      var itemSockets = searchEntryForId(gemItemIds, gemSockets, item.id.toString());
                      if (gemsToConsider > 0 && (item.gems == null || item.gems.length == null || item.gems.length == 0 || item.gems.length < itemSockets)) {
                        if (!alreadyFilled.includes(item.name + " [" + getStringForLang("noGem", langKeys, langTrans, "", "", "", "") + "]")) {
                          var numberOfGems = itemSockets;
                          if (item.gems != null && item.gems.length != null && item.gems.length >= 0)
                            numberOfGems -= item.gems.length;
                          for (var m = numberOfGems; m > 0; m--) {
                            sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#f7cfe1").setValue(item.name + " [" + getStringForLang("noGem", langKeys, langTrans, "", "", "", "") + "]");
                            alreadyFilled.push(item.name + " [" + getStringForLang("noGem", langKeys, langTrans, "", "", "", "") + "]");
                            itemsFound++;
                          }
                        }
                      }
                    }
                    if (item.gems != null) {
                      item.gems.forEach(function (gem, gemCount) {
                        if (gem.itemLevel != null) {
                          if (!alreadyFilled.includes(item.name + " [" + getStringForLang("badGem", langKeys, langTrans, "", "", "", "") + "]")) {
                            if (gemsToConsider > 1 && gem.itemLevel < 60) {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#b7b7b7").setValue(item.name + " [" + getStringForLang("commonGem", langKeys, langTrans, "", "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("badGem", langKeys, langTrans, "", "", "", "") + "]");
                              itemsFound++;
                            } else if (gemsToConsider > 2 && gem.itemLevel == 60 && gem.id.toString() != "38549" && gem.id.toString() != "32836" && gem.id.toString() != "28118" && gem.id.toString() != "27679" && gem.id.toString() != "30571" && gem.id.toString() != "27812" && gem.id.toString() != "30598" && gem.id.toString() != "27777" && gem.id.toString() != "28362" && gem.id.toString() != "28361" && gem.id.toString() != "28363" && gem.id.toString() != "28123" && gem.id.toString() != "28119" && gem.id.toString() != "28120" && gem.id.toString() != "28360" && gem.id.toString() != "38545" && gem.id.toString() != "38550" && gem.id.toString() != "27785" && gem.id.toString() != "27809" && gem.id.toString() != "38546" && gem.id.toString() != "27820" && gem.id.toString() != "38548" && gem.id.toString() != "27786" && gem.id.toString() != "38547") {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#bfeeae").setValue(item.name + " [" + getStringForLang("uncommonGem", langKeys, langTrans, "", "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("badGem", langKeys, langTrans, "", "", "", "") + "]");
                              itemsFound++;
                            } else if (gemsToConsider > 3 && gem.itemLevel < 100 && gem.id.toString() != "38549" && gem.id.toString() != "32836" && gem.id.toString() != "28118" && gem.id.toString() != "27679" && gem.id.toString() != "30571" && gem.id.toString() != "27812" && gem.id.toString() != "30598" && gem.id.toString() != "27777" && gem.id.toString() != "28362" && gem.id.toString() != "28361" && gem.id.toString() != "28363" && gem.id.toString() != "28123" && gem.id.toString() != "28119" && gem.id.toString() != "28120" && gem.id.toString() != "28360" && gem.id.toString() != "38545" && gem.id.toString() != "38550" && gem.id.toString() != "27785" && gem.id.toString() != "27809" && gem.id.toString() != "38546" && gem.id.toString() != "27820" && gem.id.toString() != "38548" && gem.id.toString() != "27786" && gem.id.toString() != "38547" && gem.id.toString() != "32409" && gem.id.toString() != "34220" && gem.id.toString() != "28118" && gem.id.toString() != "25896" && gem.id.toString() != "25897" && gem.id.toString() != "28556" && gem.id.toString() != "25901" && gem.id.toString() != "28557" && gem.id.toString() != "25893" && gem.id.toString() != "25894" && gem.id.toString() != "34831" && gem.id.toString() != "25898" && gem.id.toString() != "35503" && gem.id.toString() != "32410" && gem.id.toString() != "35501" && gem.id.toString() != "25899" && gem.id.toString() != "25895" && gem.id.toString() != "32641" && gem.id.toString() != "25890" && gem.id.toString() != "32640" && gem.id.toString() != "33633" && gem.id.toString() != "30549" && gem.id.toString() != "30556" && gem.id.toString() != "30582" && gem.id.toString() != "30550" && gem.id.toString() != "30602" && gem.id.toString() != "30564" && gem.id.toString() != "30588" && gem.id.toString() != "30600" && gem.id.toString() != "30605" && gem.id.toString() != "30555" && gem.id.toString() != "30603" && gem.id.toString() != "30606" && gem.id.toString() != "30585" && gem.id.toString() != "30547" && gem.id.toString() != "30581" && gem.id.toString() != "30551" && gem.id.toString() != "22459" && gem.id.toString() != "31116" && gem.id.toString() != "30593" && gem.id.toString() != "30590" && gem.id.toString() != "30563" && gem.id.toString() != "30584" && gem.id.toString() != "30553" && gem.id.toString() != "30554" && gem.id.toString() != "30592" && gem.id.toString() != "30586" && gem.id.toString() != "31118" && gem.id.toString() != "30546" && gem.id.toString() != "30572" && gem.id.toString() != "31117" && gem.id.toString() != "30552" && gem.id.toString() != "30591" && gem.id.toString() != "30573" && gem.id.toString() != "30559" && gem.id.toString() != "30558" && gem.id.toString() != "30566" && gem.id.toString() != "30575" && gem.id.toString() != "30604" && gem.id.toString() != "30560" && gem.id.toString() != "30548" && gem.id.toString() != "30587" && gem.id.toString() != "30589" && gem.id.toString() != "34256" && gem.id.toString() != "32735" && gem.id.toString() != "30607" && gem.id.toString() != "30583" && gem.id.toString() != "30594" && gem.id.toString() != "30574" && gem.id.toString() != "30608" && gem.id.toString() != "30601" && gem.id.toString() != "30565" && gem.id.toString() != "35759" && gem.id.toString() != "35758") {
                              sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn + itemsFound + 1).setBackground("#a9c2f1").setValue(item.name + " [" + getStringForLang("rareGem", langKeys, langTrans, "", "", "", "") + "]");
                              alreadyFilled.push(item.name + " [" + getStringForLang("badGem", langKeys, langTrans, "", "", "", "") + "]");
                              itemsFound++;
                            }
                          }
                        }
                      })
                    }
                  }
                }
              }
            })
          }
        }
      })
      if ((listPlayersWithNoIssues.indexOf("yes") < 0 && itemsFound > 0) || listPlayersWithNoIssues.indexOf("yes") > -1) {
        var range = sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn);
        range.setValue(playerName);
        range.setBackground(getColourForPlayerClass(playerType));
        playersFound++;
      }
      if (listPlayersWithNoIssues.indexOf("yes") > -1 && itemsFound == 0) {
        sheet.getRange(playersFound + firstPlayerNameRow - 1, firstPlayerNameColumn + 1).setValue("---------------------------------------------------------------");
      }
    }
  })
}

function isEnchantBad(enchantId, badIds, slot) {
  var isEnchantBad = false;
  badIds.forEach(function (badId, badIdCount) {
    if (badId.toString().split(" [")[0].split("[")[0] == enchantId.toString()) {
      if (badId.toString().indexOf("[") > -1) {
        if (badId.toString().split("[")[1].split("]")[0] == slot) {
          isEnchantBad = true;
        }
      }
      else
        isEnchantBad = true;
    }
  })
  return isEnchantBad;
}

function getEnchantBadName(enchantId, badIds, badIdNames, itemSlot) {
  var enchantName = "";
  badIds.forEach(function (badId, badIdCount) {
    if (badId.toString().split(" [")[0].split("[")[0] == enchantId.toString()) {
      if (badId.toString().indexOf("[") > - 1) {
        if (badId.toString().split("[")[1].split("]")[0] == itemSlot)
          enchantName = badIdNames[badIdCount];
      } else
        enchantName = badIdNames[badIdCount];
    }
  })
  return enchantName;
}

function isItemToBeIgnored(itemId, gearToIgnoreIds) {
  var isItemToBeIgnored = false;
  gearToIgnoreIds.forEach(function (gearToIgnoreId, gearToIgnoreIdCount) {
    if (itemId.id.toString() == gearToIgnoreId.toString()) {
      isItemToBeIgnored = true;
    }
  })
  return isItemToBeIgnored;
}