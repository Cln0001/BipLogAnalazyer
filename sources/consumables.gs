function populateBuffConsumables() {
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

  if (currentVersion.indexOf(codeVersion) < 0) {
    SpreadsheetApp.getUi().alert(getStringForLang("sheetOutdated", langKeys, langTrans, "", "", "", ""));
  }

  var firstPlayerNameRow = 6;
  var firstPlayerNameColumn = 2;

  var sheets = ss.getSheets();
  for (var c = sheets.length - 1; c >= 0; c--) {
    var sheetNameSearch = sheets[c].getName();
    if (sheetNameSearch.indexOf("buffConsumables") > -1) {
      ss.deleteSheet(sheets[c]);
    }
  }
  var confSheet = confSpreadSheet.getSheetByName("buffConsumables");
  var conf = confSheet.copyTo(ss).hideSheet().setName("buffConsumables");

  instructionsSheet.getRange(26, 2).setValue("");
  instructionsSheet.getRange(27, 2).setValue("");

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, 45, 1).clearContent();
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn + 3, 45, 9).clearContent().clearNote().setFontStyle("normal").setFontLine("none");
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, 45, 12).setFontColor("black").setFontStyle("normal");
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn + 3, 45, 3).setFontWeight("normal");
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn + 5, 45, 1).setFontWeight("normal");

  if (darkMode)
    sheet.getRange(1, 1, 50, 13).setBackground("#d9d9d9");
  else
    sheet.getRange(1, 1, 50, 13).setBackground("white");
  sheet.getRange(4, 3, 1, 1).setBackground("#cccccc");
  sheet.getRange(2, 11, 1, 1).setBackground("#cccccc");

  var baseUrl = "https://classic.warcraftlogs.com:443/v1/";
  var baseUrlFrontEnd = "https://classic.warcraftlogs.com/reports/"
  if (lang != "EN") {
    baseUrl = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com:443/v1/";
    baseUrlFrontEnd = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com/reports/";
  }

  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue().replace(/\s/g, "");
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var information = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 2), 1);
  addColumnsToRange(sheet, shiftRangeByColumns(sheet, information, 1), -1).clearContent();

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

  var allPlayersUrl = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndString;
  var urlPeopleTracked = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndString;
  var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString + startEndString;
  var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));

  var baseSheetName = getStringForLang("buffConsumablesTab", langKeys, langTrans, "", "", "", "")
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

  var spellIds = [];
  var spellIdStyles = [];

  var headersRaw = conf.getRange(1, 1, 1, 100).getValues();
  const headers = headersRaw.filter(function (x) {
    return !(x.every(element => element === (undefined || null || '')))
  });
  var headerDoneCount = 0;
  var headersToPrint = [];
  headersToPrint[0] = [];
  headers[0].forEach(function (header, headerCount) {
    if (header != null && header.length != null && header.length > 0) {
      headersToPrint[0].push(header.replace(" (V1.2.3+)", ""));
      spellIds[headerDoneCount] = getEntriesForHeader(headers[0], header, conf);
      spellIdStyles[headerDoneCount] = getStylesForHeader(headers[0], header, conf);
      headerDoneCount++;
    }
  })

  var allPlayersData = JSON.parse(UrlFetchApp.fetch(urlPeopleTracked));

  const allPlayersByNameAsc = sortByProperty(sortByProperty(allPlayersData.entries, 'name'), "type");

  var bossSummaryDataAll = [];
  var bossDamageDataAll = [];
  var fightsParsed = 0;
  var previousFight = null;
  allFightsData.fights.forEach(function (fight, fightCount) {
    if (fight.boss != null && fight.boss > 0 && fight.start_time >= sectionToLookAtStart && fight.end_time <= sectionToLookAtEnd) {
      bossSummaryDataAll[fightCount] = [];
      if (previousFight != null && previousFight.originalBoss != null && previousFight.originalBoss > 0 && (previousFight.end_time - previousFight.start_time < 100)) {
        bossSummaryDataAll[fightCount].push(JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace(startEndString, "&start=" + previousFight.start_time + "&end=" + fight.end_time).replace("damage-taken", "summary"))));
      } else {
        bossSummaryDataAll[fightCount].push(JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace(startEndString, "&start=" + fight.start_time + "&end=" + fight.end_time).replace("damage-taken", "summary"))));
      }
      bossSummaryDataAll[fightCount].push(fight.id.toString());
      bossDamageDataAll[fightCount] = [];
      if (previousFight != null && previousFight.originalBoss != null && previousFight.originalBoss > 0 && previousFight.end_time - previousFight.start_time < 100) {
        bossDamageDataAll[fightCount].push(JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace(startEndString, "&start=" + previousFight.start_time + "&end=" + fight.end_time).replace("damage-taken", "damage-done") + "&abilityid=27187")));
      } else {
        bossDamageDataAll[fightCount].push(JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace(startEndString, "&start=" + fight.start_time + "&end=" + fight.end_time).replace("damage-taken", "damage-done") + "&abilityid=27187")));
      }
      bossDamageDataAll[fightCount].push(fight.id.toString());
      fightsParsed++;
      if (fightsParsed % 5 == 0)
        Utilities.sleep(200);
    }
    previousFight = fight;
  })
  Utilities.sleep(100);

  var playersFound = 0;
  allPlayersByNameAsc.forEach(function (playerByNameAsc, playerCountByNameAsc) {
    if ((playerByNameAsc.type == "Druid" || playerByNameAsc.type == "Hunter" || playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Priest" || playerByNameAsc.type == "Paladin" || playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Shaman" || playerByNameAsc.type == "Warlock" || playerByNameAsc.type == "Warrior") && playerByNameAsc.total > 20) {
      var playerRow = [];
      playerRow[0] = [];
      var bossBuffDataAll = [];
      var playerFightCount = 0;
      var allNecksBuffData = JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace("damage-taken", "buffs") + "&targetid=" + playerByNameAsc.id + "&encounter=-2"));
      var bossCovered = [];
      var jcNeckFound = [];
      var suboptimalStuffFound = "";
      var badBuffFoodFound = [];
      var wellFedName = "";

      var previousFightBuffs = null;
      allFightsData.fights.forEach(function (fight, fightCount) {
        if (fight.boss != null && fight.boss > 0 && fight.start_time >= sectionToLookAtStart && fight.end_time <= sectionToLookAtEnd) {
          bossBuffDataAll[fightCount] = [];
          var bossUrl = null;
          if (previousFightBuffs != null && previousFightBuffs.originalBoss != null && previousFightBuffs.originalBoss > 0 && (previousFightBuffs.end_time - previousFightBuffs.start_time < 100)) {
            bossUrl = allPlayersUrl.replace(startEndString, "&start=" + previousFightBuffs.start_time + "&end=" + fight.end_time).replace("damage-taken", "buffs") + "&sourceid=" + playerByNameAsc.id;
          } else {
            bossUrl = allPlayersUrl.replace(startEndString, "&start=" + fight.start_time + "&end=" + fight.end_time).replace("damage-taken", "buffs") + "&sourceid=" + playerByNameAsc.id;
          }
          var bossData = JSON.parse(UrlFetchApp.fetch(bossUrl))
          bossBuffDataAll[fightCount].push(bossData);
          bossBuffDataAll[fightCount].push(fight.id.toString());
          if (bossData != null && bossData.auras != null && bossData.auras.length > 0) {
            playerFightCount++;
            if (playerFightCount % 5 == 0)
              Utilities.sleep(500);
          }
        }
        previousFightBuffs = fight;
      })
      var range = sheet.getRange(playersFound + firstPlayerNameRow, firstPlayerNameColumn);
      range.setValue(playerByNameAsc.name).setBackground(getColourForPlayerClass(playerByNameAsc.type)).setFontWeight("bold");
      var names = "";
      spellIds.forEach(function (spellIdsColumn, spellIdsColumnCount) {
        var headerValue = sheet.getRange(firstPlayerNameRow - 1, firstPlayerNameColumn + spellIdsColumnCount + 3).getValue();
        if (headerValue != getStringForLang("weaponEnhancement", langKeys, langTrans, "", "", "", "") && headerValue != getStringForLang("nrBossFightsWithNeck", langKeys, langTrans, "", "", "", "")) {
          var bossIdsFound = [];
          var isPlayerACheapass = false;
          spellIdsColumn.forEach(function (spellIdsCell, spellIdsCellCount) {
            allFightsData.fights.forEach(function (fight, fightRawCount) {
              if (fight.id != null && Number(fight.id) > 0 && fight.start_time >= sectionToLookAtStart && fight.end_time <= sectionToLookAtEnd) {
                bossBuffDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
                  if (bossSummaryData[1] == fight.id) {
                    bossSummaryData[0].auras.forEach(function (buffData, buffDataCount) {
                      if (buffData.guid != null && buffData.guid.toString().length > 0 && buffData.guid.toString() == spellIdsCell.toString().split(" [")[0]) {
                        if (bossIdsFound.indexOf(fight.id) < 0 && (headerValue != getStringForLang("flaskText", langKeys, langTrans, "", "", "", "") || (headerValue == getStringForLang("flaskText", langKeys, langTrans, "", "", "", "") && bossCovered.indexOf(fight.id) < 0))) {
                          bossIdsFound.push(fight.id);
                          try {
                            if (((spellIdStyles[spellIdsColumnCount][spellIdsCellCount] == "italic") && !(buffData.guid.toString() == "17539" && playerByNameAsc.type == "Paladin") && !(buffData.guid.toString() == "24799" && (playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Warrior"))) || (buffData.guid.toString() == "33721" && playerByNameAsc.type != "Mage" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Shaman") || ((buffData.guid.toString() == "11406" || buffData.guid.toString() == "28497" || buffData.guid.toString() == "43764" || buffData.guid.toString() == "28520" || buffData.guid.toString() == "41606" || buffData.guid.toString() == "11374") && playerByNameAsc.type != "Rogue" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Warrior" && playerByNameAsc.type != "Shaman" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Paladin") || ((buffData.guid.toString() == "28491" || buffData.guid.toString() == "33268" || buffData.guid.toString() == "17627") && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Priest" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Shaman") || (buffData.guid.toString() == "28493" && playerByNameAsc.type != "Mage") || ((buffData.guid.toString() == "28501" || buffData.guid.toString() == "43722") && playerByNameAsc.type != "Mage" && playerByNameAsc.type != "Warlock") || (buffData.guid.toString() == "28503" && playerByNameAsc.type != "Warlock" && playerByNameAsc.type != "Priest") || ((buffData.guid.toString() == "33726" || buffData.guid.toString() == "28502" || buffData.guid.toString() == "28518" || buffData.guid.toString() == "41607") && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Warrior") || ((buffData.guid.toString() == "28521" || buffData.guid.toString() == "46840") && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Mage" && playerByNameAsc.type != "Shaman") || ((buffData.guid.toString() == "28540" || buffData.guid.toString() == "46838") && playerByNameAsc.type != "Priest" && playerByNameAsc.type != "Warlock" && playerByNameAsc.type != "Mage") || (buffData.guid.toString() == "28509" && (playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Warrior" || playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Warlock")) || ((buffData.guid.toString() == "39627" || buffData.guid.toString() == "33263" || buffData.guid.toString() == "33265") && (playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Warrior" || playerByNameAsc.type == "Hunter")) || ((buffData.guid.toString() == "33256" || buffData.guid.toString() == "44106" || buffData.guid.toString() == "40323") && playerByNameAsc.type != "Shaman" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Warrior") || (buffData.guid.toString() == "33261" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Rogue") || ((buffData.guid.toString() == "39628") && playerByNameAsc.type != "Druid") || (buffData.guid.toString() == "17538" && playerByNameAsc.type != "Warrior" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Rogue")) {
                              isPlayerACheapass = true;
                              if (suboptimalStuffFound.indexOf(buffData.name) < 0) {
                                if (headerValue == getStringForLang("foodBuff", langKeys, langTrans, "", "", "", "")) {
                                  if (badBuffFoodFound.indexOf(buffData.guid.toString()) < 0) {
                                    badBuffFoodFound.push(buffData.guid.toString());
                                    wellFedName = buffData.name;
                                  }
                                } else {
                                  suboptimalStuffFound += buffData.name + ", ";
                                }
                              }
                            }
                          } catch { }
                          if (headerValue == getStringForLang("battleElixir", langKeys, langTrans, "", "", "", "") || headerValue == getStringForLang("guardianElixir", langKeys, langTrans, "", "", "", ""))
                            bossCovered.push(fight.id);
                        }
                        var name = spellIdsCell.toString().split(" [")[1];
                        if (name != null && name.length > 0 && name.indexOf("*") > -1) {
                          if (names.indexOf(getStringForLang(name.replace("]", "").replace("*", ""), langKeys, langTrans, "", "", "", "") + "*") < 0) {
                            names += getStringForLang(name.replace("]", "").replace("*", ""), langKeys, langTrans, "", "", "", "") + "*,";
                          }
                        }
                        else if (name != null && name.length > 0) {
                          if (names.indexOf(getStringForLang(name.replace("]", ""), langKeys, langTrans, "", "", "", "")) < 0) {
                            names += getStringForLang(name.replace("]", ""), langKeys, langTrans, "", "", "", "") + ",";
                          }
                        }
                      }
                    })
                  }
                })
              }
            })
          })
          if (bossIdsFound != null && bossIdsFound.length > 0) {
            if (names.length > 0) {
              playerRow[0].push(Math.round(bossIdsFound.length * 100 / playerFightCount) + "% (" + names.substr(0, names.length - 1) + ")");
              if (isPlayerACheapass)
                sheet.getRange(firstPlayerNameRow + playersFound, firstPlayerNameColumn + spellIdsColumnCount + 3).setFontWeight("bold").setFontStyle("italic").setBackground("#cccccc");
            } else {
              playerRow[0].push(Math.round(bossIdsFound.length * 100 / playerFightCount) + "%");
              if (isPlayerACheapass)
                sheet.getRange(firstPlayerNameRow + playersFound, firstPlayerNameColumn + spellIdsColumnCount + 3).setFontWeight("bold").setFontStyle("italic").setBackground("#cccccc");
            }
          } else {
            playerRow[0].push("0%");
          }
        } else if (headerValue == getStringForLang("weaponEnhancement", langKeys, langTrans, "", "", "", "")) {
          var usedTemporaryWeaponEnchant = 0;
          var atLeastOneConsecrationItem = "no";
          var isPlayerACheapass = false;
          var found = 0;
          var total = 0;
          bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
            var increaseTotal = false;
            var increaseFound = false;
            if (increaseTotal == false && bossSummaryData[0].playerDetails != null && bossSummaryData[0].playerDetails.dps != null && bossSummaryData[0].playerDetails.dps.length > 0) {
              bossSummaryData[0].playerDetails.dps.forEach(function (playerInfo, playerInfoCount) {
                if (playerInfo.name == playerByNameAsc.name) {
                  if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                    increaseTotal = true;
                    playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                      if (item.id != null && item.id.toString().length > 0) {
                        if (jcNeckFound.indexOf(bossSummaryData[1] < 0) && (item.id == "20966" || item.id == "24106" || item.id == "24110" || item.id == "24114" || item.id == "24116" || item.id == "24117" || item.id == "24121" || item.id == "24092" || item.id == "24093" || item.id == "24095" || item.id == "24097" || item.id == "24098")) {
                          jcNeckFound.push(bossSummaryData[1]);
                        }
                        if (item.slot != null && item.slot.toString() == "15" || item.slot.toString() == "16") {
                          if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                            increaseTotal = false;
                          } else {
                            if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266" && item.temporaryEnchant.toString() != "283" && item.temporaryEnchant.toString() != "284" && item.temporaryEnchant.toString() != "525" && item.temporaryEnchant.toString() != "563" && item.temporaryEnchant.toString() != "564" && item.temporaryEnchant.toString() != "1669" && item.temporaryEnchant.toString() != "1783" && item.temporaryEnchant.toString() != "2636" && item.temporaryEnchant.toString() != "2638" && item.temporaryEnchant.toString() != "2639") {
                              if ((item.temporaryEnchant.toString() == "2684" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Rogue" && playerByNameAsc.type != "Warrior" && playerByNameAsc.type != "Shaman" && playerByNameAsc.type != "Paladin") || (item.temporaryEnchant.toString() == "2685" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Mage" && playerByNameAsc.type != "Priest" && playerByNameAsc.type != "Warlock" && playerByNameAsc.type != "Shaman" && playerByNameAsc.type != "Paladin") || (item.temporaryEnchant.toString() == "2677" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Priest") || (item.temporaryEnchant.toString() == "2678" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Priest") || item.temporaryEnchant.toString() == "2627" || item.temporaryEnchant.toString() == "2625" || item.temporaryEnchant.toString() == "2626" || item.temporaryEnchant.toString() == "2624" || item.temporaryEnchant.toString() == "2623" || (item.temporaryEnchant.toString() == "2712" && playerByNameAsc.type != "Hunter") || item.temporaryEnchant.toString() == "1643" || item.temporaryEnchant.toString() == "2954" || item.temporaryEnchant.toString() == "13" || item.temporaryEnchant.toString() == "40" || item.temporaryEnchant.toString() == "20" || item.temporaryEnchant.toString() == "1703" || item.temporaryEnchant.toString() == "14" || item.temporaryEnchant.toString() == "19" || item.temporaryEnchant.toString() == "483" || item.temporaryEnchant.toString() == "484") {
                                isPlayerACheapass = true;
                                if (suboptimalStuffFound.indexOf(item.temporaryEnchantName) < 0)
                                  suboptimalStuffFound += item.temporaryEnchantName + ", ";
                              } else if (item.temporaryEnchant.toString() == "2684" || item.temporaryEnchant.toString() == "2685")
                                atLeastOneConsecrationItem = "yes";
                              increaseFound = true;
                            }
                          }
                        }
                      }
                    })
                  }
                }
              })
            }
            if (increaseTotal == false && bossSummaryData[0].playerDetails != null && bossSummaryData[0].playerDetails.healers != null && bossSummaryData[0].playerDetails.healers.length > 0) {
              bossSummaryData[0].playerDetails.healers.forEach(function (playerInfo, playerInfoCount) {
                if (playerInfo.name == playerByNameAsc.name) {
                  if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                    increaseTotal = true;
                    playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                      if (item.id != null && item.id.toString().length > 0) {
                        if (jcNeckFound.indexOf(bossSummaryData[1] < 0) && (item.id == "20966" || item.id == "24106" || item.id == "24110" || item.id == "24114" || item.id == "24116" || item.id == "24117" || item.id == "24121" || item.id == "24092" || item.id == "24093" || item.id == "24095" || item.id == "24097" || item.id == "24098")) {
                          jcNeckFound.push(bossSummaryData[1]);
                        }
                        if (item.slot != null && item.slot.toString() == "15" || item.slot.toString() == "16") {
                          if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                            increaseTotal = false;
                          } else {
                            if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266" && item.temporaryEnchant.toString() != "283" && item.temporaryEnchant.toString() != "284" && item.temporaryEnchant.toString() != "525" && item.temporaryEnchant.toString() != "563" && item.temporaryEnchant.toString() != "564" && item.temporaryEnchant.toString() != "1669" && item.temporaryEnchant.toString() != "1783" && item.temporaryEnchant.toString() != "2636" && item.temporaryEnchant.toString() != "2638" && item.temporaryEnchant.toString() != "2639") {
                              if ((item.temporaryEnchant.toString() == "2684" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Rogue" && playerByNameAsc.type != "Warrior" && playerByNameAsc.type != "Shaman" && playerByNameAsc.type != "Paladin") || (item.temporaryEnchant.toString() == "2685" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Mage" && playerByNameAsc.type != "Priest" && playerByNameAsc.type != "Warlock" && playerByNameAsc.type != "Shaman" && playerByNameAsc.type != "Paladin") || (item.temporaryEnchant.toString() == "2677" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Priest") || (item.temporaryEnchant.toString() == "2678" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Priest") || item.temporaryEnchant.toString() == "2627" || item.temporaryEnchant.toString() == "2625" || item.temporaryEnchant.toString() == "2626" || item.temporaryEnchant.toString() == "2624" || item.temporaryEnchant.toString() == "2623" || (item.temporaryEnchant.toString() == "2712" && playerByNameAsc.type != "Hunter") || item.temporaryEnchant.toString() == "1643" || item.temporaryEnchant.toString() == "2954" || item.temporaryEnchant.toString() == "13" || item.temporaryEnchant.toString() == "40" || item.temporaryEnchant.toString() == "20" || item.temporaryEnchant.toString() == "1703" || item.temporaryEnchant.toString() == "14" || item.temporaryEnchant.toString() == "19" || item.temporaryEnchant.toString() == "483" || item.temporaryEnchant.toString() == "484") {
                                isPlayerACheapass = true;
                                if (suboptimalStuffFound.indexOf(item.temporaryEnchantName) < 0)
                                  suboptimalStuffFound += item.temporaryEnchantName + ", ";
                              } else if (item.temporaryEnchant.toString() == "2684" || item.temporaryEnchant.toString() == "2685")
                                atLeastOneConsecrationItem = "yes";
                              increaseFound = true;
                            }
                          }
                        }
                      }
                    })
                  }
                }
              })
            }
            if (increaseTotal == false && bossSummaryData[0].playerDetails != null && bossSummaryData[0].playerDetails.tanks != null && bossSummaryData[0].playerDetails.tanks.length > 0) {
              bossSummaryData[0].playerDetails.tanks.forEach(function (playerInfo, playerInfoCount) {
                if (playerInfo.name == playerByNameAsc.name) {
                  if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                    increaseTotal = true;
                    playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                      if (item.id != null && item.id.toString().length > 0) {
                        if (jcNeckFound.indexOf(bossSummaryData[1] < 0) && (item.id == "20966" || item.id == "24106" || item.id == "24110" || item.id == "24114" || item.id == "24116" || item.id == "24117" || item.id == "24121" || item.id == "24092" || item.id == "24093" || item.id == "24095" || item.id == "24097" || item.id == "24098")) {
                          jcNeckFound.push(bossSummaryData[1]);
                        }
                        if (item.slot != null && item.slot.toString() == "15" || item.slot.toString() == "16") {
                          if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                            increaseTotal = false;
                          } else {
                            if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266" && item.temporaryEnchant.toString() != "283" && item.temporaryEnchant.toString() != "284" && item.temporaryEnchant.toString() != "525" && item.temporaryEnchant.toString() != "563" && item.temporaryEnchant.toString() != "564" && item.temporaryEnchant.toString() != "1669" && item.temporaryEnchant.toString() != "1783" && item.temporaryEnchant.toString() != "2636" && item.temporaryEnchant.toString() != "2638" && item.temporaryEnchant.toString() != "2639") {
                              if ((item.temporaryEnchant.toString() == "2684" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Rogue" && playerByNameAsc.type != "Warrior" && playerByNameAsc.type != "Shaman" && playerByNameAsc.type != "Paladin") || (item.temporaryEnchant.toString() == "2685" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Mage" && playerByNameAsc.type != "Priest" && playerByNameAsc.type != "Warlock" && playerByNameAsc.type != "Shaman" && playerByNameAsc.type != "Paladin") || (item.temporaryEnchant.toString() == "2677" && playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Priest") || (item.temporaryEnchant.toString() == "2678" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Priest") || item.temporaryEnchant.toString() == "2627" || item.temporaryEnchant.toString() == "2625" || item.temporaryEnchant.toString() == "2626" || item.temporaryEnchant.toString() == "2624" || item.temporaryEnchant.toString() == "2623" || (item.temporaryEnchant.toString() == "2712" && playerByNameAsc.type != "Hunter") || item.temporaryEnchant.toString() == "1643" || item.temporaryEnchant.toString() == "2954" || item.temporaryEnchant.toString() == "13" || item.temporaryEnchant.toString() == "40" || item.temporaryEnchant.toString() == "20" || item.temporaryEnchant.toString() == "1703" || item.temporaryEnchant.toString() == "14" || item.temporaryEnchant.toString() == "19" || item.temporaryEnchant.toString() == "483" || item.temporaryEnchant.toString() == "484") {
                                isPlayerACheapass = true;
                                if (suboptimalStuffFound.indexOf(item.temporaryEnchantName) < 0)
                                  suboptimalStuffFound += item.temporaryEnchantName + ", ";
                              } else if (item.temporaryEnchant.toString() == "2684" || item.temporaryEnchant.toString() == "2685")
                                atLeastOneConsecrationItem = "yes";
                              increaseFound = true;
                            }
                          }
                        }
                      }
                    })
                  }
                }
              })
            }
            if (increaseFound)
              found++;
            else if (playerByNameAsc.type == "Rogue" && increaseTotal) {
              bossDamageDataAll.forEach(function (bossDamageData, bossDamageDataCount) {
                if (bossDamageData[1] == bossSummaryData[1]) {
                  bossDamageData[0].entries.forEach(function (bossDamageData, bossDamageDataCount) {
                    if (bossDamageData != null && bossDamageData.id != null && bossDamageData.id == playerByNameAsc.id) {
                      found++;
                    }
                  })
                }
              })
            }
            if (increaseTotal)
              total++;
          })
          usedTemporaryWeaponEnchant = Math.round(found * 100 / total) + "%";
          if (total > 0) {
            playerRow[0].push(usedTemporaryWeaponEnchant);
            if (isPlayerACheapass)
              sheet.getRange(firstPlayerNameRow + playersFound, firstPlayerNameColumn + spellIdsColumnCount + 3).setFontWeight("bold").setFontStyle("italic").setBackground("#cccccc");
          } else
            playerRow[0].push("0%");
          if (atLeastOneConsecrationItem == "yes")
            sheet.getRange(firstPlayerNameRow + playersFound, firstPlayerNameColumn + spellIdsColumnCount + 3).setBackground("#666666");
        } else if (headerValue == getStringForLang("nrBossFightsWithNeck", langKeys, langTrans, "", "", "", "")) {
          var bossesFound = [];
          allNecksBuffData.auras.forEach(function (neckbuffData, neckbuffDataCount) {
            if (neckbuffData.guid != null) {
              spellIdsColumn.forEach(function (spellIdsCell, spellIdsCellCount) {
                if (spellIdsCell.toString().split(" [")[1].replace("]", "") == neckbuffData.guid.toString()) {
                  neckbuffData.bands.forEach(function (neckBuffBand, neckBuffBandCount) {
                    allFightsData.fights.forEach(function (fight, fightCount) {
                      if (((neckBuffBand.endTime >= fight.start_time && neckBuffBand.endTime <= fight.end_time) || (neckBuffBand.startTime >= fight.startTime && neckBuffBand.endTime <= fight.end_time)) && fight.start_time >= sectionToLookAtStart && fight.end_time <= sectionToLookAtEnd && bossesFound.indexOf(fight.id.toString()) < 0) {
                        bossesFound.push(fight.id.toString());
                      }
                    })
                  })
                }
              })
            }
          })
          if (bossesFound.length > 0) {
            if ((jcNeckFound.length - bossesFound.length) > 0) {
              playerRow[0].push(bossesFound.length + " --- " + getStringForLang("inactiveNeckEquipped", langKeys, langTrans, (jcNeckFound.length - bossesFound.length).toString(), "", "", ""));
              var noteText = "";
              jcNeckFound.forEach(function (jcNeckFoundOnBoss, jcNeckFoundOnBossCount) {
                if (bossesFound.indexOf(jcNeckFoundOnBoss) < 0) {
                  allFightsData.fights.forEach(function (fight, fightCount) {
                    if (fight.id == jcNeckFoundOnBoss && noteText.indexOf(fight.name + " (" + getStringForLang("fightId", langKeys, langTrans, "", "", "", "") + " " + fight.id.toString() + ")") < 0) {
                      noteText += fight.name + " (" + getStringForLang("fightId", langKeys, langTrans, "", "", "", "") + " " + fight.id.toString() + "), ";
                    }
                  })
                }
              })
              sheet.getRange(firstPlayerNameRow + playersFound, firstPlayerNameColumn + 3 + spellIdsColumnCount, 1, 1).setNote(noteText.substr(0, noteText.length - 2));
            }
            else
              playerRow[0].push(bossesFound.length);
          } else
            playerRow[0].push("-");
        }
        sheet.getRange(firstPlayerNameRow + playersFound, firstPlayerNameColumn + 3, 1, spellIdsColumnCount + 1).setValues(playerRow);
      })
      if (suboptimalStuffFound.length > 0 || badBuffFoodFound.length > 0) {
        var rangeSuboptimal = sheet.getRange(firstPlayerNameRow + playersFound, firstPlayerNameColumn + 3 + spellIds.length, 1, 1);
        var initialLength = suboptimalStuffFound.length;
        badBuffFoodFound.forEach(function (badBuffFood, badBuffFoodCount) {
          suboptimalStuffFound += wellFedName + ", ";
        })
        var richValue = SpreadsheetApp.newRichTextValue().setText(suboptimalStuffFound.substr(0, suboptimalStuffFound.length - 2));
        var wowheadBaseUrl = "https://tbc.wowhead.com/spell=";
        if (lang != "EN")
          wowheadBaseUrl = wowheadBaseUrl.replace("https://tbc", "https://" + lang.toLowerCase() + ".tbc");
        badBuffFoodFound.forEach(function (badBuffFood, badBuffFoodCount) {
          richValue.setLinkUrl(initialLength + (wellFedName.length + 2) * badBuffFoodCount, initialLength + (wellFedName.length + 2) * badBuffFoodCount + wellFedName.length, wowheadBaseUrl + badBuffFood);
        })
        rangeSuboptimal.setRichTextValue(richValue.build());
      }

      playersFound++;
    }
  })
  var sheets = ss.getSheets();
  for (var c = sheets.length - 1; c >= 0; c--) {
    var sheetNameSearch = sheets[c].getName();
    if (sheetNameSearch.indexOf("buffConsumables") > -1) {
      ss.deleteSheet(sheets[c]);
    }
  }
}

function getEntriesForHeader(headers, headerToSearch, conf) {
  var columnNumber = -1;
  headers.forEach(function (header, headerCount) {
    if (header == headerToSearch)
      columnNumber = headerCount;
  })
  var entriesRaw = conf.getRange(2, columnNumber + 1, 200, 1).getValues();
  return entriesRaw.reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
  }, []);
}

function getStylesForHeader(headers, headerToSearch, conf) {
  var columnNumber = -1;
  headers.forEach(function (header, headerCount) {
    if (header == headerToSearch)
      columnNumber = headerCount;
  })
  var entriesRaw = conf.getRange(2, columnNumber + 1, 200, 1).getFontStyles();
  return entriesRaw.reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
  }, []);
}