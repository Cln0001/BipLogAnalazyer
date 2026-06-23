function populateShadowResistance() {
  var firstPlayerNameRow = 5;
  var firstPlayerNameColumn = 2;
  var confSpreadSheet = SpreadsheetApp.openById('1pIbbPkn9i5jxyQ60Xt86fLthtbdCAmFriIpPSvmXiu0');
  var priestShadowResiBuffFixed = confSpreadSheet.getSheetByName("currentVersion").getRange(3, 1).getValue();

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

  if (priestShadowResiBuffFixed == "yes")
    sheet.getRange(2, 9).setValue("");
  else
    sheet.getRange(2, 9).setValue("=VLOOKUP(\"priestBuffNotTracked\",trans!$I$2:$J$1000,2,FALSE)");

  instructionsSheet.getRange(26, 2).setValue("");
  instructionsSheet.getRange(27, 2).setValue("");

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, 34, 21).clearContent();
  if (darkMode)
    sheet.getRange(1, 1, 38, 22).setBackground("#d9d9d9");
  else
    sheet.getRange(1, 1, 38, 22).setBackground("white");
  sheet.getRange(2, 3, 1, 1).setBackground("#cccccc");

  var bossName = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("selectBoss", langKeys, langTrans, "", "", "", "")).useRegularExpression(true).findNext(), 1).getValue();
  var bossId = 607;
  if (bossName != null && bossName.length > 0 && bossName.indexOf(getStringForLang("Kaz", langKeys, langTrans, "", "", "", "")) > -1)
    bossId = 320;
  else if (bossName != null && bossName.length > 0 && bossName.indexOf(getStringForLang("Azga", langKeys, langTrans, "", "", "", "")) > -1)
    bossId = 321;

  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue().replace(/\s/g, "");
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var onlyFightNr = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("onlyFightId", langKeys, langTrans, "", "", "", "")).findNext(), 1).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var information = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 2), 1);
  shiftRangeByColumns(sheet, information, 1).clearContent()
  var confShadowResistanceConfig = ss.getSheetByName("shadow resistance config");

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

  var allPlayersUrl = baseUrl + "report/tables/casts/" + logId + apiKeyString + "&start=0&end=999999999999";

  var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString;
  var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));
  var baseSheetName = getStringForLang("shadowResiTab", langKeys, langTrans, "", "", "", "") + " (" + getStringForLang("Mother", langKeys, langTrans, "", "", "", "") + ")";
  if (bossName != null && bossName.length > 0 && bossName.indexOf(getStringForLang("Kaz", langKeys, langTrans, "", "", "", "")) > -1)
    baseSheetName = getStringForLang("shadowResiTab", langKeys, langTrans, "", "", "", "") + " (" + getStringForLang("Kaz", langKeys, langTrans, "", "", "", "") + ")";
  else if (bossName != null && bossName.length > 0 && bossName.indexOf(getStringForLang("Azga", langKeys, langTrans, "", "", "", "")) > -1)
    baseSheetName = getStringForLang("shadowResiTab", langKeys, langTrans, "", "", "", "") + " (" + getStringForLang("Azga", langKeys, langTrans, "", "", "", "") + ")";
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

  var shadowResiInfoIdsRaw = confShadowResistanceConfig.getRange(1, 1, 1500, 1).getValues();
  var shadowResiInfoIds = shadowResiInfoIdsRaw.reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
  }, []);

  var shadowResiInfoFRRaw = confShadowResistanceConfig.getRange(1, 2, 1500, 1).getValues();
  var shadowResiInfoFRs = shadowResiInfoFRRaw.reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
  }, []);

  var fightDataArr = [];
  var fightDataIndexArr = [];
  var playersDone = 0;
  var fightIDToEvaluate = "";
  var longestFight = "";
  var longestFightLength = 0;
  allFightsData.fights.forEach(function (fight, fightRawCount) {
    if (((onlyFightNr == null || onlyFightNr.toString().length == 0) && fight.boss != null && fight.boss.toString().endsWith(bossId.toString())) || (onlyFightNr != null && onlyFightNr.toString().length > 0 && fight.id.toString() == onlyFightNr.toString())) {
      if (Number(fight.fightPercentage) == 100 || fight.kill == true)
        fightIDToEvaluate = fight.id.toString();
      else if ((fight.end_time - fight.start_time) > longestFightLength) {
        longestFightLength = fight.end_time - fight.start_time;
        longestFight = fight.id.toString();
      }
    }
  })
  if (fightIDToEvaluate == "" && longestFight != "")
    fightIDToEvaluate = longestFight;
  var rangeBoss = sheet.getRange(firstPlayerNameRow - 2, firstPlayerNameColumn);
  if (fightIDToEvaluate == "")
    rangeBoss.setValue(getStringForLang("noFightFound", langKeys, langTrans, "", "", "", ""))
  var allPlayersData = JSON.parse(UrlFetchApp.fetch(allPlayersUrl));
  const allPlayersByNameAsc = sortByProperty(sortByProperty(allPlayersData.entries, "name"), "type");
  allPlayersByNameAsc.forEach(function (playerByNameAsc, playerCountByNameAsc) {
    if ((playerByNameAsc.type == "Druid" || playerByNameAsc.type == "Hunter" || playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Priest" || playerByNameAsc.type == "Paladin" || playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Shaman" || playerByNameAsc.type == "Warlock" || playerByNameAsc.type == "Warrior") && playerByNameAsc.total > 20) {
      var fightCount = 0;
      allFightsData.fights.forEach(function (fight, fightRawCount) {
        if (fight.id.toString() == fightIDToEvaluate) {
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
              var shadowResistanceTotal = 0;
              var shadowResistanceByBuffs = 0;
              var shadowResistanceByGear = 0;
              var urlPlayerBuffs = baseUrl + "report/tables/buffs/" + logId + apiKeyString + "&start=" + fight.start_time + "&end=" + fight.end_time + "&sourceid=" + playerByNameAsc.id;
              var priestBuffFound = 0;
              buffsData = JSON.parse(UrlFetchApp.fetch(urlPlayerBuffs));
              buffsData.auras.forEach(function (playerBuff, playerBuffCount) {
                if (playerBuff.guid != null) {
                  if (priestShadowResiBuffFixed == "yes") {
                    if (playerBuff.guid.toString() == "25433") {
                      if (priestBuffFound < 70) {
                        shadowResistanceByBuffs += 70;
                        priestBuffFound = 70;
                      }
                    } else if (playerBuff.guid.toString() == "10958") {
                      if (priestBuffFound < 60) {
                        shadowResistanceByBuffs += 60;
                        priestBuffFound = 60;
                      }
                    } else if (playerBuff.guid.toString() == "10957") {
                      if (priestBuffFound < 45) {
                        shadowResistanceByBuffs += 45;
                        priestBuffFound = 45;
                      }
                    } else if (playerBuff.guid.toString() == "976") {
                      if (priestBuffFound < 30) {
                        shadowResistanceByBuffs += 30;
                        priestBuffFound = 30;
                      }
                    } else if (playerBuff.guid.toString() == "39374") {
                      if (priestBuffFound < 70) {
                        shadowResistanceByBuffs += 70;
                        priestBuffFound = 70;
                      }
                    } else if (playerBuff.guid.toString() == "27683") {
                      if (priestBuffFound < 60) {
                        shadowResistanceByBuffs += 60;
                        priestBuffFound = 60;
                      }
                    }
                    else if (playerBuff.guid.toString() == "27125")
                      shadowResistanceByBuffs += 18;
                    else if (playerBuff.guid.toString() == "22783")
                      shadowResistanceByBuffs += 15;
                    else if (playerBuff.guid.toString() == "22782")
                      shadowResistanceByBuffs += 10;
                    else if (playerBuff.guid.toString() == "6117")
                      shadowResistanceByBuffs += 5;
                    else if (playerBuff.guid.toString() == "27260")
                      shadowResistanceByBuffs += 18;
                    else if (playerBuff.guid.toString() == "11735")
                      shadowResistanceByBuffs += 15;
                    else if (playerBuff.guid.toString() == "11734")
                      shadowResistanceByBuffs += 12;
                    else if (playerBuff.guid.toString() == "11733")
                      shadowResistanceByBuffs += 9;
                    else if (playerBuff.guid.toString() == "1086")
                      shadowResistanceByBuffs += 6;
                    else if (playerBuff.guid.toString() == "706")
                      shadowResistanceByBuffs += 3;
                  }
                  if (playerBuff.guid.toString() == "42735")
                    shadowResistanceByBuffs += 35;
                  else if (playerBuff.guid.toString() == "17629")
                    shadowResistanceByBuffs += 25;
                  else if (playerBuff.guid.toString() == "45619")
                    shadowResistanceByBuffs += 8;
                  else if (playerBuff.guid.toString() == "1138")
                    shadowResistanceByBuffs += 10;
                  else if (playerBuff.guid.toString() == "11371")
                    shadowResistanceByBuffs += 10;
                }
              })
              shadowResistanceTotal += shadowResistanceByBuffs;
              if (player.gear != null && player.gear.length > 0) {
                player.gear.forEach(function (item, itemCount) {
                  if (item.id != null && item.id.toString().length > 0 && item.id.toString() != "0" && item.slot != 3 && item.slot != 18) {
                    var gearShadowResi = searchEntryForId(shadowResiInfoIds, shadowResiInfoFRs, item.id.toString());
                    var enchantShadowResi = 0;
                    var gemShadowResi = 0;
                    if (item.permanentEnchant != null && item.permanentEnchant.toString().length > 1) {
                      if (item.permanentEnchant.toString() == "804") {
                        enchantShadowResi = 10;
                      } else if (item.permanentEnchant.toString() == "1888") {
                        enchantShadowResi = 5;
                      } else if (item.permanentEnchant.toString() == "2984") {
                        enchantShadowResi = 8;
                      } else if (item.permanentEnchant.toString() == "3009") {
                        enchantShadowResi = 20;
                      } else if (item.permanentEnchant.toString() == "2998") {
                        enchantShadowResi = 7;
                      } else if (item.permanentEnchant.toString() == "2664") {
                        enchantShadowResi = 7;
                      } else if (item.permanentEnchant.toString() == "1441") {
                        enchantShadowResi = 15;
                      } else if (item.permanentEnchant.toString() == "2683") {
                        enchantShadowResi = 10;
                      }
                      shadowResistanceTotal += enchantShadowResi;
                    }
                    if (item.gems != null) {
                      item.gems.forEach(function (gem, gemCount) {
                        if (gem.id.toString() == "22459") {
                          gemShadowResi += 4;
                        } else if (gem.id.toString() == "22460") {
                          gemShadowResi += 3;
                        }
                      })
                      shadowResistanceTotal += gemShadowResi;
                    }
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
                    var rangeTarget = sheet.getRange(playersDone + firstPlayerNameRow, firstPlayerNameColumn + 4 + fightCount + itemPos);
                    if (gearShadowResi != "") {
                      shadowResistanceTotal += Number(gearShadowResi);
                      shadowResistanceByGear += Number(gearShadowResi);
                      confShadowResistanceConfig.createTextFinder(item.id.toString()).useRegularExpression(true).findNext().copyTo(rangeTarget, { formatOnly: true });
                      if ((enchantShadowResi + gemShadowResi) > 0) {
                        rangeTarget.setValue(item.name + " (~" + gearShadowResi + " " + getStringForLang("SRtext", langKeys, langTrans, "", "", "", "") + ") +" + (enchantShadowResi + gemShadowResi) + " " + getStringForLang("SRtext", langKeys, langTrans, "", "", "", ""));
                        shadowResistanceByGear += (enchantShadowResi + gemShadowResi);
                      } else
                        rangeTarget.setValue(item.name + " (~" + gearShadowResi + " " + getStringForLang("SRtext", langKeys, langTrans, "", "", "", "") + ")");
                    } else if ((enchantShadowResi + gemShadowResi) > 0) {
                      rangeTarget.setValue(item.name + " +" + (enchantShadowResi + gemShadowResi) + " " + getStringForLang("SRtext", langKeys, langTrans, "", "", "", ""));
                      shadowResistanceByGear += (enchantShadowResi + gemShadowResi);
                    }
                  }
                })
              }
              if (fightCount == 0) {
                var range = sheet.getRange(playersDone + firstPlayerNameRow, firstPlayerNameColumn);
                range.setValue(player.name);
                range.setBackground(getColourForPlayerClass(player.type));
                playersDone++;
              }
              sheet.getRange(playersDone + firstPlayerNameRow - 1, firstPlayerNameColumn + 1 + fightCount).setValue(shadowResistanceTotal);
              sheet.getRange(playersDone + firstPlayerNameRow - 1, firstPlayerNameColumn + 2 + fightCount).setValue(shadowResistanceByGear);
              sheet.getRange(playersDone + firstPlayerNameRow - 1, firstPlayerNameColumn + 3 + fightCount).setValue(shadowResistanceByBuffs);
              fightCount++;
            }
          })
        }
      })
    }
  })
}