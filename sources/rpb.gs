function generateAllSheet() {
  var codeVersion = '1.6.0';
  var confSpreadSheet = SpreadsheetApp.openById('1pIbbPkn9i5jxyQ60Xt86fLthtbdCAmFriIpPSvmXiu0');
  var currentVersion = confSpreadSheet.getSheetByName("currentVersion").getRange(1, 1).getValue();
  var maxColumns = 69;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var instructionsSheet = ss.getSheetByName("Instructions");

  var confSpellHasteConfig = ss.getSheetByName("spell haste config");

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
  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue();
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var information = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), 2), 1);

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^" + getStringForLang("email", langKeys, langTrans, "", "", "", "") + "$").useRegularExpression(true).findNext(), -1), 4).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  if (darkMode)
    sheet.getRange(4, 63).setFontColor("#d9d9d9").setValue("");
  else
    sheet.getRange(4, 63).setFontColor("white").setValue("");
  sheet.getRange(3, 5).setValue("");
  sheet.getRange(4, 5).setValue("");

  var onlyBosses = false;
  var onlyTrash = false;
  var noWipes = false;
  var modeSelection = sheet.createTextFinder(getStringForLang("trashBosses", langKeys, langTrans, "", "", "", "")).useRegularExpression(true).findNext();
  if (modeSelection != null) {
    var modeSelectionValue = shiftRangeByColumns(sheet, modeSelection, 1);
    if (modeSelectionValue != null) {
      var value = modeSelectionValue.getValue();
      if (value != null && (value.toString().indexOf("only bosses") > -1 || value.toString().indexOf(getStringForLang("onlyBosses", langKeys, langTrans, "", "", "", "")) > -1)) {
        onlyBosses = true;
      } else if (value != null && (value.toString().indexOf("only trash") > -1 || value.toString().indexOf(getStringForLang("onlyTrash", langKeys, langTrans, "", "", "", "")) > -1)) {
        onlyTrash = true;
      }
      if (value != null && (value.toString().indexOf("(no wipes)") > -1 || value.toString().indexOf(getStringForLang("noWipes", langKeys, langTrans, "", "", "", "")) > -1)) {
        noWipes = true;
      }
    }
  }
  var onlyFightNr = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("onlyFightId", langKeys, langTrans, "", "", "", "")).findNext(), 1).getValue();
  var manualStartAndEnd = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("startEndOptional", langKeys, langTrans, "", "", "", "")).findNext(), 1).getValue();
  if ((onlyFightNr != null && onlyFightNr.toString().length > 0) && (manualStartAndEnd != null && manualStartAndEnd.toString().length > 0)) {
    SpreadsheetApp.getUi().alert(getStringForLang("fighIdOrStartEnd", langKeys, langTrans, "", "", "", ""));
    return;
  }
  if (manualStartAndEnd != null && manualStartAndEnd.toString().length > 0)
    manualStartAndEnd = manualStartAndEnd.replaceAll(" ", "");

  var characterNames = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("characterNames", langKeys, langTrans, "", "", "", "")).findNext(), 1).getValue();

  if (currentVersion.indexOf(codeVersion) < 0) {
    SpreadsheetApp.getUi().alert(getStringForLang("sheetOutdated", langKeys, langTrans, "", "", "", ""));
  }

  cleanSheet(sheet, information, darkMode);

  if (reportPathOrId.length > 5) {
    //build urls of apiQueries
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
    var startEndStringNoFilter = "&start=0&end=999999999999";
    var startEndString = "&start=0&end=999999999999&filter=encounterid%20%21%3D%20724";
    var apiKeyString = "?translate=true&api_key=" + api_key;
    if (onlyBosses)
      apiKeyString += "&encounter=-2";
    if (onlyTrash)
      apiKeyString += "&encounter=0";
    if (noWipes)
      apiKeyString += "&wipes=2";
    var baseUrl = "https://classic.warcraftlogs.com:443/v1/"
    var baseUrlFrontEnd = "https://classic.warcraftlogs.com/reports/"
    if (lang != "EN") {
      baseUrl = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com:443/v1/";
      baseUrlFrontEnd = "https://" + lang.toLowerCase() + ".classic.warcraftlogs.com/reports/";
    }
    var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString;
    var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));
    var fightName = "";
    var lastFightName = "";
    var lastFight = "";
    var lastId = "";
    if (onlyFightNr != null && onlyFightNr.toString().length > 0) {
      allFightsData.fights.forEach(function (fight, fightCount) {
        if (fight.id.toString() == onlyFightNr && fight.start_time >= 0 && fight.end_time >= 0) {
          startEndString = "&start=" + fight.start_time + "&end=" + fight.end_time;
          fightName = fight.name;
        }
        if ((fight.boss > 0 && !onlyTrash) || onlyTrash) {
          lastFight = "&start=" + fight.start_time + "&end=" + fight.end_time;
          lastFightName = fight.name;
          lastId = fight.id.toString();
        }
      })
      if (onlyFightNr == "last") {
        startEndString = lastFight;
        fightName = lastFightName;
        onlyFightNr = lastId;
      }
    }
    if (manualStartAndEnd != null && manualStartAndEnd.toString().length > 0) {
      var startEndParts = manualStartAndEnd.split("-");
      startEndString = "&start=" + startEndParts[0] + "&end=" + startEndParts[1] + "&filter=encounterid%20%21%3D%20724";
    }
    var urlDamageTakenTop = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndString + "&options=4098&by=ability";
    var urlDebuffsTop = baseUrl + "report/tables/debuffs/" + logId + apiKeyString + startEndString + "&options=2&hostility=1&by=target";
    var urlPeopleTracked = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndString;
    var urlDebuffInfo = baseUrl + "report/tables/debuffs/" + logId + apiKeyString + startEndString + "&options=2&hostility=1&by=target&abilityid=";
    var urlPlayersOnTrash = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndString + "&encounter=0&sourceid=";
    var urlPlayersRacials = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndStringNoFilter + "&filter=ability.id%3D7744%20OR%20ability.id%3D20554%20OR%20ability.id%3D20549%20OR%20ability.id%3D20572";
    var urlPlayers = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndString + "&sourceid=";
    var urlPlayersSunderArmorOnLessThan5Stacks = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndStringNoFilter + "&filter=ability.id%3D25225%20AND%20NOT%20IN%20RANGE%20FROM%20type%20%3D%20%22applydebuffstack%22%20AND%20ability.id%20%3D%2025225%20AND%20stack%20%3D%205%20TO%20type%3D%22removedebuff%22%20AND%20ability.id%3D25225%20GROUP%20BY%20target%20ON%20target%20END%20AND%20encounterid%20%21%3D%20724&by=source";
    var urlPlayersScorchOnLessThan5Stacks = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndStringNoFilter + "&filter=ability.id%20IN%20%2827073,27074,10207,10206,10205,8446,8445,8444,2948%29%20AND%20NOT%20IN%20RANGE%20FROM%20type%20%3D%20%22applydebuffstack%22%20AND%20ability.id%20%3D%2022959%20AND%20stack%20%3D%205%20TO%20type%3D%22removedebuff%22%20AND%20ability.id%3D22959%20GROUP%20BY%20target%20END%20AND%20encounterid%20%21%3D%20724&by=source";
    var urlSummary = baseUrl + "report/tables/summary/" + logId + apiKeyString + startEndString;
    var urlDamageDone = baseUrl + "report/tables/damage-done/" + logId + apiKeyString + startEndString + "&options=2&sourceid=";
    var urlBuffsOnTrash = baseUrl + "report/tables/buffs/" + logId + apiKeyString + startEndString + "&by=target&encounter=0&targetid=";
    var urlBuffsTotal = baseUrl + "report/tables/buffs/" + logId + apiKeyString + startEndString + "&by=target&targetid=";
    var urlDeathsOnTrash = baseUrl + "report/tables/deaths/" + logId + apiKeyString + startEndString + "&encounter=0";
    var urlDeaths = baseUrl + "report/tables/deaths/" + logId + apiKeyString + startEndString + "&sourceid=";
    var urlDamageTakenOilOfImmo = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndString + "&hostility=1&abilityid=11351&by=target";
    var urlDamageTakenEngineering = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndStringNoFilter + "&hostility=1&filter=ability.id%20IN%20%2823063%2C13241%2C17291%2C30486%2C4062%2C19821%2C15239%2C19784%2C12543%2C30461%2C30217%2C39965%2C4068%2C19769%2C4100%2C30216%2C22792%2C30526%2C4072%2C19805%2C27661%2C23000%2C11350%29%20AND%20encounterid%20%21%3D%20724&by=target";
    var urlDamageTakenTotal = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndString + "&options=4134&sourceid=";
    var urlDebuffs = baseUrl + "report/tables/debuffs/" + logId + apiKeyString + startEndString + "&options=2&hostility=1&by=target&targetid=";
    var urlDebuffsApplied = baseUrl + "report/tables/debuffs/" + logId + apiKeyString + startEndString + "&hostility=1&targetid=";
    var urlDebuffsAppliedTotal = baseUrl + "report/tables/debuffs/" + logId + apiKeyString + startEndString + "&hostility=1";
    var urlDebuffsAppliedBosses = baseUrl + "report/tables/debuffs/" + logId + apiKeyString + startEndString + "&encounter=-2&hostility=1&targetid=";
    var urlDebuffsAppliedBossesJudgement = baseUrl + "report/tables/debuffs/" + logId + apiKeyString + startEndStringNoFilter + "&encounter=-2&hostility=1&filter=ability.id%20IN%20%2827164%2C27162%2C31898%2C41461%2C32220%2C27172%2C27171%2C356112%2C31896%2C27162%2C27163%2C27164%2C27165%2C31804%2C27159%2C27157%2C348702%29%20AND%20encounterid%20%21%3D%20724&targetid=";
    var urlDebuffsAppliedBossesTotal = baseUrl + "report/tables/debuffs/" + logId + apiKeyString + startEndString + "&encounter=-2&hostility=1";
    var urlHostilePlayers = baseUrl + "report/tables/damage-done/" + logId + apiKeyString + startEndString + "&targetclass=player&by=source";
    var urlHealing = baseUrl + "report/tables/healing/" + logId + apiKeyString + startEndString + "&sourceid=";
    var urlHealingTarget = baseUrl + "report/tables/healing/" + logId + apiKeyString + startEndString + "&targetid=";
    var urlDamageReflected = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndStringNoFilter + "&filter=target.name%3Dsource.name%20AND%20ability.id!%3D%27348191%27%20AND%20ability.id!%3D%2716666%27%20AND%20ability.id!%3D%2711684%27%20AND%20ability.id!%3D%2711683%27%20AND%20ability.id!%3D%271949%27%20AND%20ability.id!%3D%2726557%27%20AND%20ability.id!%3D%2728622%27%20AND%20ability.id!%3D%27290025%27%20AND%20ability.id!%3D%2727869%27%20AND%20ability.id!%3D%2716666%27%20AND%20ability.id!%3D%2713241%27AND%20ability.id!%3D%2720476%27AND%20ability.id!%3D%2732221%27AND%20ability.id!%3D%2732220%27AND%20ability.id!%3D%2730486%27AND%20ability.id!%3D%27351761%27AND%20ability.id!%3D%2737852%27AND%20ability.id!%3D%2727213%27AND%20ability.id!%3D%2738281%27AND%20ability.id!%3D%2729766%27AND%20ability.id!%3D%27348703%27AND%20ability.id!%3D%2741352%27AND%20ability.id!%3D%2740871%27AND%20ability.id!%3D%27348703%27AND%20ability.id!%3D%2741352%27AND%20ability.id!%3D%2745348%27AND%20ability.id!%3D%2741352%27AND%20ability.id!%3D%2745034%27AND%20ability.id!%3D%2745642%27%20AND%20encounterid%20%21%3D%20724";
    var urlInterrupted = baseUrl + "report/tables/interrupts/" + logId + apiKeyString + startEndString;
    var urlVTManaGain = baseUrl + "report/tables/resources-gains/" + logId + apiKeyString + startEndStringNoFilter + "&filter=ability.id%20%3D%2034919%20AND%20encounterid%20%21%3D%20724&abilityid=100&sourceid=";
    var urlShadowDamageDone = baseUrl + "report/tables/damage-done/" + logId + apiKeyString + startEndStringNoFilter + "&filter=ability.id%20IN%20%288129%2C8131%2C10874%2C10875%2C10876%2C25379%2C25380%2C8092%2C8102%2C8104%2C8105%2C8106%2C10945%2C10946%2C10947%2C25372%2C25375%2C25387%2C18807%2C17314%2C17313%2C17312%2C17311%2C15407%2C32379%2C32996%2C25368%2C25367%2C10894%2C10893%2C10892%2C2767%2C992%2C970%2C594%2C589%2C34914%2C34916%2C34917%2C25467%2C45055%29%20AND%20encounterid%20%21%3D%20724&by=source&sourceid=";
    var urlTwistsDoneOnBosses = baseUrl + "report/tables/damage-done/" + logId + apiKeyString + startEndStringNoFilter + "&abilityid=1&by=source&options=2&sourceAurasPresent=20375,31892&encounter=-2&sourceid=";
    var urlWindfuryAttacksOnTwistsDoneOnBosses = baseUrl + "report/tables/damage-done/" + logId + apiKeyString + startEndStringNoFilter + "&abilityid=1&by=source&options=2&sourceAurasPresent=20375,31892,25584&encounter=-2&sourceid=";
    var urlWindfuryAttacksOnBosses = baseUrl + "report/tables/buffs/" + logId + apiKeyString + startEndString + "&abilityid=25584&by=source&options=2&encounter=-2&sourceid=";
    var urlDamageDoneOnBosses = baseUrl + "report/tables/damage-done/" + logId + apiKeyString + startEndString + "&options=2&abilityid=1&by=source&options=2&encounter=-2&sourceid=";

    var bossString = "-3";
    if (onlyBosses)
      bossString = "-2";

    if (onlyTrash)
      bossString = "0";

    if (noWipes)
      bossString += "&wipes=2";

    var urlDamageReflectedLink = urlDamageReflected.replace(baseUrl + "report/tables/damage-taken/", baseUrlFrontEnd).replace(logId, logId + "#type=damage-taken").replace(apiKeyString, "").replace(startEndStringNoFilter, "").replace("&filter=", "&pins=2%24Off%24%23244F4B%24expression%24") + "&translate=true&boss=" + bossString + "&difficulty=0&view=events";

    var maxColumnWidth = 0;

    var interruptedData = JSON.parse(UrlFetchApp.fetch(urlInterrupted));

    var damageTakenOilOfImmoData = JSON.parse(UrlFetchApp.fetch(urlDamageTakenOilOfImmo));
    var damageTakenEngineeringData = JSON.parse(UrlFetchApp.fetch(urlDamageTakenEngineering));

    var hostilePlayersData = JSON.parse(UrlFetchApp.fetch(urlHostilePlayers));

    var urlHostilePlayersLink = urlHostilePlayers.replace(baseUrl + "report/tables/damage-done/", baseUrlFrontEnd).replace(logId, logId + "#type=damage").replace(apiKeyString, "").replace(startEndStringNoFilter, "").replace("&by=source", "&by=target") + "&translate=true&boss=" + bossString + "&difficulty=0";

    var damageReflectedData = JSON.parse(UrlFetchApp.fetch(urlDamageReflected));

    var deathsDataTrash = JSON.parse(UrlFetchApp.fetch(urlDeathsOnTrash));
    var deathsData = JSON.parse(UrlFetchApp.fetch(urlDeaths));

    var playerDataSunderArmorOnLessThan5Stacks = JSON.parse(UrlFetchApp.fetch(urlPlayersSunderArmorOnLessThan5Stacks));
    var playerDataScorchOnLessThan5Stacks = JSON.parse(UrlFetchApp.fetch(urlPlayersScorchOnLessThan5Stacks));

    var allPlayersCasting = JSON.parse(UrlFetchApp.fetch(urlPeopleTracked));
    var allPlayersCastingOnTrash = JSON.parse(UrlFetchApp.fetch(urlPeopleTracked + "&encounter=0"));

    Utilities.sleep(500);

    var debuffsAppliedDataTotal = JSON.parse(UrlFetchApp.fetch(urlDebuffsAppliedTotal));
    var debuffsAppliedDataBossesTotal = JSON.parse(UrlFetchApp.fetch(urlDebuffsAppliedBossesTotal));

    var allPlayersCastingRacials = JSON.parse(UrlFetchApp.fetch(urlPlayersRacials));
    if (allPlayersCastingRacials != null && allPlayersCastingRacials.entries != null && allPlayersCastingRacials.entries.length > 0)
      var factionName = "Horde";
    else
      var factionName = "Alliance";

    if (factionName == "Alliance")
      urlWindfuryAttacksOnTwistsDoneOnBosses = urlWindfuryAttacksOnTwistsDoneOnBosses.replace("31892", "348700");
    if (factionName == "Alliance")
      urlTwistsDoneOnBosses = urlTwistsDoneOnBosses.replace("31892", "348700");

    var totalClassCount = 0;
    var currentClass = "";
    //load general queries into datastructures
    var allPlayersByNameAsc = sortByProperty(sortByProperty(JSON.parse(UrlFetchApp.fetch(urlPeopleTracked)).entries, 'name'), "type");
    allPlayersByNameAsc.forEach(function (playerByNameAsc, playerCount) {
      if (playerByNameAsc.total > 20 || fightName != "") {
        if (currentClass != playerByNameAsc.type) {
          currentClass = playerByNameAsc.type;
          totalClassCount++;
        }
      }
    })

    maxColumns += totalClassCount;

    sheet.setRowHeights(8, sheet.getLastRow() - 7, 21);
    sheet.setRowHeights(5, 1, 38);
    sheet.setRowHeights(6, 1, 18);
    sheet.setRowHeight(7, 26);
    sheet.setColumnWidth(1, 48);
    sheet.setColumnWidth(2, 335);
    sheet.setColumnWidths(3, maxColumns - 1, 74);
    sheet.setColumnWidths(3 + maxColumns, sheet.getMaxColumns() - maxColumns - 2, 50);

    Utilities.sleep(500);

    var nameSet = false;
    allFightsData.fights.forEach(function (fight, fightCount) {
      if (fight.zoneName != null && fight.zoneName.length > 0 && !nameSet) {
        if (fightName != "")
          sheet.getRange(information.getRow() + 1, information.getColumn() + 1).setValue(fight.zoneName + " (" + getStringForLang("onlyOnly", langKeys, langTrans, "", "", "", "") + fightName + ")");
        else
          sheet.getRange(information.getRow() + 1, information.getColumn() + 1).setValue(fight.zoneName);
        nameSet = true;
      }
    })
    if (allFightsData.zone != null && allFightsData.zone > 0 && (allFightsData.zone < 1007 || (allFightsData.zone >= 2000 && allFightsData.zone < 2007)))
      SpreadsheetApp.getUi().alert(getStringForLang("vanillaExecution", langKeys, langTrans, "", "", "", ""));
    else if (allFightsData.zone <= 0)
      SpreadsheetApp.getUi().alert(getStringForLang("zoneNotRecognized", langKeys, langTrans, "", "", "", ""));

    var raidDuration = 0;
    var returnVal = getRaidStartAndEnd(allFightsData, ss, baseUrl + "report/events/summary/" + logId + apiKeyString, confSpreadSheet);
    var zonesFound = [];
    if (returnVal != null && returnVal.zonesFound != null) {
      zonesFound = returnVal.zonesFound;
    }
    var zoneTimesString = " (";
    if (zonesFound != null && zonesFound.length > 0) {
      zonesFound.forEach(function (raidZone, raidZoneCount) {
        zoneTimesString += raidZone[5] + " " + getStringForLang("in", langKeys, langTrans, "", "", "", "") + " ";
        if (raidZone[10] > 0) {
          zoneTimesString += getStringForTimeStamp(raidZone[10], true) + ", ";
        } else {
          zoneTimesString += getStringForTimeStamp(raidZone[2] - raidZone[1], true) + ", ";
        }
        raidDuration += raidZone[4] - raidZone[3];
      })
      if (manualStartAndEnd != null && manualStartAndEnd.toString().length > 0) {
        raidDuration = Number(manualStartAndEnd.split("-")[1]) - Number(manualStartAndEnd.split("-")[0]);
      }
      zoneTimesString = zoneTimesString.substr(0, zoneTimesString.length - 2);
      if (zoneTimesString.length > 0)
        sheet.getRange(information.getRow(), information.getColumn() + 1).setValue(allFightsData.title + zoneTimesString + ")");
      else
        sheet.getRange(information.getRow(), information.getColumn() + 1).setValue(allFightsData.title);
    } else
      SpreadsheetApp.getUi().alert(getStringForLang("noRaidZone", langKeys, langTrans, "", "", "", ""));

    var sheetName = getStringForLang("all", langKeys, langTrans, "", "", "", "");
    var spreadsheetNameAppendix = "";
    if (includeReportTitleInSheetNames.indexOf("yes") > -1)
      sheetName += " " + allFightsData.title;
    if (onlyBosses) {
      spreadsheetNameAppendix += " (" + getStringForLang("onlyBosses", langKeys, langTrans, "", "", "", "") + ")";
    }
    if (onlyTrash) {
      spreadsheetNameAppendix += " (" + getStringForLang("onlyTrash", langKeys, langTrans, "", "", "", "") + ")";
    }
    if (noWipes) {
      spreadsheetNameAppendix += " (" + getStringForLang("noWipes", langKeys, langTrans, "", "", "", "") + ")";
    }
    if (darkMode)
      sheet.getRange(4, 65).setFontColor("#d9d9d9").setValue(spreadsheetNameAppendix);
    else
      sheet.getRange(4, 65).setFontColor("white").setValue(spreadsheetNameAppendix);
    nameSheetSafely(sheet, sheetName);
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

    var sheets = ss.getSheets();
    for (var c = sheets.length - 1; c >= 0; c--) {
      var sheetNameSearch = sheets[c].getName();
      if (sheetNameSearch.indexOf("Caster") > -1 || sheetNameSearch.indexOf("Healer") > -1 || sheetNameSearch.indexOf("Physical") > -1 || sheetNameSearch.indexOf("Tank") > -1 || sheetNameSearch.indexOf("configNew") > -1 || sheetNameSearch.indexOf(getStringForLang("Caster", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Healer", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Physical", langKeys, langTrans, "", "", "", "")) > -1 || sheetNameSearch.indexOf(getStringForLang("Tank", langKeys, langTrans, "", "", "", "")) > -1) {
        ss.deleteSheet(sheets[c]);
      }
    }

    var rnd = (Math.floor(Math.random() * 100000000) + 1).toString();
    var conf = copyConfigSheet(rnd, confSpreadSheet, ss);

    const configPatterns = {
      singleTargetCasts: "^singleTargetCasts \\[",
      aoeCasts: "^aoeCasts \\[",
      classCooldowns: "^classCooldowns \\[",
      secondsActive: "^secondsActive \\[",

      showDamageReflectRow: "showDamageReflectRow \\[",
      showFriendlyFireRow: "showFriendlyFireRow \\[",
      showDeathCountRow: "showDeathCountRow \\[",
      showInterruptedSpells: "showInterruptedSpells \\[",
      showInterruptedSpellsNamesRow: "showInterruptedSpellsNamesRow \\[",
      showConditionalFormattingDamageTaken: "showConditionalFormattingDamageTaken \\[",
      showOilOfImmolationDmg: "showDamageDoneWithOilOfImmolation \\[",
      showEngineeringDmg: "showDamageDoneWithEngineering \\[",
      showAverageOfHitsPerAoeCast: "showAverageOfHitsPerAoeCast \\[",
      totalAndInformationRowsDefaultTemplate: "totalAndInformationRowsDefaultTemplate \\[",
      showUsedTemporaryWeaponEnchant: "showUsedTemporaryWeaponEnchant \\[",

      damageTakenToTrack: "^damageTaken tracked \\[",
      damageTaken: "^damageTaken \\[",
      debuffsToTrack: "debuffs tracked \\[",
      debuffs: "^debuffs \\[",
      statsAndMiscToTrack: "statsAndMisc tracked \\[",
      statsAndMisc: "^statsAndMisc \\[",
      trinketsAndRacialsToTrack: "trinketsAndRacials tracked \\[",
      trinketsAndRacials: "^trinketsAndRacials \\[",
      engineeringToTrack: "engineering tracked \\[",
      engineering: "^engineering \\[",
      otherCastsToTrack: "otherCasts tracked \\[",
      otherCasts: "^otherCasts \\[",
      absorbsToTrack: "absorbs tracked \\[",
      absorbs: "^absorbs \\[",
      interrupts: "^interrupts \\[",
    };

    // Language-specific patterns
    const langPatterns = {};
    if (lang !== "EN") {
      Object.keys(configPatterns).forEach(key => {
        if (!key.startsWith('show') && !key.includes('WCL')) {
          const basePattern = configPatterns[key];
          if (key.includes('ToTrack')) {

            if (key === 'damageTakenToTrack') {
              langPatterns[key + 'Lang'] = "^damageTaken tracked " + lang;
            } else if (key === 'debuffsToTrack') {
              langPatterns[key + 'Lang'] = "^debuffs tracked " + lang;
            } else {
              langPatterns[key + 'Lang'] = key.replace('ToTrack', '') + " tracked " + lang;
            }
          } else {
            if (basePattern.startsWith('^')) {
              langPatterns[key + 'Lang'] = "^" + key + " " + lang;
            } else {
              langPatterns[key + 'Lang'] = key + " " + lang;
            }
          }
        }
      });
    }

    // Find all ranges at once
    const allPatterns = { ...configPatterns, ...langPatterns };
    const foundRanges = {};

    // Get all text finders in one batch
    Object.entries(allPatterns).forEach(([key, pattern]) => {
      foundRanges[key] = conf.createTextFinder(pattern).useRegularExpression(true).findNext();
    });

    // Extract config references
    const confRefs = {};
    Object.keys(configPatterns).forEach(key => {
      const capitalizedKey = 'conf' + key.charAt(0).toUpperCase() + key.slice(1);
      confRefs[capitalizedKey] = foundRanges[key];

      // Handle language variants
      if (lang !== "EN") {
        const langKey = key + 'Lang';
        if (foundRanges[langKey]) {
          confRefs[capitalizedKey + 'Lang'] = foundRanges[langKey];
        } else {
          confRefs[capitalizedKey + 'Lang'] = foundRanges[key]; // Fallback to EN
        }
      } else {
        confRefs[capitalizedKey + 'Lang'] = foundRanges[key];
      }
    });

    // Destructure for easier access
    const confSingleTargetCasts = confRefs.confSingleTargetCasts;
    const confSingleTargetCastsLang = confRefs.confSingleTargetCastsLang;
    const confAoeCasts = confRefs.confAoeCasts;
    const confAoeCastsLang = confRefs.confAoeCastsLang;
    const confClassCooldowns = confRefs.confClassCooldowns;
    const confClassCooldownsLang = confRefs.confClassCooldownsLang;
    const confSecondsActive = confRefs.confSecondsActive;

    const confShowDamageReflectRow = confRefs.confShowDamageReflectRow;
    const confShowFriendlyFireRow = confRefs.confShowFriendlyFireRow;
    const confShowDeathCountRow = confRefs.confShowDeathCountRow;
    const confShowInterruptedSpells = confRefs.confShowInterruptedSpells;
    const confShowInterruptedSpellsNamesRow = confRefs.confShowInterruptedSpellsNamesRow;
    const confShowConditionalFormattingDamageTaken = confRefs.confShowConditionalFormattingDamageTaken;
    const confShowOilOfImmolationDmg = confRefs.confShowOilOfImmolationDmg;
    const confShowEngineeringDmg = confRefs.confShowEngineeringDmg;
    const confShowAverageOfHitsPerAoeCast = confRefs.confShowAverageOfHitsPerAoeCast;
    const confTotalAndInformationRowsDefaultTemplate = confRefs.confTotalAndInformationRowsDefaultTemplate;
    const confShowUsedTemporaryWeaponEnchant = confRefs.confShowUsedTemporaryWeaponEnchant;

    const confDamageTakenToTrack = confRefs.confDamageTakenToTrack;
    const confDamageTakenToTrackLang = confRefs.confDamageTakenToTrackLang;
    const confDamageTaken = confRefs.confDamageTaken;
    const confDamageTakenLang = confRefs.confDamageTakenLang;
    const confDebuffsToTrack = confRefs.confDebuffsToTrack;
    const confDebuffsToTrackLang = confRefs.confDebuffsToTrackLang;
    const confDebuffs = confRefs.confDebuffs;
    const confDebuffsLang = confRefs.confDebuffsLang;
    const confStatsAndMiscToTrack = confRefs.confStatsAndMiscToTrack;
    const confStatsAndMiscToTrackLang = confRefs.confStatsAndMiscToTrackLang;
    const confStatsAndMisc = confRefs.confStatsAndMisc;
    const confStatsAndMiscLang = confRefs.confStatsAndMiscLang;
    const confTrinketsAndRacialsToTrack = confRefs.confTrinketsAndRacialsToTrack;
    const confTrinketsAndRacialsToTrackLang = confRefs.confTrinketsAndRacialsToTrackLang;
    const confTrinketsAndRacials = confRefs.confTrinketsAndRacials;
    const confTrinketsAndRacialsLang = confRefs.confTrinketsAndRacialsLang;
    const confEngineeringToTrack = confRefs.confEngineeringToTrack;
    const confEngineeringToTrackLang = confRefs.confEngineeringToTrackLang;
    const confEngineering = confRefs.confEngineering;
    const confEngineeringLang = confRefs.confEngineeringLang;
    const confOtherCastsToTrack = confRefs.confOtherCastsToTrack;
    const confOtherCastsToTrackLang = confRefs.confOtherCastsToTrackLang;
    const confOtherCasts = confRefs.confOtherCasts;
    const confOtherCastsLang = confRefs.confOtherCastsLang;
    const confAbsorbsToTrack = confRefs.confAbsorbsToTrack;
    const confAbsorbsToTrackLang = confRefs.confAbsorbsToTrackLang;
    const confAbsorbs = confRefs.confAbsorbs;
    const confAbsorbsLang = confRefs.confAbsorbsLang;
    const confInterrupts = confRefs.confInterrupts;
    const confInterruptsLang = confRefs.confInterruptsLang;

    // Debug logging for missing patterns
    const debugMissing = false; // Set to true to enable debug logging
    if (debugMissing) {
      Object.entries(confRefs).forEach(([name, value]) => {
        if (!value) {
          console.log(`Missing config: ${name}`);
        }
      });
    }

    //initialize variables
    var damageTakenMaxEntries = 15;
    var debuffsMaxEntries = 5;

    //initialize functionalities
    var showDamageReflectRow = confShowDamageReflectRow.getValue().split("[")[1].split("]")[0] == "true";
    var showFriendlyFireRow = confShowFriendlyFireRow.getValue().split("[")[1].split("]")[0] == "true";
    var showDeathCountRow = confShowDeathCountRow.getValue().split("[")[1].split("]")[0] == "true";
    var showInterruptedSpells = confShowInterruptedSpells.getValue().split("[")[1].split("]")[0] == "true";
    var showInterruptedSpellsNamesRow = confShowInterruptedSpellsNamesRow.getValue().split("[")[1].split("]")[0] == "true";
    var showConditionalFormattingDamageTaken = confShowConditionalFormattingDamageTaken.getValue().split("[")[1].split("]")[0] == "true";
    var showOilOfImmolationDmg = confShowOilOfImmolationDmg.getValue().split("[")[1].split("]")[0] == "true";
    var showEngineeringDmg = confShowEngineeringDmg.getValue().split("[")[1].split("]")[0] == "true";
    var showUsedTemporaryWeaponEnchant = confShowUsedTemporaryWeaponEnchant.getValue().split("[")[1].split("]")[0] == "true";
    var showWCLActivePercentage = conf.createTextFinder("showWCLActivePercentage \\[").useRegularExpression(true).findNext().getValue().split("[")[1].split("]")[0] == "true";
    var showAverageOfHitsPerAoeCast = conf.createTextFinder("showAverageOfHitsPerAoeCast \\[").useRegularExpression(true).findNext().getValue().split("[")[1].split("]")[0] == "true";

    //read tracked casts from conf spreadsheet
    var damageTakenToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confDamageTakenToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var damageTakenToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confDamageTakenToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var debuffsToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confDebuffsToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var debuffsToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confDebuffsToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var statsAndMiscToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confStatsAndMiscToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var statsAndMiscToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confStatsAndMiscToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var trinketsAndRacialsToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confTrinketsAndRacialsToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var trinketsAndRacialsToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confTrinketsAndRacialsToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var engineeringToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confEngineeringToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var engineeringToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confEngineeringToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var otherCastsToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confOtherCastsToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var otherCastsToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confOtherCastsToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var absorbsToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confAbsorbsToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
    var absorbsToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confAbsorbsToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);

    //define outputRanges
    var numberOfDamageTakenRows = Number(damageTakenMaxEntries) + 3;
    if (showDamageReflectRow)
      numberOfDamageTakenRows += 1;
    if (showFriendlyFireRow)
      numberOfDamageTakenRows += 2;
    if (showDeathCountRow)
      numberOfDamageTakenRows += 1;

    var secondsActiveLines = 7;
    if (showWCLActivePercentage)
      secondsActiveLines += 1;

    var aoeCastStartRow = confAoeCasts.getValue().split("[")[1].split("]")[0];
    var singleTargetCastsStartRow = confSingleTargetCasts.getValue().split("[")[1].split("]")[0]
    var secondsActiveStartRow = confSecondsActive.getValue().split("[")[1].split("]")[0];
    var classCooldownsStartRow = confClassCooldowns.getValue().split("[")[1].split("]")[0];
    var trinketsAndRacialsStartRow = confTrinketsAndRacials.getValue().split("[")[1].split("]")[0];
    var statsAndMiscStartRow = confStatsAndMisc.getValue().split("[")[1].split("]")[0];
    var otherCastsStartRow = confOtherCasts.getValue().split("[")[1].split("]")[0];
    var singleTargetCastsLines = aoeCastStartRow - singleTargetCastsStartRow + 1;
    var aoeCastsLines = showAverageOfHitsPerAoeCast ? secondsActiveStartRow - aoeCastStartRow + 2 : secondsActiveStartRow - aoeCastStartRow + 1;
    var classCooldownsLines = trinketsAndRacialsStartRow - classCooldownsStartRow + 1;

    if (darkMode) {
      sheet.getRange(4, 64).setFontColor("#d9d9d9").setValue(statsAndMiscStartRow);
      sheet.getRange(3, 64).setFontColor("#d9d9d9").setValue(otherCastsStartRow);
    }
    else {
      sheet.getRange(4, 64).setFontColor("white").setValue(statsAndMiscStartRow);
      sheet.getRange(3, 64).setFontColor("white").setValue(otherCastsStartRow);
    }

    if (confDamageTaken != null) var damageTaken = getOutputRange(sheet, confDamageTaken, numberOfDamageTakenRows, maxColumns + 1);
    if (confDebuffs != null) var debuffs = getOutputRange(sheet, confDebuffs, Number(debuffsMaxEntries) + 1, maxColumns + 1);
    var bonusEngi = 0;
    if (showEngineeringDmg)
      bonusEngi = 1;
    if (confEngineering != null) var engineering = getOutputRange(sheet, confEngineering, showOilOfImmolationDmg ? engineeringToTrack.length + 2 + bonusEngi : engineeringToTrack.length + 1 + bonusEngi, maxColumns + 1);
    if (confStatsAndMisc != null) var statsAndMisc = getOutputRange(sheet, confStatsAndMisc, statsAndMiscToTrack.length + 1, maxColumns + 1);
    if (confOtherCasts != null) var otherCasts = getOutputRange(sheet, confOtherCasts, showUsedTemporaryWeaponEnchant ? otherCastsToTrack.length + 2 : otherCastsToTrack.length + 1, maxColumns + 1);
    if (confAbsorbs != null) var absorbs = getOutputRange(sheet, confAbsorbs, absorbsToTrack.length + 2, maxColumns + 1);
    if (confInterrupts != null) var interrupts = getOutputRange(sheet, confInterrupts, showInterruptedSpellsNamesRow ? 3 : 2, maxColumns + 1);
    if (confTrinketsAndRacials != null) var trinketsAndRacials = getOutputRange(sheet, confTrinketsAndRacials, trinketsAndRacialsToTrack.length + 1, maxColumns + 1);
    if (confSingleTargetCasts != null) var singleTargetCasts = getOutputRange(sheet, confSingleTargetCasts, singleTargetCastsLines, maxColumns + 1);
    if (confAoeCasts != null) var aoeCasts = getOutputRange(sheet, confAoeCasts, aoeCastsLines, maxColumns + 1);
    if (confSecondsActive != null) var secondsActive = getOutputRange(sheet, confSecondsActive, secondsActiveLines, maxColumns + 1);
    if (confClassCooldowns != null) var classCooldowns = getOutputRange(sheet, confClassCooldowns, classCooldownsLines, maxColumns + 1);

    //define headers
    if (confSingleTargetCasts != null)
      var singleTargetCastsHeader = getHeaderFromConfig(confSingleTargetCastsLang);
    if (confAoeCastsLang != null)
      var aoeCastsHeader = getHeaderFromConfig(confAoeCastsLang);
    var damageTakenHeader = getHeaderFromConfig(confDamageTakenLang);
    var debuffsHeader = getHeaderFromConfig(confDebuffsLang);
    if (confClassCooldownsLang != null)
      var classCooldownsHeader = getHeaderFromConfig(confClassCooldownsLang);
    var statsAndMiscHeader = getHeaderFromConfig(confStatsAndMiscLang);
    var trinketsAndRacialsHeader = getHeaderFromConfig(confTrinketsAndRacialsLang);
    var engineeringHeader = getHeaderFromConfig(confEngineeringLang);
    var otherCastsHeader = getHeaderFromConfig(confOtherCastsLang);
    var absorbsHeader = getHeaderFromConfig(confAbsorbsLang);
    var interruptsHeader = getHeaderFromConfig(confInterruptsLang);

    //initialize output arrays with their respective header texts
    if (singleTargetCastsHeader != null) {
      var singleTargetCastsArr = [];
      fillUpMultiDimArrayWithEmptyValuesWithNumberOfRows(singleTargetCastsArr, maxColumns, singleTargetCastsLines)
      singleTargetCastsArr[0][0] = singleTargetCastsHeader;
      sheet.getRange(singleTargetCasts.getRow(), singleTargetCasts.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    }
    if (aoeCastsHeader != null) {
      var aoeCastsArr = [];
      fillUpMultiDimArrayWithEmptyValuesWithNumberOfRows(aoeCastsArr, maxColumns, aoeCastsLines);
      aoeCastsArr[0][0] = aoeCastsHeader;
      sheet.getRange(aoeCasts.getRow(), aoeCasts.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    }
    var damageTakenArr = [];
    addSingleEntryToMultiDimArray(damageTakenArr, damageTakenHeader);
    sheet.getRange(damageTaken.getRow(), damageTaken.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    var debuffsArr = [];
    addSingleEntryToMultiDimArray(debuffsArr, debuffsHeader);
    sheet.getRange(debuffs.getRow(), debuffs.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    if (classCooldownsHeader != null) {
      var classCooldownsArr = [];
      fillUpMultiDimArrayWithEmptyValuesWithNumberOfRows(classCooldownsArr, maxColumns, classCooldownsLines);
      classCooldownsArr[0][0] = classCooldownsHeader;
      sheet.getRange(classCooldowns.getRow(), classCooldowns.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    }
    var statsAndMiscArr = [];
    addSingleEntryToMultiDimArray(statsAndMiscArr, statsAndMiscHeader);
    sheet.getRange(statsAndMisc.getRow(), statsAndMisc.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    var trinketsAndRacialsArr = [];
    addSingleEntryToMultiDimArray(trinketsAndRacialsArr, trinketsAndRacialsHeader);
    sheet.getRange(trinketsAndRacials.getRow(), trinketsAndRacials.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    var engineeringArr = [];
    addSingleEntryToMultiDimArray(engineeringArr, engineeringHeader);
    sheet.getRange(engineering.getRow(), engineering.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    var otherCastsArr = [];
    addSingleEntryToMultiDimArray(otherCastsArr, otherCastsHeader);
    sheet.getRange(otherCasts.getRow(), otherCasts.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    var absorbsArr = [];
    addSingleEntryToMultiDimArray(absorbsArr, absorbsHeader);
    sheet.getRange(absorbs.getRow(), absorbs.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    var interruptsArr = [];
    addSingleEntryToMultiDimArray(interruptsArr, interruptsHeader);
    sheet.getRange(interrupts.getRow(), interrupts.getColumn()).setFontWeight("bold").setHorizontalAlignment("right");
    var topDamageTakenDoneArr = [];
    var topDamageTakenDoneArrOriginal = [];
    addSingleEntryToMultiDimArray(topDamageTakenDoneArr, damageTakenHeader);
    addSingleEntryToMultiDimArray(topDamageTakenDoneArrOriginal, damageTakenHeader);
    var debuffsDoneArr = [];
    var debuffsDoneArrOriginal = [];
    addSingleEntryToMultiDimArray(debuffsDoneArr, debuffsHeader);
    addSingleEntryToMultiDimArray(debuffsDoneArrOriginal, debuffsHeader);
    if (confSecondsActive != null) {
      var secondsActiveArr = [];
      addSingleEntryToMultiDimArray(secondsActiveArr, "");
    }

    //fill in the information section
    var totalTimeElapsedBosses = 0;
    var totalTimeElapsedRaw = 0;
    allFightsData.fights.forEach(function (fight, fightcount) {
      if (!onlyTrash && onlyFightNr != null && onlyFightNr.toString().length > 0 && onlyFightNr.toString() == fight.id.toString()) {
        totalTimeElapsedBosses = Number(fight.end_time) - Number(fight.start_time);
        totalTimeElapsedRaw = Number(fight.end_time) - Number(fight.start_time);
      } else if (fight.boss > 0 && !fight.boss.toString().endsWith("724") && !onlyTrash && (!noWipes || (noWipes && fight.kill != null && fight.kill.toString() == "true"))) {
        totalTimeElapsedBosses += Number(fight.end_time) - Number(fight.start_time);
      }
      if ((onlyFightNr == null || onlyFightNr.toString().length == 0) && ((onlyBosses && fight.boss > 0) || !onlyBosses) && (!noWipes || (noWipes && fight.kill != null && fight.kill.toString() == "true"))) {
        if ((!onlyTrash || (onlyTrash && fight.boss.toString() == "0")) && !fight.boss.toString().endsWith("724")) {
          totalTimeElapsedRaw += Number(fight.end_time) - Number(fight.start_time);
        }
      }
    })

    var hideArr = [];
    for (v = 8, w = sheet.getMaxRows(); v <= w; v++) {
      hideArr.push("=IF(ISERROR(MATCH(IF(ISERROR(MATCH(" + sheet.getRange(v, 2).getA1Notation() + ",\" (\",0)),INDEX(SPLIT(INDEX(SPLIT(" + sheet.getRange(v, 2).getA1Notation() + ",\" (\", FALSE, TRUE), 0, 1),\" [\", FALSE, TRUE), 0, 1), " + sheet.getRange(v, 2).getA1Notation() + "),settings!$B$2:$B$10000,0)),\"yes\", \"no\")")
    }
    sheet.getRange(8, 1, sheet.getMaxRows() - 7, 1).setValues(convertMultiRowSingleColumnArraytoMultidimensionalArray(hideArr));

    var damageTakenTop = JSON.parse(UrlFetchApp.fetch(urlDamageTakenTop));
    var damageTakenTopByTotalDesc = [];
    var damageTakenTopByTotalDescOriginal = [];
    damageTakenToTrack.forEach(function (ability, abilityCount) {
      var total = 0;
      var name = "";
      var nameOriginal = "";
      var sourcesString = "";
      var triggerString = "";
      if (ability.indexOf("[") > -1) {
        var abilityIds = ability.split("[")[1].split("]")[0];
        if (ability.indexOf("Cleave") > -1)
          abilityIds += ",15754";
        abilityIds.split(",").forEach(function (spellId, spellIdCount) {
          damageTakenTop.entries.forEach(function (abilityFromLogs, abilityFromLogsCount) {
            if (abilityFromLogs.guid != null && abilityFromLogs.guid.toString().length > 0 && spellId == abilityFromLogs.guid.toString()) {
              total += abilityFromLogs.total;
              name = damageTakenToTrackLang[abilityCount];
              nameOriginal = damageTakenToTrack[abilityCount];
              if (abilityFromLogs.sources != null && abilityFromLogs.sources.length > 0) {
                abilityFromLogs.sources.forEach(function (abilitySource, abilitySourceCount) {
                  var abilitySourceName = abilitySource.name;
                  if (abilitySource.name.endsWith(" "))
                    abilitySourceName = abilitySource.name.substring(0, abilitySource.name.length - 1);
                  var abilitySourceNameCorrected = abilitySourceName.replace("[", "").replace("] ", "").replace("]", "").replace("UNUSED", "");
                  if (!(sourcesString.indexOf(abilitySourceNameCorrected) > -1)) {
                    sourcesString += abilitySourceNameCorrected + "/";
                  }
                })
              }
            }
          })
        })
        if (sourcesString.length > 0) {
          if ((triggerString.length == 0 && sourcesString.length > (58 - name.split(" [")[0].length)) || (triggerString.length > 0 && sourcesString.length > (43 - name.split(" [")[0].length - triggerString.length))) {
            sourcesString = shortenSources(sourcesString);
          }
          if (triggerString.length > 0) {
            if (!damageTakenTopByTotalDesc.includes(name.split(" [")[0] + " [" + nameOriginal.split("[")[1].split("]")[0] + "] (" + sourcesString.substr(0, sourcesString.length - 1) + ", " + getStringForLang("triggeredBy", langKeys, langTrans, "", "", "", "") + triggerString + ")" + ": " + total)) {
              damageTakenTopByTotalDesc.push(name.split(" [")[0] + " [" + nameOriginal.split("[")[1].split("]")[0] + "] (" + sourcesString.substr(0, sourcesString.length - 1) + ", " + getStringForLang("triggeredBy", langKeys, langTrans, "", "", "", "") + triggerString + ")" + ": " + total);
              damageTakenTopByTotalDescOriginal.push(nameOriginal + " (" + sourcesString.substr(0, sourcesString.length - 1) + ", " + getStringForLang("triggeredBy", langKeys, langTrans, "", "", "", "") + triggerString + ")" + ": " + total);
            }
          } else {
            if (!damageTakenTopByTotalDesc.includes(name.split("[")[0] + " [" + nameOriginal.split("[")[1].split("]")[0] + "] (" + sourcesString.substr(0, sourcesString.length - 1) + ")" + ": " + total)) {
              damageTakenTopByTotalDesc.push(name.split(" [")[0] + " [" + nameOriginal.split("[")[1].split("]")[0] + "] (" + sourcesString.substr(0, sourcesString.length - 1) + ")" + ": " + total);
              damageTakenTopByTotalDescOriginal.push(nameOriginal + " (" + sourcesString.substr(0, sourcesString.length - 1) + ")" + ": " + total);
            }
          }
        }
      }
    })
    damageTakenTopByTotalDesc = damageTakenTopByTotalDesc.sort(function (a, b) {
      a = a.split(": ")[1].toString().match(/^\d+$/) ? +a.split(": ")[1] : a.split(": ")[1];
      b = b.split(": ")[1].toString().match(/^\d+$/) ? +b.split(": ")[1] : b.split(": ")[1];
      return ((a < b) ? 1 : ((a > b) ? -1 : 0));
    })

    //fill in names of top damageTaken
    damageTakenTopByTotalDesc.forEach(function (abilityByTotalDesc, abilityByTotalDescCount) {
      if (!(topDamageTakenDoneArr.some(e => new RegExp(abilityByTotalDesc.split(": ")[0] + ".*", "g").test(e[0])))) {
        if (topDamageTakenDoneArr.length <= damageTakenMaxEntries) {
          addSingleEntryToMultiDimArray(topDamageTakenDoneArr, abilityByTotalDesc.split(": ")[0]);
        }
      }
      if (!(topDamageTakenDoneArrOriginal.some(e => new RegExp(damageTakenTopByTotalDescOriginal[abilityByTotalDescCount].split(": ")[0] + ".*", "g").test(e[0])))) {
        if (topDamageTakenDoneArrOriginal.length <= damageTakenMaxEntries) {
          addSingleEntryToMultiDimArray(topDamageTakenDoneArrOriginal, damageTakenTopByTotalDescOriginal[abilityByTotalDescCount].split(": ")[0]);
        }
      }
    })
    for (var i = 1, j = topDamageTakenDoneArr.length; i <= damageTakenMaxEntries; i++) {
      if (i < j) {
        var rangeRow = sheet.getRange(damageTaken.getRow() + i, damageTaken.getColumn() + 1, 1, maxColumns);

        var rng = conf.createTextFinder(topDamageTakenDoneArrOriginal[i][0].split(' (')[0]).findNext();
        if (rng != null)
          copyRangeStyle(rng, rangeRow, null, "center", null);
      }
      else
        addSingleEntryToMultiDimArray(topDamageTakenDoneArr, "");
    }
    if (showDamageReflectRow) {
      addSingleEntryToMultiDimArray(topDamageTakenDoneArr, getStringForLang("damageReflected", langKeys, langTrans, "", "", "", ""));
      copyRangeStyle(confShowDamageReflectRow, sheet.getRange(damageTaken.getRow() + topDamageTakenDoneArr.length - 1, damageTaken.getColumn() + 1, 1, maxColumns), null, "center", null);
    }
    if (showFriendlyFireRow) {
      addSingleEntryToMultiDimArray(topDamageTakenDoneArr, getStringForLang("damageHostile", langKeys, langTrans, "", "", "", ""));
      copyRangeStyle(confShowFriendlyFireRow, sheet.getRange(damageTaken.getRow() + topDamageTakenDoneArr.length - 1, damageTaken.getColumn() + 1, 1, maxColumns), null, "center", null);
      addSingleEntryToMultiDimArray(topDamageTakenDoneArr, getStringForLang("friendlyFire", langKeys, langTrans, "", "", "", ""));
      copyRangeStyle(confShowFriendlyFireRow, sheet.getRange(damageTaken.getRow() + topDamageTakenDoneArr.length - 1, damageTaken.getColumn() + 1, 1, maxColumns), null, "center", null);
    }
    if (showDeathCountRow) {
      if (onlyBosses || onlyTrash || (onlyFightNr != null && onlyFightNr.toString().length > 0))
        addSingleEntryToMultiDimArray(topDamageTakenDoneArr, getStringForLang("totalDeaths", langKeys, langTrans, "", "", "", ""));
      else
        addSingleEntryToMultiDimArray(topDamageTakenDoneArr, getStringForLang("totalDeaths", langKeys, langTrans, "", "", "", "") + " (" + getStringForLang("justOnTrash", langKeys, langTrans, "", "", "", "") + ")");
      copyRangeStyle(confShowDeathCountRow, sheet.getRange(damageTaken.getRow() + topDamageTakenDoneArr.length - 1, damageTaken.getColumn() + 1, 1, maxColumns), null, "center", null);
    }
    addSingleEntryToMultiDimArray(topDamageTakenDoneArr, getStringForLang("avoidableDamageTaken", langKeys, langTrans, "", "", "", ""));
    copyRangeStyle(confTotalAndInformationRowsDefaultTemplate, sheet.getRange(damageTaken.getRow() + topDamageTakenDoneArr.length - 1, damageTaken.getColumn() + 1, 1, maxColumns), null, "center", null);
    if (showConditionalFormattingDamageTaken) {
      var rangeCellStart = sheet.getRange(damageTaken.getRow() + 1, damageTaken.getColumn() + 1);
      var rangeCellEnd = sheet.getRange(damageTaken.getRow() + 1, damageTaken.getColumn() + maxColumns);
      var ruleRange = sheet.getRange(damageTaken.getRow() + 1, damageTaken.getColumn() + 1, topDamageTakenDoneArr.length - 1, maxColumns);
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=(' + rangeCellStart.getA1Notation() + '/AVERAGE($' + rangeCellStart.getA1Notation() + ':$' + rangeCellEnd.getA1Notation() + '))>1.25')
        .setBackground("#f9cb9c")
        .setRanges([ruleRange])
        .build();

      var rule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=(' + rangeCellStart.getA1Notation() + '/AVERAGE($' + rangeCellStart.getA1Notation() + ':$' + rangeCellEnd.getA1Notation() + '))>1.75')
        .setBackground("#ea9999")
        .setRanges([ruleRange])
        .build();

      var rules = sheet.getConditionalFormatRules();
      rules.push(rule2);
      rules.push(rule);
      sheet.setConditionalFormatRules(rules);
    }
    addSingleEntryToMultiDimArray(topDamageTakenDoneArr, "");
    shiftRangeByRows(sheet, addRowsToRange(sheet, addColumnsToRange(sheet, damageTaken, 1 - damageTaken.getNumColumns()).setValues(topDamageTakenDoneArr), -1), 1).setFontSize(8).setHorizontalAlignment("right");
    sheet.getRange(damageTaken.getRow() + topDamageTakenDoneArr.length - 1, damageTaken.getColumn() + 1, 1, maxColumns).setHorizontalAlignment("center").setNumberFormat("0%");

    var debuffsTakenTop = JSON.parse(UrlFetchApp.fetch(urlDebuffsTop));
    var debuffsTopByTotalDesc = [];
    var debuffsTopByTotalDescOriginal = [];
    debuffsToTrack.forEach(function (ability, abilityCount) {
      var total = 0;
      var name = "";
      var nameOriginal = "";
      var sourcesString = "";
      if (ability.indexOf("[") > -1) {
        ability.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
          debuffsTakenTop.auras.forEach(function (abilityFromLogs, abilityFromLogsCount) {
            if (abilityFromLogs.guid != null && abilityFromLogs.guid.toString().length > 0 && spellId == abilityFromLogs.guid.toString()) {
              total += abilityFromLogs.totalUses;
              name = debuffsToTrackLang[abilityCount];
              nameOriginal = debuffsToTrack[abilityCount];
              var debuffInfoData = JSON.parse(UrlFetchApp.fetch(urlDebuffInfo + spellId));
              if (debuffInfoData.auras != null && debuffInfoData.auras.length > 0) {
                debuffInfoData.auras.forEach(function (abilitySource, abilitySourceCount) {
                  sourcesString += abilitySource.name + "/";
                })
              }
            }
          })
        })
        if (sourcesString.length > 0) {
          if (sourcesString.length > (58 - name.split(" [")[0].length)) {
            sourcesString = shortenSources(sourcesString);
          }
          if (!debuffsTopByTotalDesc.includes(name.split(" [")[0] + " [" + nameOriginal.split("[")[1].split("]")[0] + "] (" + sourcesString.substr(0, sourcesString.length - 1) + ")" + ": " + total)) {
            debuffsTopByTotalDesc.push(name.split(" [")[0] + " [" + nameOriginal.split("[")[1].split("]")[0] + "] (" + sourcesString.substr(0, sourcesString.length - 1) + ")" + ": " + total);
            debuffsTopByTotalDescOriginal.push(nameOriginal + " (" + sourcesString.substr(0, sourcesString.length - 1) + ")" + ": " + total);
          }
        }
      }
    })
    debuffsTopByTotalDesc = debuffsTopByTotalDesc.sort(function (a, b) {
      a = a.split(": ")[1].toString().match(/^\d+$/) ? +a.split(": ")[1] : a.split(": ")[1];
      b = b.split(": ")[1].toString().match(/^\d+$/) ? +b.split(": ")[1] : b.split(": ")[1];
      return ((a < b) ? 1 : ((a > b) ? -1 : 0));
    })

    //fill in names of debuffs applied
    debuffsTopByTotalDesc.forEach(function (debuffByTotalDesc, debuffByTotalDescCount) {
      if (!(debuffsDoneArr.some(e => new RegExp(debuffByTotalDesc.split(": ")[0] + ".*", "g").test(e[0])))) {
        if (debuffsDoneArr.length <= debuffsMaxEntries) {
          addSingleEntryToMultiDimArray(debuffsDoneArr, debuffByTotalDesc.split(": ")[0]);
        }
      }
      if (!(debuffsDoneArrOriginal.some(e => new RegExp(debuffsTopByTotalDescOriginal[debuffByTotalDescCount].split(": ")[0] + ".*", "g").test(e[0])))) {
        if (debuffsDoneArrOriginal.length <= debuffsMaxEntries) {
          addSingleEntryToMultiDimArray(debuffsDoneArrOriginal, debuffsTopByTotalDescOriginal[debuffByTotalDescCount].split(": ")[0]);
        }
      }
    })
    for (var i = 1, j = debuffsDoneArr.length; i <= debuffsMaxEntries; i++) {
      if (i < j) {
        var rangeRow = sheet.getRange(debuffs.getRow() + i, debuffs.getColumn() + 1, 1, maxColumns);

        var rng = conf.createTextFinder(debuffsDoneArrOriginal[i][0].split(' (')[0]).findNext();
        if (rng != null)
          copyRangeStyle(rng, rangeRow, null, "center", null);
      }
      else
        addSingleEntryToMultiDimArray(debuffsDoneArr, "");
    }
    shiftRangeByRows(sheet, addRowsToRange(sheet, addColumnsToRange(sheet, debuffs, 1 - debuffs.getNumColumns()).setValues(debuffsDoneArr), -1), 1).setFontSize(8).setHorizontalAlignment("right");

    bossSummaryDataAll = [];
    bossDamageDataAll = [];
    allFightsData.fights.forEach(function (fight, fightCount) {
      if ((fight.start_time != fight.end_time) && fight.boss > 0 && ((onlyFightNr != null && onlyFightNr.toString() == fight.id.toString()) || onlyFightNr.toString().length == 0)) {
        bossSummaryDataAll.push(JSON.parse(UrlFetchApp.fetch(urlSummary.replace(startEndString, "&start=" + fight.start_time + "&end=" + fight.end_time).replace("&encounter=0", ""))));
        bossDamageDataAll.push(JSON.parse(UrlFetchApp.fetch(urlDamageDone.replace("&sourceid=", "").replace(startEndString, "&start=" + fight.start_time + "&end=" + fight.end_time) + "&abilityid=27187")));
        if (fightCount % 10 == 0)
          Utilities.sleep(500);
      }
    })

    var spellHasteInfoIdsRaw = confSpellHasteConfig.getRange(1, 1, 1500, 1).getValues();
    var spellHasteInfoIds = spellHasteInfoIdsRaw.reduce(function (ar, e) {
      if (e[0]) ar.push(e[0])
      return ar;
    }, []);

    var spellHasteInfoValueRaw = confSpellHasteConfig.getRange(1, 2, 1500, 1).getValues();
    var spellHasteInfoValues = spellHasteInfoValueRaw.reduce(function (ar, e) {
      if (e[0]) ar.push(e[0])
      return ar;
    }, []);

    var playerDoneCount = 0;
    var previousClass = "";
    var classDoneCount = 0;
    var singleTargetCastsToTrack = null;
    var singleTargetCastsToTrackLang = null;
    var aoeCastsToTrack = null;
    var aoeCastsToTrackLang = null;
    var classCooldownsToTrack = null;
    var classCooldownsToTrackLang = null;
    var rolesAndNames = [];
    rolesAndNames[0] = [];
    rolesAndNames[1] = [];
    rolesAndNames[2] = [];
    allPlayersByNameAsc.forEach(function (playerByNameAsc, playerCount) {
      var characterNamesArr = characterNames.replace(/\s/g, "").split(",");
      if ((playerByNameAsc.type == "Druid" || playerByNameAsc.type == "Hunter" || playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Priest" || playerByNameAsc.type == "Paladin" || playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Shaman" || playerByNameAsc.type == "Warlock" || playerByNameAsc.type == "Warrior") && (playerByNameAsc.total > 20 || fightName != "") && (characterNames.length < 1 || (characterNames.length > 0 && checkIfArrayContainsEntry(characterNamesArr, playerByNameAsc.name)))) {
        if (previousClass != playerByNameAsc.type) {
          var confSingleTargetCastsToTrack = conf.createTextFinder("^singleTargetCasts tracked " + playerByNameAsc.type + " \\[").useRegularExpression(true).findNext();
          var confSingleTargetCastsToTrackLang; if (lang == "EN") confSingleTargetCastsToTrackLang = confSingleTargetCastsToTrack; else confSingleTargetCastsToTrackLang = conf.createTextFinder("^singleTargetCasts tracked " + playerByNameAsc.type + " " + lang).useRegularExpression(true).findNext();
          var confClassCooldownsToTrack = conf.createTextFinder("^classCooldowns tracked " + playerByNameAsc.type + " \\[").useRegularExpression(true).findNext();
          var confClassCooldownsToTrackLang; if (lang == "EN") confClassCooldownsToTrackLang = confClassCooldownsToTrack; else confClassCooldownsToTrackLang = conf.createTextFinder("^classCooldowns tracked " + playerByNameAsc.type + " " + lang).useRegularExpression(true).findNext();
          var confAoeCastsToTrack = conf.createTextFinder("aoeCasts tracked " + playerByNameAsc.type + " \\[").useRegularExpression(true).findNext();
          var confAoeCastsToTrackLang; if (lang == "EN") confAoeCastsToTrackLang = confAoeCastsToTrack; else confAoeCastsToTrackLang = conf.createTextFinder("^aoeCasts tracked " + playerByNameAsc.type + " " + lang).useRegularExpression(true).findNext();
          if (confSingleTargetCastsToTrack != null) {
            singleTargetCastsToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confSingleTargetCastsToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
            singleTargetCastsToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confSingleTargetCastsToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
          }
          if (confAoeCastsToTrack != null) {
            aoeCastsToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confAoeCastsToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
            aoeCastsToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confAoeCastsToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
          }
          if (confClassCooldownsToTrack != null) {
            classCooldownsToTrack = addRowsToRange(conf, shiftRangeByRows(conf, confClassCooldownsToTrack, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
            classCooldownsToTrackLang = addRowsToRange(conf, shiftRangeByRows(conf, confClassCooldownsToTrackLang, 1), 200).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
          }

          var playersInThisClass = 0;
          allPlayersByNameAsc.forEach(function (playerByNameAscSearch, playerCountSearch) {
            if (playerByNameAscSearch.type == playerByNameAsc.type && (playerByNameAscSearch.total > 20 || fightName != "") && (characterNames.length < 1 || (characterNames.length > 0 && checkIfArrayContainsEntry(characterNamesArr, playerByNameAscSearch.name)))) {
              playersInThisClass++;
            }
          })

          classDoneCount++;
        }
        Utilities.sleep(150);
        //load player-related queries into datastructures
        var playerData = JSON.parse(UrlFetchApp.fetch(urlPlayers + playerByNameAsc.id));
        var playerDataTrash = JSON.parse(UrlFetchApp.fetch(urlPlayersOnTrash + playerByNameAsc.id));
        var damageDoneData = JSON.parse(UrlFetchApp.fetch(urlDamageDone + playerByNameAsc.id));
        var buffsDataTrash = JSON.parse(UrlFetchApp.fetch(urlBuffsOnTrash + playerByNameAsc.id));
        var buffsData = JSON.parse(UrlFetchApp.fetch(urlBuffsTotal + playerByNameAsc.id));
        var damageTakenTotalData = JSON.parse(UrlFetchApp.fetch(urlDamageTakenTotal + playerByNameAsc.id));
        var debuffsAppliedData = JSON.parse(UrlFetchApp.fetch(urlDebuffsApplied + playerByNameAsc.id));
        var debuffsAppliedDataBosses = JSON.parse(UrlFetchApp.fetch(urlDebuffsAppliedBosses + playerByNameAsc.id));
        if (playerByNameAsc.type == "Paladin")
          var debuffsAppliedDataBossesJudgement = JSON.parse(UrlFetchApp.fetch(urlDebuffsAppliedBossesJudgement + playerByNameAsc.id));
        var debuffsData = JSON.parse(UrlFetchApp.fetch(urlDebuffs + playerByNameAsc.id));
        var healingData = JSON.parse(UrlFetchApp.fetch(urlHealing + playerByNameAsc.id));
        var healingDataTarget = JSON.parse(UrlFetchApp.fetch(urlHealingTarget + playerByNameAsc.id + "&by=ability"));
        if (playerByNameAsc.type == "Priest") {
          var VTManaGainData = JSON.parse(UrlFetchApp.fetch(urlVTManaGain + playerByNameAsc.id));
          var shadowDamageDoneData = JSON.parse(UrlFetchApp.fetch(urlShadowDamageDone + playerByNameAsc.id));
        } else if (playerByNameAsc.type == "Paladin" && !onlyTrash) {
          var TwistsDoneOnBosses = JSON.parse(UrlFetchApp.fetch(urlTwistsDoneOnBosses + playerByNameAsc.id));
          var WindfuryAttacksOnTwistsDoneOnBosses = JSON.parse(UrlFetchApp.fetch(urlWindfuryAttacksOnTwistsDoneOnBosses + playerByNameAsc.id));
          var WindfuryAttacksOnTwistsDoneOnBossesRank1 = JSON.parse(UrlFetchApp.fetch(urlWindfuryAttacksOnTwistsDoneOnBosses.replace("25584", "8516") + playerByNameAsc.id));
          var WindfuryAttacksOnBosses = JSON.parse(UrlFetchApp.fetch(urlWindfuryAttacksOnBosses + playerByNameAsc.id));
          var WindfuryAttacksOnBossesRank1 = JSON.parse(UrlFetchApp.fetch(urlWindfuryAttacksOnBosses.replace("25584", "8516") + playerByNameAsc.id));
          var DamageDoneOnBosses = JSON.parse(UrlFetchApp.fetch(urlDamageDoneOnBosses + playerByNameAsc.id));
          Utilities.sleep(250);
        }

        var hostilePlayersTotal = 0;
        hostilePlayersData.entries.forEach(function (hostilePlayersDataEntry, hostilePlayersDataEntryCount) {
          if (hostilePlayersDataEntry.id == playerByNameAsc.id)
            hostilePlayersTotal += hostilePlayersDataEntry.total;
        })

        var friendlyFireTotal = 0;
        var urlFriendlyFire = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndStringNoFilter + "&filter=NOT%20IN%20RANGE%20FROM%20type%20%3D%20%22applydebuff%22%20AND%20ability.id%20%3D%20%2229546%22%20AND%20target.name%3D%22Qlap%22%20TO%20type%20%3D%20%22removedebuff%22%20and%20ability.id%3D%2229546%22%20AND%20target.name%3D%22Qlap%22%20END%20AND%20NOT%20IN%20RANGE%20FROM%20type%20%3D%20%22applydebuff%22%20AND%20ability.id%20%3D%20%2245717%22%20AND%20target.name%3D%22Qlap%22%20TO%20type%20%3D%20%22removedebuff%22%20and%20ability.id%3D%2245717%22%20AND%20target.name%3D%22Qlap%22%20END%20AND%20NOT%20IN%20RANGE%20FROM%20type%20%3D%20%22applydebuff%22%20AND%20ability.id%20%3D%20%2237122%22%20AND%20target.name%3D%22Qlap%22%20TO%20type%20%3D%20%22removedebuff%22%20and%20ability.id%3D%2237122%22%20AND%20target.name%3D%22Qlap%22%20END%20AND%20NOT%20IN%20RANGE%20FROM%20type%20%3D%20%22applydebuff%22%20AND%20ability.id%20%3D%20%2237135%22%20AND%20target.name%3D%22Qlap%22%20TO%20type%20%3D%20%22removedebuff%22%20and%20ability.id%3D%2237135%22%20AND%20target.name%3D%22Qlap%22%20END%20AND%20NOT%20IN%20RANGE%20FROM%20type%20%3D%20%22applydebuff%22%20AND%20ability.id%20%3D%20%2241345%22%20AND%20target.name%3D%22Qlap%22%20TO%20type%20%3D%20%22removedebuff%22%20and%20ability.id%3D%2241345%22%20AND%20target.name%3D%22Qlap%22%20END%20AND%20NOT%20IN%20RANGE%20FROM%20type%20%3D%20%22applydebuff%22%20AND%20ability.id%20%3D%20%2243361%22%20AND%20target.name%3D%22Qlap%22%20TO%20type%20%3D%20%22removedebuff%22%20and%20ability.id%3D%2243361%22%20AND%20target.name%3D%22Qlap%22%20END%20AND%20encounterid%20%21%3D%20724%20AND%20ability.id%20%21%3D%2046768%20&options=4135&by=target&targetid=";
        var urlFriendlyFireReplaced = replaceAll(urlFriendlyFire, "Qlap", playerByNameAsc.name) + playerByNameAsc.id;
        var friendlyFireData = JSON.parse(UrlFetchApp.fetch(urlFriendlyFireReplaced));
        if (friendlyFireData != null && friendlyFireData.entries != null && friendlyFireData.entries.length > 0 && friendlyFireData.entries[0].total != null && friendlyFireData.entries[0].total > 0) {
          friendlyFireTotal = friendlyFireData.entries[0].total;
        }
        var urlFriendlyFireLinkPlayer = urlFriendlyFireReplaced.replace(baseUrl + "report/tables/damage-taken/", baseUrlFrontEnd).replace(logId, logId + "#type=damage-taken").replace(apiKeyString, "").replace(startEndStringNoFilter, "").replace("&targetid=" + playerByNameAsc.id, "").replace("&filter=", "&pins=2%24Off%24%23244F4B%24expression%24") + "&translate=true&boss=" + bossString + "&difficulty=0&view=events&target=" + playerByNameAsc.id;

        var urlHostilePlayersLinkPlayer = urlHostilePlayersLink + "&source=" + playerByNameAsc.id;
        var urlDamageReflectedLinkPlayer = urlDamageReflectedLink + "&target=" + playerByNameAsc.id;

        //get totalUses of band-driven casts
        var gotAtLeastOnePI = false;
        var PIonTrash = 0;
        var PIOnBosses = 0;
        var PITotalUptime = 0;
        var PIonTrashSelf = 0;
        var PIOnBossesSelf = 0;
        var gotAtLeastOneAPComb = false;
        var APCombonTrash = 0;
        var APCombOnBosses = 0;
        var APCombTotalUptime = 0;
        var gotAtLeastOneEvocation = false;
        var EvocationOnTrash = 0;
        var EvocationOnBosses = 0;
        var EvocationTotalUptime = 0;
        var gotAtLeastOneEvasion = false;
        var EvasionOnTrash = 0;
        var EvasionOnBosses = 0;
        var EvasionTotalUptime = 0;
        var gotAtLeastOneAdrenalineRush = false;
        var AdrenalineRushOnTrash = 0;
        var AdrenalineRushOnBosses = 0;
        var AdrenalineRushTotalUptime = 0;
        var gotAtLeastOneBerserkerRage = false;
        var BerserkerRageOnTrash = 0;
        var BerserkerRageOnBosses = 0;
        var BerserkerRageTotalUptime = 0;
        var gotAtLeastOneBloodrage = false;
        var BloodrageOnTrash = 0;
        var BloodrageOnBosses = 0;
        var BloodrageTotalUptime = 0;
        var gotAtLeastOneChallengingShout = false;
        var ChallengingShoutOnTrash = 0;
        var ChallengingShoutOnBosses = 0;
        var ChallengingShoutTotalUptime = 0;
        var gotAtLeastOneDeathWish = false;
        var DeathWishOnTrash = 0;
        var DeathWishOnBosses = 0;
        var DeathWishTotalUptime = 0;
        var gotAtLeastOneLastStand = false;
        var LastStandOnTrash = 0;
        var LastStandOnBosses = 0;
        var LastStandTotalUptime = 0;
        var gotAtLeastOneRecklessness = false;
        var RecklessnessOnTrash = 0;
        var RecklessnessOnBosses = 0;
        var RecklessnessTotalUptime = 0;
        var gotAtLeastOneRetaliation = false;
        var RetaliationOnTrash = 0;
        var RetaliationOnBosses = 0;
        var RetaliationTotalUptime = 0;
        var gotAtLeastOneShieldWall = false;
        var ShieldWallOnTrash = 0;
        var ShieldWallOnBosses = 0;
        var ShieldWallTotalUptime = 0;
        var gotAtLeastOneNaturesSwiftness = false;
        var NaturesSwiftnessOnTrash = 0;
        var NaturesSwiftnessOnBosses = 0;
        var NaturesSwiftnessTotalUptime = 0;
        var gotAtLeastOneElementalMastery = false;
        var ElementalMasteryOnTrash = 0;
        var ElementalMasteryOnBosses = 0;
        var ElementalMasteryTotalUptime = 0;
        var gotAtLeastOneRebirth = false;
        var RebirthOnTrash = 0;
        var RebirthOnBosses = 0;
        var RebirthTotalUptime = 0;
        var gotAtLeastOneChallengingRoar = false;
        var ChallengingRoarOnTrash = 0;
        var ChallengingRoarOnBosses = 0;
        var ChallengingRoarTotalUptime = 0;
        var gotAtLeastOneDash = false;
        var DashOnTrash = 0;
        var DashOnBosses = 0;
        var DashTotalUptime = 0;
        var gotAtLeastOneFrenziedRegeneration = false;
        var FrenziedRegenerationOnTrash = 0;
        var FrenziedRegenerationOnBosses = 0;
        var FrenziedRegenerationTotalUptime = 0;
        var gotAtLeastOneInnervate = false;
        var InnervateOnTrash = 0;
        var InnervateOnBosses = 0;
        var InnervateTotalUptime = 0;
        var InnervateOnTrashSelf = 0;
        var InnervateOnBossesSelf = 0;
        var gotAtLeastOneTranquility = false;
        var TranquilityOnTrash = 0;
        var TranquilityOnBosses = 0;
        var TranquilityTotalUptime = 0;
        var gotAtLeastOneRapidFire = false;
        var RapidFireOnTrash = 0;
        var RapidFireOnBosses = 0;
        var RapidFireTotalUptime = 0;
        var gotAtLeastOneReadiness = false;
        var ReadinessOnTrash = 0;
        var ReadinessOnBosses = 0;
        var ReadinessTotalUptime = 0;
        var gotAtLeastOneDeterrence = false;
        var DeterrenceOnTrash = 0;
        var DeterrenceOnBosses = 0;
        var DeterrenceTotalUptime = 0;
        var gotAtLeastOneBestialWrath = false;
        var BestialWrathOnTrash = 0;
        var BestialWrathOnBosses = 0;
        var BestialWrathTotalUptime = 0;
        var gotAtLeastOneInnerFocus = false;
        var InnerFocusOnTrash = 0;
        var InnerFocusOnBosses = 0;
        var InnerFocusTotalUptime = 0;
        var gotAtLeastOneShadowfiend = false;
        var ShadowfiendOnTrash = 0;
        var ShadowfiendOnBosses = 0;
        var ShadowfiendTotalUptime = 0;
        var gotAtLeastOneBoP = false;
        var BoPOnTrash = 0;
        var BoPOnBosses = 0;
        var BoPTotalUptime = 0;
        var BoPOnTrashSelf = 0;
        var BoPOnBossesSelf = 0;
        var gotAtLeastOneDivineFavor = false;
        var DivineFavorOnTrash = 0;
        var DivineFavorOnBosses = 0;
        var DivineFavorTotalUptime = 0;
        var gotAtLeastOneDivineIntervention = false;
        var DivineInterventionOnTrash = 0;
        var DivineInterventionOnBosses = 0;
        var DivineInterventionTotalUptime = 0;
        var gotAtLeastOneDivineProtection = false;
        var DivineProtectionOnTrash = 0;
        var DivineProtectionOnBosses = 0;
        var DivineProtectionTotalUptime = 0;
        var gotAtLeastOneDivineShield = false;
        var DivineShieldOnTrash = 0;
        var DivineShieldOnBosses = 0;
        var DivineShieldTotalUptime = 0;
        var gotAtLeastOneLayOnHands = false;
        var LayOnHandsOnTrash = 0;
        var LayOnHandsOnBosses = 0;
        var LayOnHandsTotalUptime = 0;
        var gotAtLeastOneManaTideTotem = false;
        var ManaTideTotemOnTrash = 0;
        var ManaTideTotemOnBosses = 0;
        var ManaTideTotalUptime = 0;
        var gotAtLeastOneVanish = false;
        var VanishOnTrash = 0;
        var VanishOnBosses = 0;
        var VanishTotalUptime = 0;
        var gotAtLeastOneBarkskin = false;
        var BarkskinOnTrash = 0;
        var BarkskinOnBosses = 0;
        var BarkskinTotalUptime = 0;
        var gotAtLeastOneForceOfNature = false;
        var ForceOfNatureOnTrash = 0;
        var ForceOfNatureOnBosses = 0;
        var ForceOfNatureTotalUptime = 0;
        var gotAtLeastOneSummonWaterElemental = false;
        var SummonWaterElementalOnTrash = 0;
        var SummonWaterElementalOnBosses = 0;
        var SummonWaterElementalTotalUptime = 0;
        var gotAtLeastOneInvisibility = false;
        var InvisibilityOnTrash = 0;
        var InvisibilityOnBosses = 0;
        var InvisibilityTotalUptime = 0;
        var gotAtLeastOneColdSnap = false;
        var ColdSnapOnTrash = 0;
        var ColdSnapOnBosses = 0;
        var ColdSnapTotalUptime = 0;
        var gotAtLeastOneAvengingWrath = false;
        var AvengingWrathOnTrash = 0;
        var AvengingWrathOnBosses = 0;
        var AvengingWrathTotalUptime = 0;
        var gotAtLeastOneDivineIllumination = false;
        var DivineIlluminationOnTrash = 0;
        var DivineIlluminationOnBosses = 0;
        var DivineIlluminationTotalUptime = 0;
        var gotAtLeastOneDevouringPlague = false;
        var DevouringPlagueOnTrash = 0;
        var DevouringPlagueOnBosses = 0;
        var DevouringPlagueTotalUptime = 0;
        var gotAtLeastOneDesperatePrayer = false;
        var DesperatePrayerOnTrash = 0;
        var DesperatePrayerOnBosses = 0;
        var DesperatePrayerTotalUptime = 0;
        var gotAtLeastOneGiftOfTheNaaru = false;
        var GiftOfTheNaaruOnTrash = 0;
        var GiftOfTheNaaruOnBosses = 0;
        var GiftOfTheNaaruTotalUptime = 0;
        var gotAtLeastOnePainSuppression = false;
        var PainSuppressionOnTrash = 0;
        var PainSuppressionOnBosses = 0;
        var PainSuppressionTotalUptime = 0;
        var gotAtLeastOneCloakOfShadows = false;
        var CloakOfShadowsOnTrash = 0;
        var CloakOfShadowsOnBosses = 0;
        var CloakOfShadowsTotalUptime = 0;
        var gotAtLeastOneBloodlust = false;
        var BloodlustOnTrash = 0;
        var BloodlustOnBosses = 0;
        var BloodlustTotalUptime = 0;
        var gotAtLeastOneAmplifyCurse = false;
        var AmplifyCurseOnTrash = 0;
        var AmplifyCurseOnBosses = 0;
        var AmplifyCurseTotalUptime = 0;
        var gotAtLeastOneHeroism = false;
        var HeroismOnTrash = 0;
        var HeroismOnBosses = 0;
        var HeroismTotalUptime = 0;
        var gotAtLeastOneShamanisticRage = false;
        var ShamanisticRageOnTrash = 0;
        var ShamanisticRageOnBosses = 0;
        var ShamanisticRageTotalUptime = 0;
        var gotAtLeastOneSoulshatter = false;
        var SoulshatterOnTrash = 0;
        var SoulshatterOnBosses = 0;
        var SoulshatterTotalUptime = 0;
        var gotAtLeastOneIcyVeins = false;
        var IcyVeinsOnTrash = 0;
        var IcyVeinsOnBosses = 0;
        var IcyVeinsTotalUptime = 0;
        var gotAtLeastOnePoM = false;
        var PoMOnTrash = 0;
        var PoMOnBosses = 0;
        var PoMTotalUptime = 0;
        var GreaterStoneshieldUses = 0;
        var LIPUses = 0;
        var FAPUses = 0;
        var RestorativeUses = 0;
        var SkullUsesOnTrash = 0;
        var SkullUsesOnBosses = 0;
        var BerserkingUsesOnTrash = 0;
        var BerserkingUsesOnBosses = 0;
        var ScrollsOfBlindingLightUsesOnBosses = 0;
        var ScrollsOfBlindingLightUsesOnTrash = 0;
        var QuagmirransEyeUsesOnBosses = 0;
        var QuagmirransEyeUsesOnTrash = 0;
        var ScarabOfTheInfiniteCycleUsesOnBosses = 0;
        var ScarabOfTheInfiniteCycleUsesOnTrash = 0;
        var MysticalSkyfireDiamondUsesOnBosses = 0;
        var MysticalSkyfireDiamondUsesOnTrash = 0;
        var MindQuickeningGemUsesOnBosses = 0;
        var MindQuickeningGemUsesOnTrash = 0;
        var BladeOfWizardryUsesOnBosses = 0;
        var BladeOfWizardryUsesOnTrash = 0;

        if (!onlyBosses) {
          playerDataTrash.entries.forEach(function (spellSelf, spellSelfCount) {
            if (spellSelf.guid == 10060 && spellSelf.total > 0 && spellSelf.targets != null && spellSelf.targets.length > 0) {
              spellSelf.targets.forEach(function (spellSelfTarget, spellSelfTargetCount) {
                if (spellSelfTarget.name == playerByNameAsc.name) {
                  gotAtLeastOnePI = true;
                  PIonTrashSelf = spellSelfTarget.total;
                }
              })
            } else if (spellSelf.guid == 29166 && spellSelf.total > 0 && spellSelf.targets != null && spellSelf.targets.length > 0) {
              spellSelf.targets.forEach(function (spellSelfTarget, spellSelfTargetCount) {
                if (spellSelfTarget.name == playerByNameAsc.name) {
                  gotAtLeastOneInnervate = true;
                  InnervateOnTrashSelf = spellSelfTarget.total;
                }
              })
            } else if (spellSelf.guid == 10278 && spellSelf.total > 0 && spellSelf.targets != null && spellSelf.targets.length > 0) {
              spellSelf.targets.forEach(function (spellSelfTarget, spellSelfTargetCount) {
                if (spellSelfTarget.name == playerByNameAsc.name) {
                  gotAtLeastOneBoP = true;
                  BoPOnTrashSelf = spellSelfTarget.total;
                }
              })
            }
          })

          buffsDataTrash.auras.forEach(function (buff, buffCount) {
            if (buff.guid == 10060 && buff.totalUses > 0) {
              gotAtLeastOnePI = true;
              PIonTrash = buff.bands.length;
            } else if ((buff.guid == 12042 || buff.guid == 28682) && buff.totalUses > 0) {
              gotAtLeastOneAPComb = true;
              APCombonTrash = buff.bands.length;
            } else if ((buff.guid == 12051) && buff.totalUses > 0) {
              gotAtLeastOneEvocation = true;
              EvocationOnTrash = buff.bands.length;
            } else if ((buff.guid == 5277 || buff.guid == 26669) && buff.totalUses > 0) {
              gotAtLeastOneEvasion = true;
              EvasionOnTrash = buff.bands.length;
            } else if ((buff.guid == 13750) && buff.totalUses > 0) {
              gotAtLeastOneAdrenalineRush = true;
              AdrenalineRushOnTrash = buff.bands.length;
            } else if ((buff.guid == 18499) && buff.totalUses > 0) {
              gotAtLeastOneBerserkerRage = true;
              BerserkerRageOnTrash = buff.bands.length;
            } else if ((buff.guid == 2687) && buff.totalUses > 0) {
              gotAtLeastOneBloodrage = true;
              BloodrageOnTrash = buff.bands.length;
            } else if ((buff.guid == 1161) && buff.totalUses > 0) {
              gotAtLeastOneChallengingShout = true;
              ChallengingShoutOnTrash = buff.bands.length;
            } else if ((buff.guid == 12292) && buff.totalUses > 0) {
              gotAtLeastOneDeathWish = true;
              DeathWishOnTrash = buff.bands.length;
            } else if ((buff.guid == 12975) && buff.totalUses > 0) {
              gotAtLeastOneLastStand = true;
              LastStandOnTrash = buff.bands.length;
            } else if ((buff.guid == 1719) && buff.totalUses > 0) {
              gotAtLeastOneRecklessness = true;
              RecklessnessOnTrash = buff.bands.length;
            } else if ((buff.guid == 20230) && buff.totalUses > 0) {
              gotAtLeastOneRetaliation = true;
              RetaliationOnTrash = buff.bands.length;
            } else if ((buff.guid == 871) && buff.totalUses > 0) {
              gotAtLeastOneShieldWall = true;
              ShieldWallOnTrash = buff.bands.length;
            } else if ((buff.guid == 17116 || buff.guid == 16188) && buff.totalUses > 0) {
              gotAtLeastOneNaturesSwiftness = true;
              NaturesSwiftnessOnTrash = buff.bands.length;
            } else if ((buff.guid == 26994) && buff.totalUses > 0) {
              gotAtLeastOneRebirth = true;
              RebirthOnTrash = buff.bands.length;
            } else if ((buff.guid == 5209) && buff.totalUses > 0) {
              gotAtLeastOneChallengingRoar = true;
              ChallengingRoarOnTrash = buff.bands.length;
            } else if ((buff.guid == 33357) && buff.totalUses > 0) {
              gotAtLeastOneDash = true;
              DashOnTrash = buff.bands.length;
            } else if ((buff.guid == 26999) && buff.totalUses > 0) {
              gotAtLeastOneFrenziedRegeneration = true;
              FrenziedRegenerationOnTrash = buff.bands.length;
            } else if ((buff.guid == 29166) && buff.totalUses > 0) {
              gotAtLeastOneInnervate = true;
              InnervateOnTrash = buff.bands.length;
            } else if ((buff.guid == 3045) && buff.totalUses > 0) {
              gotAtLeastOneRapidFire = true;
              RapidFireOnTrash = buff.bands.length;
            } else if ((buff.guid == 23989) && buff.totalUses > 0) {
              gotAtLeastOneReadiness = true;
              ReadinessOnTrash = buff.bands.length;
            } else if ((buff.guid == 19263) && buff.totalUses > 0) {
              gotAtLeastOneDeterrence = true;
              DeterrenceOnTrash = buff.bands.length;
            } else if ((buff.guid == 19574) && buff.totalUses > 0) {
              gotAtLeastOneBestialWrath = true;
              BestialWrathOnTrash = buff.bands.length;
            } else if ((buff.guid == 14751) && buff.totalUses > 0) {
              gotAtLeastOneInnerFocus = true;
              InnerFocusOnTrash = buff.bands.length;
            } else if ((buff.guid == 10278) && buff.totalUses > 0) {
              gotAtLeastOneBoP = true;
              BoPOnTrash = buff.bands.length;
            } else if ((buff.guid == 20216) && buff.totalUses > 0) {
              gotAtLeastOneDivineFavor = true;
              DivineFavorOnTrash = buff.bands.length;
            } else if ((buff.guid == 19752) && buff.totalUses > 0) {
              gotAtLeastOneDivineIntervention = true;
              DivineInterventionOnTrash = buff.bands.length;
            } else if ((buff.guid == 1020) && buff.totalUses > 0) {
              gotAtLeastOneDivineShield = true;
              DivineShieldOnTrash = buff.bands.length;
            } else if ((buff.guid == 5573) && buff.totalUses > 0) {
              gotAtLeastOneDivineProtection = true;
              DivineProtectionOnTrash = buff.bands.length;
            } else if ((buff.guid == 27154) && buff.totalUses > 0) {
              gotAtLeastOneLayOnHands = true;
              LayOnHandsOnTrash = buff.bands.length;
            } else if ((buff.guid == 16190 || buff.guid == 17359 || buff.guid == 17354) && buff.totalUses > 0) {
              gotAtLeastOneManaTideTotem = true;
              ManaTideTotemOnTrash = buff.bands.length;
            } else if ((buff.guid == 26889) && buff.totalUses > 0) {
              gotAtLeastOneVanish = true;
              VanishOnTrash = buff.bands.length;
            } else if ((buff.guid == 22812) && buff.totalUses > 0) {
              gotAtLeastOneBarkskin = true;
              BarkskinOnTrash = buff.bands.length;
            } else if ((buff.guid == 31884) && buff.totalUses > 0) {
              gotAtLeastOneAvengingWrath = true;
              AvengingWrathOnTrash = buff.bands.length;
            } else if ((buff.guid == 31842) && buff.totalUses > 0) {
              gotAtLeastOneDivineIllumination = true;
              DivineIlluminationOnTrash = buff.bands.length;
            } else if ((buff.guid == 33206) && buff.totalUses > 0) {
              gotAtLeastOnePainSuppression = true;
              PainSuppressionOnTrash = buff.bands.length;
            } else if ((buff.guid == 31224) && buff.totalUses > 0) {
              gotAtLeastOneCloakOfShadows = true;
              CloakOfShadowsOnTrash = buff.bands.length;
            } else if ((buff.guid == 2825) && buff.totalUses > 0) {
              gotAtLeastOneBloodlust = true;
              BloodlustOnTrash = buff.bands.length;
            } else if ((buff.guid == 18288) && buff.totalUses > 0) {
              gotAtLeastOneAmplifyCurse = true;
              AmplifyCurseOnTrash = buff.bands.length;
            } else if ((buff.guid == 32182) && buff.totalUses > 0) {
              gotAtLeastOneHeroism = true;
              HeroismOnTrash = buff.bands.length;
            } else if ((buff.guid == 30823) && buff.totalUses > 0) {
              gotAtLeastOneShamanisticRage = true;
              ShamanisticRageOnTrash = buff.bands.length;
            } else if ((buff.guid == 12472) && buff.totalUses > 0) {
              gotAtLeastOneIcyVeins = true;
              IcyVeinsOnTrash = buff.bands.length;
            } else if ((buff.guid == 26635) && buff.totalUses > 0) {
              BerserkingUsesOnTrash = buff.totalUses;
            } else if ((buff.guid == 18803) && buff.totalUses > 0) {
              MysticalSkyfireDiamondUsesOnTrash = buff.totalUses;
            } else if ((buff.guid == 38317) && buff.totalUses > 0) {
              BladeOfWizardryUsesOnTrash = buff.totalUses;
            }
          })
          playerDataTrash.entries.forEach(function (spell, spellCount) {
            if (playerByNameAsc.type == "Priest" && (spell.guid == 10060) && spell.total > 0) {
              gotAtLeastOnePI = true;
              PIonTrash = spell.total;
            } else if (playerByNameAsc.type == "Druid" && (spell.guid == 29166) && spell.total > 0) {
              gotAtLeastOneInnervate = true;
              InnervateOnTrash = spell.total;
            } else if (playerByNameAsc.type == "Shaman" && (spell.guid == 32182) && spell.total > 0) {
              gotAtLeastOneHeroism = true;
              HeroismOnTrash = spell.total;
            } else if (playerByNameAsc.type == "Shaman" && (spell.guid == 2825) && spell.total > 0) {
              gotAtLeastOneBloodlust = true;
              BloodlustOnTrash = spell.total;
            } else if (playerByNameAsc.type == "Paladin" && (spell.guid == 10278) && spell.total > 0) {
              gotAtLeastOneBoP = true;
              BoPOnTrash = spell.total;
            } else if ((spell.guid == 12042 || spell.guid == 28682) && spell.total > 0) {
              gotAtLeastOneAPComb = true;
              APCombonTrash = spell.total;
            } else if ((spell.guid == 12051) && spell.total > 0) {
              gotAtLeastOneEvocation = true;
              EvocationOnTrash = spell.total;
            } else if ((spell.guid == 1161) && spell.total > 0) {
              gotAtLeastOneChallengingShout = true;
              ChallengingShoutOnTrash = spell.total;
            } else if ((spell.guid == 2687) && spell.total > 0) {
              gotAtLeastOneBloodrage = true;
              BloodrageOnTrash = spell.total;
            } else if ((spell.guid == 12292) && spell.total > 0) {
              gotAtLeastOneDeathWish = true;
              DeathWishOnTrash = spell.total;
            } else if ((spell.guid == 12975) && spell.total > 0) {
              gotAtLeastOneLastStand = true;
              LastStandOnTrash = spell.total;
            } else if ((spell.guid == 20230) && spell.total > 0) {
              gotAtLeastOneRetaliation = true;
              RetaliationOnTrash = spell.total;
            } else if ((spell.guid == 17116 || spell.guid == 16188) && spell.total > 0) {
              gotAtLeastOneNaturesSwiftness = true;
              NaturesSwiftnessOnTrash = spell.total;
            } else if ((spell.guid == 5209) && spell.total > 0) {
              gotAtLeastOneChallengingRoar = true;
              ChallengingRoarOnTrash = spell.total;
            } else if ((spell.guid == 16166) && spell.total > 0) {
              gotAtLeastOneElementalMastery = true;
              ElementalMasteryOnTrash = spell.total;
            } else if ((spell.guid == 26994) && spell.total > 0) {
              gotAtLeastOneRebirth = true;
              RebirthOnTrash = spell.total;
            } else if ((spell.guid == 1719) && spell.total > 0) {
              gotAtLeastOneRecklessness = true;
              RecklessnessOnTrash = spell.total;
            } else if ((spell.guid == 871) && spell.total > 0) {
              gotAtLeastOneShieldWall = true;
              ShieldWallOnTrash = spell.total;
            } else if ((spell.guid == 33206) && spell.total > 0) {
              gotAtLeastOnePainSuppression = true;
              PainSuppressionOnTrash = spell.total;
            } else if ((spell.guid == 3045) && spell.total > 0) {
              gotAtLeastOneRapidFire = true;
              RapidFireOnTrash = spell.total;
            } else if ((spell.guid == 23989) && spell.total > 0) {
              gotAtLeastOneReadiness = true;
              ReadinessOnTrash = spell.total;
            } else if ((spell.guid == 19263) && spell.total > 0) {
              gotAtLeastOneDeterrence = true;
              DeterrenceOnTrash = spell.total;
            } else if ((spell.guid == 19574) && spell.total > 0) {
              gotAtLeastOneBestialWrath = true;
              BestialWrathOnTrash = spell.total;
            } else if ((spell.guid == 14751) && spell.total > 0) {
              gotAtLeastOneInnerFocus = true;
              InnerFocusOnTrash = spell.total;
            } else if ((spell.guid == 34433) && spell.total > 0) {
              gotAtLeastOneShadowfiend = true;
              ShadowfiendOnTrash = spell.total;
            } else if ((spell.guid == 20216) && spell.total > 0) {
              gotAtLeastOneDivineFavor = true;
              DivineFavorOnTrash = spell.total;
            } else if ((spell.guid == 19752) && spell.total > 0) {
              gotAtLeastOneDivineIntervention = true;
              DivineInterventionOnTrash = spell.total;
            } else if ((spell.guid == 1020) && spell.total > 0) {
              gotAtLeastOneDivineShield = true;
              DivineShieldOnTrash = spell.total;
            } else if ((spell.guid == 5573) && spell.total > 0) {
              gotAtLeastOneDivineProtection = true;
              DivineProtectionOnTrash = spell.total;
            } else if ((spell.guid == 27154) && spell.total > 0) {
              gotAtLeastOneLayOnHands = true;
              LayOnHandsOnTrash = spell.total;
            } else if ((spell.guid == 16190 || spell.guid == 17359 || spell.guid == 17354) && spell.total > 0) {
              gotAtLeastOneManaTideTotem = true;
              ManaTideTotemOnTrash = spell.total;
            } else if ((spell.guid == 26889) && spell.total > 0) {
              gotAtLeastOneVanish = true;
              VanishOnTrash = spell.total;
            } else if ((spell.guid == 26983) && spell.total > 0) {
              gotAtLeastOneTranquility = true;
              TranquilityOnTrash = spell.total;
            } else if ((spell.guid == 33831) && spell.total > 0) {
              gotAtLeastOneForceOfNature = true;
              ForceOfNatureOnTrash = spell.total;
            } else if ((spell.guid == 31687) && spell.total > 0) {
              gotAtLeastOneSummonWaterElemental = true;
              SummonWaterElementalOnTrash = spell.total;
            } else if ((spell.guid == 66) && spell.total > 0) {
              gotAtLeastOneInvisibility = true;
              InvisibilityOnTrash = spell.total;
            } else if ((spell.guid == 11958) && spell.total > 0) {
              gotAtLeastOneColdSnap = true;
              ColdSnapOnTrash = spell.total;
            } else if ((spell.guid == 25437) && spell.total > 0) {
              gotAtLeastOneDesperatePrayer = true;
              DesperatePrayerOnTrash = spell.total;
            } else if ((spell.guid == 28880) && spell.total > 0) {
              gotAtLeastOneGiftOfTheNaaru = true;
              GiftOfTheNaaruOnTrash = spell.total;
            } else if ((spell.guid == 29858) && spell.total > 0) {
              gotAtLeastOneSoulshatter = true;
              SoulshatterOnTrash = spell.total;
            } else if ((spell.guid == 25467) && spell.total > 0) {
              gotAtLeastOneDevouringPlague = true;
              DevouringPlagueOnTrash = spell.total;
            } else if ((spell.guid == 12472) && spell.total > 0) {
              gotAtLeastOneIcyVeins = true;
              IcyVeinsOnTrash = spell.total;
            } else if ((spell.guid == 12043) && spell.total > 0) {
              gotAtLeastOnePoM = true;
              PoMOnTrash = spell.total;
            } else if ((spell.guid == 40396) && spell.total > 0) {
              SkullUsesOnTrash = spell.total;
            } else if ((spell.guid == 23733) && spell.total > 0) {
              ScrollsOfBlindingLightUsesOnTrash = spell.total;
            } else if ((spell.guid == 33370 || spell.guid == 33369) && spell.total > 0) {
              QuagmirransEyeUsesOnTrash = spell.total;
            } else if ((spell.guid == 33953 || spell.guid == 33370) && spell.total > 0) {
              ScarabOfTheInfiniteCycleUsesOnTrash = spell.total;
            } else if ((spell.guid == 23723) && spell.total > 0) {
              MindQuickeningGemUsesOnTrash = spell.total;
            }
          })
        }

        playerData.entries.forEach(function (spellSelf, spellSelfCount) {
          if (spellSelf.guid == 10060 && spellSelf.total > 0 && spellSelf.targets != null && spellSelf.targets.length > 0) {
            spellSelf.targets.forEach(function (spellSelfTarget, spellSelfTargetCount) {
              if (spellSelfTarget.name == playerByNameAsc.name) {
                gotAtLeastOnePI = true;
                PIOnBossesSelf = spellSelfTarget.total - PIonTrashSelf;
              }
            })
          } else if (spellSelf.guid == 29166 && spellSelf.total > 0 && spellSelf.targets != null && spellSelf.targets.length > 0) {
            spellSelf.targets.forEach(function (spellSelfTarget, spellSelfTargetCount) {
              if (spellSelfTarget.name == playerByNameAsc.name) {
                gotAtLeastOneInnervate = true;
                InnervateOnBossesSelf = spellSelfTarget.total - InnervateOnTrashSelf;
              }
            })
          } else if (spellSelf.guid == 10278 && spellSelf.total > 0 && spellSelf.targets != null && spellSelf.targets.length > 0) {
            spellSelf.targets.forEach(function (spellSelfTarget, spellSelfTargetCount) {
              if (spellSelfTarget.name == playerByNameAsc.name) {
                gotAtLeastOneBoP = true;
                BoPOnBossesSelf = spellSelfTarget.total - BoPOnTrashSelf;
              }
            })
          }
        })

        var PIoverwrittenWithCasts = false;
        var InnervateUsedBySelf = false;
        var BoPUsedBySelf = false;
        buffsData.auras.forEach(function (buff, buffCount) {
          if (buff.guid == 10060 && buff.totalUses > 0) {
            gotAtLeastOnePI = true;
            PITotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            PIOnBosses = buff.bands.length - PIonTrash;
          } else if ((buff.guid == 12042 || buff.guid == 28682) && buff.totalUses > 0) {
            gotAtLeastOneAPComb = true;
            APCombTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            APCombOnBosses = buff.bands.length - APCombonTrash;
          } else if ((buff.guid == 12051) && buff.totalUses > 0) {
            gotAtLeastOneEvocation = true;
            EvocationTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            EvocationOnBosses = buff.bands.length - EvocationOnTrash;
          } else if ((buff.guid == 5277 || buff.guid == 26669) && buff.totalUses > 0) {
            gotAtLeastOneEvasion = true;
            EvasionTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            EvasionOnBosses = buff.bands.length - EvasionOnTrash;
          } else if ((buff.guid == 13750) && buff.totalUses > 0) {
            gotAtLeastOneAdrenalineRush = true;
            AdrenalineRushTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            AdrenalineRushOnBosses = buff.bands.length - AdrenalineRushOnTrash;
          } else if ((buff.guid == 18499) && buff.totalUses > 0) {
            gotAtLeastOneBerserkerRage = true;
            BerserkerRageTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            BerserkerRageOnBosses = buff.bands.length - BerserkerRageOnTrash;
          } else if ((buff.guid == 2687) && buff.totalUses > 0) {
            gotAtLeastOneBloodrage = true;
            BloodrageTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            BloodrageOnBosses = buff.bands.length - BloodrageOnTrash;
          } else if ((buff.guid == 1161) && buff.totalUses > 0) {
            gotAtLeastOneChallengingShout = true;
            ChallengingShoutTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            ChallengingShoutOnBosses = buff.bands.length - ChallengingShoutOnTrash;
          } else if ((buff.guid == 12292) && buff.totalUses > 0) {
            gotAtLeastOneDeathWish = true;
            DeathWishTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            DeathWishOnBosses = buff.bands.length - DeathWishOnTrash;
          } else if ((buff.guid == 12975) && buff.totalUses > 0) {
            gotAtLeastOneLastStand = true;
            LastStandTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            LastStandOnBosses = buff.bands.length - LastStandOnTrash;
          } else if ((buff.guid == 1719) && buff.totalUses > 0) {
            gotAtLeastOneRecklessness = true;
            RecklessnessTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            RecklessnessOnBosses = buff.bands.length - RecklessnessOnTrash;
          } else if ((buff.guid == 20230) && buff.totalUses > 0) {
            gotAtLeastOneRetaliation = true;
            RetaliationTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            RetaliationOnBosses = buff.bands.length - RetaliationOnTrash;
          } else if ((buff.guid == 871) && buff.totalUses > 0) {
            gotAtLeastOneShieldWall = true;
            ShieldWallTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            ShieldWallOnBosses = buff.bands.length - ShieldWallOnTrash;
          } else if ((buff.guid == 17116 || buff.guid == 16188) && buff.totalUses > 0) {
            gotAtLeastOneNaturesSwiftness = true;
            NaturesSwiftnessTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            NaturesSwiftnessOnBosses = buff.bands.length - NaturesSwiftnessOnTrash;
          } else if ((buff.guid == 26994) && buff.totalUses > 0) {
            gotAtLeastOneRebirth = true;
            RebirthTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            RebirthOnBosses = buff.bands.length - RebirthOnTrash;
          } else if ((buff.guid == 5209) && buff.totalUses > 0) {
            gotAtLeastOneChallengingRoar = true;
            ChallengingRoarTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            ChallengingRoarOnBosses = buff.bands.length - ChallengingRoarOnTrash;
          } else if ((buff.guid == 33357) && buff.totalUses > 0) {
            gotAtLeastOneDash = true;
            DashTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            DashOnBosses = buff.bands.length - DashOnTrash;
          } else if ((buff.guid == 26999) && buff.totalUses > 0) {
            gotAtLeastOneFrenziedRegeneration = true;
            FrenziedRegenerationTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            FrenziedRegenerationOnBosses = buff.bands.length - FrenziedRegenerationOnTrash;
          } else if ((buff.guid == 29166) && buff.totalUses > 0) {
            gotAtLeastOneInnervate = true;
            InnervateTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            InnervateOnBosses = buff.bands.length - InnervateOnTrash;
          } else if ((buff.guid == 3045) && buff.totalUses > 0) {
            gotAtLeastOneRapidFire = true;
            RapidFireTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            RapidFireOnBosses = buff.bands.length - RapidFireOnTrash;
          } else if ((buff.guid == 23989) && buff.totalUses > 0) {
            gotAtLeastOneReadiness = true;
            ReadinessTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            ReadinessOnBosses = buff.bands.length - ReadinessOnTrash;
          } else if ((buff.guid == 19263) && buff.totalUses > 0) {
            gotAtLeastOneDeterrence = true;
            DeterrenceTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            DeterrenceOnBosses = buff.bands.length - DeterrenceOnTrash;
          } else if ((buff.guid == 19574) && buff.totalUses > 0) {
            gotAtLeastOneBestialWrath = true;
            BestialWrathTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            BestialWrathOnBosses = buff.bands.length - BestialWrathOnTrash;
          } else if ((buff.guid == 14751) && buff.totalUses > 0) {
            gotAtLeastOneInnerFocus = true;
            InnerFocusTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            InnerFocusOnBosses = buff.bands.length - InnerFocusOnTrash;
          } else if ((buff.guid == 10278) && buff.totalUses > 0) {
            gotAtLeastOneBoP = true;
            BoPTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            BoPOnBosses = buff.bands.length - BoPOnTrash;
          } else if ((buff.guid == 20216) && buff.totalUses > 0) {
            gotAtLeastOneDivineFavor = true;
            DivineFavorTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            DivineFavorOnBosses = buff.bands.length - DivineFavorOnTrash;
          } else if ((buff.guid == 19752) && buff.totalUses > 0) {
            gotAtLeastOneDivineIntervention = true;
            DivineInterventionTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            DivineInterventionOnBosses = buff.bands.length - DivineInterventionOnTrash;
          } else if ((buff.guid == 1020) && buff.totalUses > 0) {
            gotAtLeastOneDivineShield = true;
            DivineShieldTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            DivineShieldOnBosses = buff.bands.length - DivineShieldOnTrash;
          } else if ((buff.guid == 5573) && buff.totalUses > 0) {
            gotAtLeastOneDivineProtection = true;
            DivineProtectionTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            DivineProtectionOnBosses = buff.bands.length - DivineProtectionOnTrash;
          } else if ((buff.guid == 27154) && buff.totalUses > 0) {
            gotAtLeastOneLayOnHands = true;
            LayOnHandsTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            LayOnHandsOnBosses = buff.bands.length - LayOnHandsOnTrash;
          } else if ((buff.guid == 16190 || buff.guid == 17359 || buff.guid == 17354) && buff.totalUses > 0) {
            gotAtLeastOneManaTideTotem = true;
            ManaTideTotemTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            ManaTideTotemOnBosses = buff.bands.length - ManaTideTotemOnTrash;
          } else if ((buff.guid == 26889) && buff.totalUses > 0) {
            gotAtLeastOneVanish = true;
            VanishTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            VanishOnBosses = buff.bands.length - VanishOnTrash;
          } else if ((buff.guid == 22812) && buff.totalUses > 0) {
            gotAtLeastOneBarkskin = true;
            BarkskinTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            BarkskinOnBosses = buff.bands.length - BarkskinOnTrash;
          } else if ((buff.guid == 31884) && buff.totalUses > 0) {
            gotAtLeastOneAvengingWrath = true;
            AvengingWrathTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            AvengingWrathOnBosses = buff.bands.length - AvengingWrathOnTrash;
          } else if ((buff.guid == 31842) && buff.totalUses > 0) {
            gotAtLeastOneDivineIllumination = true;
            DivineIlluminationTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            DivineIlluminationOnBosses = buff.bands.length - DivineIlluminationOnTrash;
          } else if ((buff.guid == 33206) && buff.totalUses > 0) {
            gotAtLeastOnePainSuppression = true;
            PainSuppressionTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            PainSuppressionOnBosses = buff.bands.length - PainSuppressionOnTrash;
          } else if ((buff.guid == 31224) && buff.totalUses > 0) {
            gotAtLeastOneCloakOfShadows = true;
            CloakOfShadowsTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            CloakOfShadowsOnBosses = buff.bands.length - CloakOfShadowsOnTrash;
          } else if ((buff.guid == 2825) && buff.totalUses > 0) {
            gotAtLeastOneBloodlust = true;
            BloodlustTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            BloodlustOnBosses = buff.bands.length - BloodlustOnTrash;
          } else if ((buff.guid == 18288) && buff.totalUses > 0) {
            gotAtLeastOneAmplifyCurse = true;
            AmplifyCurseTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            AvengingWrathOnBosses = buff.bands.length - AvengingWrathOnTrash;
          } else if ((buff.guid == 32182) && buff.totalUses > 0) {
            gotAtLeastOneHeroism = true;
            HeroismTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            HeroismOnBosses = buff.bands.length - HeroismOnTrash;
          } else if ((buff.guid == 30823) && buff.totalUses > 0) {
            gotAtLeastOneShamanisticRage = true;
            ShamanisticRageTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            ShamanisticRageOnBosses = buff.bands.length - ShamanisticRageOnTrash;
          } else if ((buff.guid == 12472) && buff.totalUses > 0) {
            gotAtLeastOneIcyVeins = true;
            IcyVeinsTotalUptime = buff.totalUptime * 100 / Math.abs(raidDuration);
            IcyVeinsOnBosses = buff.bands.length - IcyVeinsOnTrash;
          } else if ((buff.guid == 17540) && buff.totalUses > 0) {
            GreaterStoneshieldUses = buff.bands.length;
          } else if ((buff.guid == 24364 || buff.guid == 6615) && buff.totalUses > 0) {
            FAPUses = buff.bands.length;
          } else if ((buff.guid == 3169) && buff.totalUses > 0) {
            LIPUses = buff.bands.length;
          } else if ((buff.guid == 11359) && buff.totalUses > 0) {
            RestorativeUses = buff.bands.length;
          } else if ((buff.guid == 26635) && buff.totalUses > 0) {
            BerserkingUsesOnBosses = buff.totalUses - BerserkingUsesOnTrash;
          } else if ((buff.guid == 18803) && buff.totalUses > 0) {
            MysticalSkyfireDiamondUsesOnBosses = buff.totalUses - MysticalSkyfireDiamondUsesOnTrash;
          } else if ((buff.guid == 38317) && buff.totalUses > 0) {
            BladeOfWizardryUsesOnBosses = buff.totalUses - BladeOfWizardryUsesOnTrash;
          }
        })

        var usedIceblockOrIceBarrier = false;
        var usedFireVulnerabilityOrBlastWave = false;
        playerData.entries.forEach(function (spell, spellCount) {
          if (playerByNameAsc.type == "Priest" && spell.guid == 10060 && spell.total > 0) {
            gotAtLeastOnePI = true;
            PITotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            PIOnBosses = spell.total - PIonTrash;
            PIoverwrittenWithCasts = true;
          } else if (playerByNameAsc.type == "Druid" && (spell.guid == 29166) && spell.total > 0) {
            gotAtLeastOneInnervate = true;
            InnervateTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            InnervateOnBosses = spell.total - InnervateOnTrash;
            InnervateUsedBySelf = true;
          } else if (playerByNameAsc.type == "Shaman" && (spell.guid == 32182) && spell.total > 0) {
            gotAtLeastOneHeroism = true;
            HeroismTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            HeroismOnBosses = spell.total - HeroismOnTrash;
          } else if (playerByNameAsc.type == "Shaman" && (spell.guid == 2825) && spell.total > 0) {
            gotAtLeastOneBloodlust = true;
            BloodlustTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            BloodlustOnBosses = spell.total - BloodlustOnTrash;
          } else if (playerByNameAsc.type == "Paladin" && (spell.guid == 10278) && spell.total > 0) {
            gotAtLeastOneBoP = true;
            BoPTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            BoPOnBosses = spell.total - BoPOnTrash;
            BoPUsedBySelf = true;
          } else if ((spell.guid == 12042 || spell.guid == 28682) && spell.total > 0) {
            gotAtLeastOneAPComb = true;
            APCombTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            APCombOnBosses = spell.total - APCombonTrash;
          } else if ((spell.guid == 12051) && spell.total > 0) {
            gotAtLeastOneEvocation = true;
            EvocationTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            EvocationOnBosses = spell.total - EvocationOnTrash;
          } else if ((spell.guid == 1161) && spell.total > 0) {
            gotAtLeastOneChallengingShout = true;
            ChallengingShoutTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ChallengingShoutOnBosses = spell.total - ChallengingShoutOnTrash;
          } else if ((spell.guid == 2687) && spell.total > 0) {
            gotAtLeastOneBloodrage = true;
            BloodrageTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            BloodrageOnBosses = spell.total - BloodrageOnTrash;
          } else if ((spell.guid == 12292) && spell.total > 0) {
            gotAtLeastOneDeathWish = true;
            DeathWishTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            DeathWishOnBosses = spell.total - DeathWishOnTrash;
          } else if ((spell.guid == 12975) && spell.total > 0) {
            gotAtLeastOneLastStand = true;
            LastStandTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            LastStandOnBosses = spell.total - LastStandOnTrash;
          } else if ((spell.guid == 20230) && spell.total > 0) {
            gotAtLeastOneRetaliation = true;
            RetaliationTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            RetaliationOnBosses = spell.total - RetaliationOnTrash;
          } else if ((spell.guid == 17116 || spell.guid == 16188) && spell.total > 0) {
            gotAtLeastOneNaturesSwiftness = true;
            NaturesSwiftnessTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            NaturesSwiftnessOnBosses = spell.total - NaturesSwiftnessOnTrash;
          } else if ((spell.guid == 16166) && spell.total > 0) {
            gotAtLeastOneElementalMastery = true;
            ElementalMasteryTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ElementalMasteryOnBosses = spell.total - ElementalMasteryOnTrash;
          } else if ((spell.guid == 26994) && spell.total > 0) {
            gotAtLeastOneRebirth = true;
            RebirthTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            RebirthOnBosses = spell.total - RebirthOnTrash;
          } else if ((spell.guid == 5209) && spell.total > 0) {
            gotAtLeastOneChallengingRoar = true;
            ChallengingRoarTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ChallengingRoarOnBosses = spell.total - ChallengingRoarOnTrash;
          } else if ((spell.guid == 3045) && spell.total > 0) {
            gotAtLeastOneRapidFire = true;
            RapidFireTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            RapidFireOnBosses = spell.total - RapidFireOnTrash;
          } else if ((spell.guid == 33206) && spell.total > 0) {
            gotAtLeastOnePainSuppression = true;
            PainSuppressionTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            PainSuppressionOnBosses = spell.total - PainSuppressionOnTrash;
          } else if ((spell.guid == 23989) && spell.total > 0) {
            gotAtLeastOneReadiness = true;
            ReadinessTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ReadinessOnBosses = spell.total - ReadinessOnTrash;
          } else if ((spell.guid == 19263) && spell.total > 0) {
            gotAtLeastOneDeterrence = true;
            DeterrenceTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            DeterrenceOnBosses = spell.total - DeterrenceOnTrash;
          } else if ((spell.guid == 19574) && spell.total > 0) {
            gotAtLeastOneBestialWrath = true;
            BestialWrathTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            BestialWrathOnBosses = spell.total - BestialWrathOnTrash;
          } else if ((spell.guid == 14751) && spell.total > 0) {
            gotAtLeastOneInnerFocus = true;
            InnerFocusTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            InnerFocusOnBosses = spell.total - InnerFocusOnTrash;
          } else if ((spell.guid == 34433) && spell.total > 0) {
            gotAtLeastOneShadowfiend = true;
            ShadowfiendTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ShadowfiendOnBosses = spell.total - ShadowfiendOnTrash;
          } else if ((spell.guid == 20216) && spell.total > 0) {
            gotAtLeastOneDivineFavor = true;
            DivineFavorTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            DivineFavorOnBosses = spell.total - DivineFavorOnTrash;
          } else if ((spell.guid == 19752) && spell.total > 0) {
            gotAtLeastOneDivineIntervention = true;
            DivineInterventionTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            DivineInterventionOnBosses = spell.total - DivineInterventionOnTrash;
          } else if ((spell.guid == 5573) && spell.total > 0) {
            gotAtLeastOneDivineProtection = true;
            DivineProtectionTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            DivineProtectionOnBosses = spell.total - DivineProtectionOnTrash;
          } else if ((spell.guid == 1020) && spell.total > 0) {
            gotAtLeastOneDivineShield = true;
            DivineShieldTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            DivineShieldOnBosses = spell.total - DivineShieldOnTrash;
          } else if ((spell.guid == 1719) && spell.total > 0) {
            gotAtLeastOneRecklessness = true;
            RecklessnessTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            RecklessnessOnBosses = spell.total - RecklessnessOnTrash;
          } else if ((spell.guid == 871) && spell.total > 0) {
            gotAtLeastOneShieldWall = true;
            ShieldWallTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ShieldWallOnBosses = spell.total - ShieldWallOnTrash;
          } else if ((spell.guid == 27154) && spell.total > 0) {
            gotAtLeastOneLayOnHands = true;
            LayOnHandsTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            LayOnHandsOnBosses = spell.total - LayOnHandsOnTrash;
          } else if ((spell.guid == 16190 || spell.guid == 17359 || spell.guid == 17354) && spell.total > 0) {
            gotAtLeastOneManaTideTotem = true;
            ManaTideTotemTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ManaTideTotemOnBosses = spell.total - ManaTideTotemOnTrash;
          } else if ((spell.guid == 26889) && spell.total > 0) {
            gotAtLeastOneVanish = true;
            VanishTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            VanishOnBosses = spell.total - VanishOnTrash;
          } else if ((spell.guid == 26983) && spell.total > 0) {
            gotAtLeastOneTranquility = true;
            TranquilityTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            TranquilityOnBosses = spell.total - TranquilityOnTrash;
          } else if ((spell.guid == 33831) && spell.total > 0) {
            gotAtLeastOneForceOfNature = true;
            ForceOfNatureTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ForceOfNatureOnBosses = spell.total - ForceOfNatureOnTrash;
          } else if ((spell.guid == 31687) && spell.total > 0) {
            gotAtLeastOneSummonWaterElemental = true;
            SummonWaterElementalTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            SummonWaterElementalOnBosses = spell.total - SummonWaterElementalOnTrash;
          } else if ((spell.guid == 66) && spell.total > 0) {
            gotAtLeastOneInvisibility = true;
            InvisibilityTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            InvisibilityOnBosses = spell.total - InvisibilityOnTrash;
          } else if ((spell.guid == 11958) && spell.total > 0) {
            gotAtLeastOneColdSnap = true;
            ColdSnapTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            ColdSnapOnBosses = spell.total - ColdSnapOnTrash;
          } else if ((spell.guid == 25437) && spell.total > 0) {
            gotAtLeastOneDesperatePrayer = true;
            DesperatePrayerTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            DesperatePrayerOnBosses = spell.total - DesperatePrayerOnTrash;
          } else if ((spell.guid == 28880) && spell.total > 0) {
            gotAtLeastOneGiftOfTheNaaru = true;
            GiftOfTheNaaruTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            GiftOfTheNaaruOnBosses = spell.total - GiftOfTheNaaruOnTrash;
          } else if ((spell.guid == 29858) && spell.total > 0) {
            gotAtLeastOneSoulshatter = true;
            SoulshatterTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            SoulshatterOnBosses = spell.total - SoulshatterOnTrash;
          } else if ((spell.guid == 12472) && spell.total > 0) {
            gotAtLeastOneIcyVeins = true;
            IcyVeinsTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            IcyVeinsOnBosses = spell.total - IcyVeinsOnTrash;
          } else if ((spell.guid == 25467) && spell.total > 0) {
            gotAtLeastOneDevouringPlague = true;
            DevouringPlagueTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            DevouringPlagueOnBosses = spell.total - DevouringPlagueOnTrash;
          } else if ((spell.guid == 12043) && spell.total > 0) {
            gotAtLeastOnePoM = true;
            PoMTotalUptime = spell.uptime * 100 / Math.abs(raidDuration);
            PoMOnBosses = spell.total - PoMOnTrash;
          } else if (playerByNameAsc.type == "Mage") {
            if (spell.guid == 13021 && spell.total > 0)
              usedFireVulnerabilityOrBlastWave = true;
            else if (spell.guid == 13033 && spell.total > 0)
              usedIceblockOrIceBarrier = true;
            else if (spell.guid == 11958 && spell.total > 0)
              usedIceblockOrIceBarrier = true;
          } else if ((spell.guid == 40396) && spell.total > 0) {
            SkullUsesOnBosses = spell.total - SkullUsesOnTrash;
          } else if ((spell.guid == 23733) && spell.total > 0) {
            ScrollsOfBlindingLightUsesOnBosses = spell.total - ScrollsOfBlindingLightUsesOnTrash;
          } else if ((spell.guid == 33370 || spell.guid == 33369) && spell.total > 0) {
            QuagmirransEyeUsesOnBosses = spell.total - QuagmirransEyeUsesOnTrash;
          } else if ((spell.guid == 33953 || spell.guid == 33370) && spell.total > 0) {
            ScarabOfTheInfiniteCycleUsesOnBosses = spell.total - ScarabOfTheInfiniteCycleUsesOnTrash;
          } else if ((spell.guid == 23723) && spell.total > 0) {
            MindQuickeningGemUsesOnBosses = spell.total - MindQuickeningGemUsesOnTrash;
          }
        })

        var usedArcaniteDragonling = 0;
        if (damageDoneData.toString().indexOf('"id":16022') > -1)
          usedArcaniteDragonling = 1;

        var activePercentageTotal = 0;
        allPlayersCasting.entries.forEach(function (playerCasting, playerCastingCount) {
          if (playerCasting.name == playerByNameAsc.name) {
            activePercentageTotal = Math.round(playerCasting.activeTime * 100 / allPlayersCasting.totalTime);
          }
        })

        var activePercentageTotalOnTrash = 0;
        if (onlyTrash || (!onlyBosses && (onlyFightNr == null || onlyFightNr.toString().length == 0))) {
          allPlayersCastingOnTrash.entries.forEach(function (playerCasting, playerCastingCount) {
            if (playerCasting.name == playerByNameAsc.name) {
              activePercentageTotalOnTrash = Math.round(playerCasting.activeTime * 100 / allPlayersCastingOnTrash.totalTime);
            }
          })
        }

        var usedTemporaryWeaponEnchant = 0;
        var trinketsUsed = "";
        var found = 0;
        var total = 0;
        var healerCount = 0;
        var dpsCount = 0;
        var dpsSpec = "";
        var tankCount = 0;
        var gearSpellHaste = 0;
        bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
          bossSummaryData.composition.forEach(function (raidMember, raidMemberCount) {
            if (raidMember.id == playerByNameAsc.id) {
              raidMember.specs.forEach(function (playerSpec, playerSpecCount) {
                if (playerSpec.role != null && playerSpec.role.toString().length > 0) {
                  if (playerSpec.role == "healer")
                    healerCount++;
                  else if (playerSpec.role == "dps") {
                    dpsCount++;
                    if (playerSpec.spec != null && playerSpec.spec.toString().length > 0)
                      dpsSpec = playerSpec.spec;
                  } else if (playerSpec.role == "tank")
                    tankCount++;
                }
              })
            }
          })
          var increaseTotal = false;
          var increaseFound = false;
          if (increaseTotal == false && bossSummaryData.playerDetails != null && bossSummaryData.playerDetails.dps != null && bossSummaryData.playerDetails.dps.length > 0) {
            var fightGearSpellHaste = 0;
            bossSummaryData.playerDetails.dps.forEach(function (playerInfo, playerInfoCount) {
              if (playerInfo.name == playerByNameAsc.name) {
                if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                  increaseTotal = true;
                  playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                    if (item.slot.toString() == "15" || item.slot.toString() == "16") {
                      if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                        increaseTotal = false;
                      } else {
                        if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266") {
                          increaseFound = true;
                        }
                      }
                    } else if (item.id != null && item.id > 0 && (item.slot.toString() == "12" || item.slot.toString() == "13")) {
                      var itemName = "";
                      if (item.name != null)
                        itemName = item.name.toString();
                      else
                        itemName = getStringForLang("notFound", langKeys, langTrans, "", "", "", "");
                      if (trinketsUsed.indexOf(itemName) < 0)
                        trinketsUsed += "1x " + itemName + "\r\n";
                      else {
                        var previousNumber = 1;
                        var temp = trinketsUsed.split("x " + itemName)[0];
                        if (temp.length == 1)
                          previousNumber = Number(temp);
                        else {
                          var tempArr = temp.split("\r\n");
                          previousNumber = Number(tempArr[tempArr.length - 1]);
                        }
                        var stringOld = previousNumber.toString() + "x " + itemName;
                        previousNumber++;
                        var stringNew = previousNumber.toString() + "x " + itemName;
                        trinketsUsed = trinketsUsed.replace(stringOld, stringNew);
                      }
                    }

                    if (item.id != null && item.id.toString().length > 0 && item.id.toString() != "0" && item.slot != 3 && item.slot != 18) {
                      fightGearSpellHaste += Number(searchEntryForId(spellHasteInfoIds, spellHasteInfoValues, item.id.toString()));
                      if (item.gems != null) {
                        item.gems.forEach(function (gem, gemCount) {
                          if (gem.id.toString() == "35761") {
                            fightGearSpellHaste += 10;
                          } else if (gem.id.toString() == "35315") {
                            fightGearSpellHaste += 8;
                          } else if (gem.id.toString() == "35760") {
                            fightGearSpellHaste += 5;
                          } else if (gem.id.toString() == "35316") {
                            fightGearSpellHaste += 4;
                          } else if (gem.id.toString() == "35759") {
                            fightGearSpellHaste += 5;
                          } else if (gem.id.toString() == "35318") {
                            fightGearSpellHaste += 4;
                          }
                        })
                      }
                    }
                  })
                }
              }
            })
            if (fightGearSpellHaste > gearSpellHaste)
              gearSpellHaste = fightGearSpellHaste;
          }
          if (increaseTotal == false && bossSummaryData.playerDetails != null && bossSummaryData.playerDetails.healers != null && bossSummaryData.playerDetails.healers.length > 0) {
            var fightGearSpellHaste = 0;
            bossSummaryData.playerDetails.healers.forEach(function (playerInfo, playerInfoCount) {
              if (playerInfo.name == playerByNameAsc.name) {
                if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                  increaseTotal = true;
                  playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                    if (item.slot.toString() == "15" || item.slot.toString() == "16") {
                      if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                        increaseTotal = false;
                      } else {
                        if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266") {
                          increaseFound = true;
                        }
                      }
                    } else if (item.id != null && item.id > 0 && (item.slot.toString() == "12" || item.slot.toString() == "13")) {
                      var itemName = "";
                      if (item.name != null)
                        itemName = item.name.toString();
                      else
                        itemName = getStringForLang("notFound", langKeys, langTrans, "", "", "", "");
                      if (trinketsUsed.indexOf(itemName) < 0)
                        trinketsUsed += "1x " + itemName + "\r\n";
                      else {
                        var previousNumber = 1;
                        var temp = trinketsUsed.split("x " + itemName)[0];
                        if (temp.length == 1)
                          previousNumber = Number(temp);
                        else {
                          var tempArr = temp.split("\r\n");
                          previousNumber = Number(tempArr[tempArr.length - 1]);
                        }
                        var stringOld = previousNumber.toString() + "x " + itemName;
                        previousNumber++;
                        var stringNew = previousNumber.toString() + "x " + itemName;
                        trinketsUsed = trinketsUsed.replace(stringOld, stringNew);
                      }
                    }

                    if (item.id != null && item.id.toString().length > 0 && item.id.toString() != "0" && item.slot != 3 && item.slot != 18) {
                      fightGearSpellHaste += Number(searchEntryForId(spellHasteInfoIds, spellHasteInfoValues, item.id.toString()));
                      if (item.gems != null) {
                        item.gems.forEach(function (gem, gemCount) {
                          if (gem.id.toString() == "35761") {
                            fightGearSpellHaste += 10;
                          } else if (gem.id.toString() == "35315") {
                            fightGearSpellHaste += 8;
                          } else if (gem.id.toString() == "35760") {
                            fightGearSpellHaste += 5;
                          } else if (gem.id.toString() == "35316") {
                            fightGearSpellHaste += 4;
                          } else if (gem.id.toString() == "35759") {
                            fightGearSpellHaste += 5;
                          } else if (gem.id.toString() == "35318") {
                            fightGearSpellHaste += 4;
                          }
                        })
                      }
                    }
                  })
                }
              }
            })
            if (fightGearSpellHaste > gearSpellHaste)
              gearSpellHaste = fightGearSpellHaste;
          }
          if (increaseTotal == false && bossSummaryData.playerDetails != null && bossSummaryData.playerDetails.tanks != null && bossSummaryData.playerDetails.tanks.length > 0) {
            var fightGearSpellHaste = 0;
            bossSummaryData.playerDetails.tanks.forEach(function (playerInfo, playerInfoCount) {
              if (playerInfo.name == playerByNameAsc.name) {
                if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                  increaseTotal = true;
                  playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                    if (item.slot.toString() == "15" || item.slot.toString() == "16") {
                      if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                        increaseTotal = false;
                      } else {
                        if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266") {
                          increaseFound = true;
                        }
                      }
                    } else if (item.id != null && item.id > 0 && (item.slot.toString() == "12" || item.slot.toString() == "13")) {
                      var itemName = "";
                      if (item.name != null)
                        itemName = item.name.toString();
                      else
                        itemName = getStringForLang("notFound", langKeys, langTrans, "", "", "", "");
                      if (trinketsUsed.indexOf(itemName) < 0)
                        trinketsUsed += "1x " + itemName + "\r\n";
                      else {
                        var previousNumber = 1;
                        var temp = trinketsUsed.split("x " + itemName)[0];
                        if (temp.length == 1)
                          previousNumber = Number(temp);
                        else {
                          var tempArr = temp.split("\r\n");
                          previousNumber = Number(tempArr[tempArr.length - 1]);
                        }
                        var stringOld = previousNumber.toString() + "x " + itemName;
                        previousNumber++;
                        var stringNew = previousNumber.toString() + "x " + itemName;
                        trinketsUsed = trinketsUsed.replace(stringOld, stringNew);
                      }
                    }

                    if (item.id != null && item.id.toString().length > 0 && item.id.toString() != "0" && item.slot != 3 && item.slot != 18) {
                      fightGearSpellHaste += Number(searchEntryForId(spellHasteInfoIds, spellHasteInfoValues, item.id.toString()));
                      if (item.gems != null) {
                        item.gems.forEach(function (gem, gemCount) {
                          if (gem.id.toString() == "35761") {
                            fightGearSpellHaste += 10;
                          } else if (gem.id.toString() == "35315") {
                            fightGearSpellHaste += 8;
                          } else if (gem.id.toString() == "35760") {
                            fightGearSpellHaste += 5;
                          } else if (gem.id.toString() == "35316") {
                            fightGearSpellHaste += 4;
                          } else if (gem.id.toString() == "35759") {
                            fightGearSpellHaste += 5;
                          } else if (gem.id.toString() == "35318") {
                            fightGearSpellHaste += 4;
                          }
                        })
                      }
                    }
                  })
                }
              }
            })
            if (fightGearSpellHaste > gearSpellHaste)
              gearSpellHaste = fightGearSpellHaste;
          }
          if (increaseFound)
            found++;
          else if (playerByNameAsc.type == "Rogue" && increaseTotal) {
            bossDamageDataAll[total].entries.forEach(function (bossDamageData, bossDamageDataCount) {
              if (bossDamageData != null && bossDamageData.id != null && bossDamageData.id == playerByNameAsc.id) {
                found++;
              }
            })
          }
          if (increaseTotal)
            total++;
        })
        usedTemporaryWeaponEnchant = Math.round(found * 100 / total) + "%";

        //fill in single target casts
        if (singleTargetCastsToTrack != null) {
          var singleTargetTotalTime = 0;
          singleTargetCastsToTrack.forEach(function (singleTargetCast, singleTargetCastCount) {
            var amount = 0;
            var amountOmitted = 0;
            var uptime = 0;
            var overheal = 0;
            var debuffIdString = "";
            var lowerRankUsed = 0;
            if (singleTargetCast.indexOf("(Fire Vuln") > -1)
              debuffIdString = "22959";
            else if (singleTargetCast.indexOf("(Shadow Vuln") > -1)
              debuffIdString = "15258";
            else if (singleTargetCast.indexOf("(WC") > -1)
              debuffIdString = "12579";
            else if (singleTargetCast.indexOf("(Flurry") > -1)
              debuffIdString = "12970";
            else if (singleTargetCast.indexOf("(Expose Weakness") > -1)
              debuffIdString = "34501";
            else if (singleTargetCast.indexOf("(Deep Wounds") > -1)
              debuffIdString = "12721";
            else
              singleTargetCast.split("[")[1].split("]")[0].split(",").forEach(function (stCast, stCastCount) {
                if (stCast.indexOf("*") < 0)
                  debuffIdString = stCast;
              })
            var singleTargetCastString = "";
            if (previousClass != playerByNameAsc.type) {
              if (singleTargetCastCount == 0) {
                copyRowStyles(conf, sheet, confSingleTargetCastsToTrack, singleTargetCastsToTrack.length, singleTargetCasts.getRow() + 1, singleTargetCasts.getColumn() + playerDoneCount + classDoneCount, playersInThisClass, false, "left", darkMode);
                var confColumnWidth = conf.getColumnWidth(confSingleTargetCastsToTrack.getColumn());
                if (confColumnWidth > maxColumnWidth) {
                  maxColumnWidth = confColumnWidth;
                }
              }
              singleTargetCastString = singleTargetCastsToTrackLang[singleTargetCastCount].split(" [")[0].split(" {")[0];
              if (singleTargetCast.indexOf("Slice and Dice") < 0 && singleTargetCast.indexOf("Battle Shout") < 0 && singleTargetCast.indexOf("Commanding Shout") < 0 && singleTargetCast.indexOf("Earth Shield") < 0 && singleTargetCast.indexOf("Water Shield") < 0 && singleTargetCast.indexOf("Judgement") < 0 && singleTargetCast.indexOf("Misdirect") < 0) {
                if (!onlyTrash && singleTargetCast.indexOf("uptime") > -1 && singleTargetCast.indexOf("total") < 0 && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                  singleTargetCastString = singleTargetCastString.replace("%)", "% - " + getStringForLang("overall", langKeys, langTrans, "", "", "", "") + ": " + getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedDataBossesTotal, totalTimeElapsedBosses) + "%)");
                } else if (singleTargetCast.indexOf("uptime") > -1) {
                  if (debuffIdString != "12970")
                    singleTargetCastString = singleTargetCastString.replace("%)", "% - " + getStringForLang("overall", langKeys, langTrans, "", "", "", "") + ": " + getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedDataTotal, totalTimeElapsedRaw) + "%)");
                }
              } else if (singleTargetCast.indexOf("Judgement") > -1)
                singleTargetCastString += " (" + getStringForLang("uptimeNoPerc", langKeys, langTrans, "", "", "", "") + "%)"
              else if (singleTargetCast.indexOf("Misdirect") > -1)
                singleTargetCastString += " (" + getStringForLang("includingOOC", langKeys, langTrans, "", "", "", "") + ")";
              singleTargetCastsArr[singleTargetCastCount + 1][playerDoneCount + classDoneCount] = singleTargetCastString;
            }
            if (singleTargetCast.indexOf("Melee") > -1 || singleTargetCast.indexOf("Execute") > -1 || singleTargetCast.indexOf("Shiv") > -1 || singleTargetCast.indexOf("Envenom") > -1 || singleTargetCast.indexOf("Auto Shot") > -1) {
              if (singleTargetCast.indexOf("Expose Weakness") > -1 || singleTargetCast.indexOf("Deep Wounds") > -1) {
                uptime = getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedData, totalTimeElapsedRaw);
              }
              damageDoneData.entries.forEach(function (damDone, damDoneCount) {
                if (singleTargetCast.indexOf("[") > -1) {
                  singleTargetCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                    if (spellId.replace("*", "") == damDone.guid.toString()) {
                      if (singleTargetCast.indexOf("Melee") > -1 || singleTargetCast.indexOf("Auto Shot") > -1) {
                        if (damDone.uses && damDone.uses > 0)
                          amount += damDone.uses;
                      } else {
                        if (damDone.hitCount != null && damDone.hitCount > 0) {
                          amount += damDone.hitCount;
                          if (spellId.indexOf("*") > -1)
                            lowerRankUsed += damDone.hitCount;
                        }
                        if (damDone.missCount != null && damDone.missCount > 0) {
                          amount += damDone.missCount;
                          if (spellId.indexOf("*") > -1)
                            lowerRankUsed += damDone.missCount;
                        }
                      }
                    }
                  })
                }
              })
            } else if (singleTargetCast.indexOf("Misdirect") > -1) {
              buffsData.auras.forEach(function (buff, buffCount) {
                if (singleTargetCast.indexOf("[") > -1) {
                  singleTargetCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                    if (buff.guid.toString() == spellId.replace("*", "") && buff.totalUses > 0) {
                      amount += buff.totalUses;
                    }
                  })
                }
              })
            } else if (singleTargetCast.indexOf("[99999]") > -1) {
              playerDataSunderArmorOnLessThan5Stacks.entries.forEach(function (spellSA5, spellSA5Count) {
                if (spellSA5.id == playerByNameAsc.id) {
                  amount = spellSA5.total;
                  var saCastsOverall = 0;
                  playerData.entries.forEach(function (spell, spellCount) {
                    if (spell.guid.toString() == "11597" || spell.guid.toString() == "25225")
                      saCastsOverall = spell.total;
                  })
                  uptime = Math.round(spellSA5.total * 100 / saCastsOverall) * spellSA5.total;
                }
              })
            } else if (singleTargetCast.indexOf("[99998]") > -1) {
              playerDataScorchOnLessThan5Stacks.entries.forEach(function (spellScorch5, spellScorch5Count) {
                if (spellScorch5.id == playerByNameAsc.id) {
                  amount = spellScorch5.total;
                  var scorchCastsOverall = 0;
                  playerData.entries.forEach(function (spell, spellCount) {
                    if (spell.guid.toString() == "27074" || spell.guid.toString() == "27073" || spell.guid.toString() == "10207" || spell.guid.toString() == "10206" || spell.guid.toString() == "10205" || spell.guid.toString() == "8446" || spell.guid.toString() == "8445" || spell.guid.toString() == "8444" || spell.guid.toString() == "2948")
                      scorchCastsOverall += spell.total;
                  })
                  uptime = Math.round(spellScorch5.total * 100 / scorchCastsOverall) * spellScorch5.total;
                }
              })
            } else {
              playerData.entries.forEach(function (spell, spellCount) {
                if (singleTargetCast.indexOf("[") > -1) {
                  singleTargetCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                    if (spellId.indexOf("*") > 0 && spellId.substring(0, spellId.length - 1) == spell.guid.toString() && spell.total > 0) {
                      lowerRankUsed += spell.total;
                    }
                    spellId = spellId.replace("*", "");
                    if (singleTargetCast.indexOf("Hamstring") > -1) {
                      uptime = getUptimeForDebuffSpellId(debuffIdString, buffsData, totalTimeElapsedRaw);
                      if (uptime == 0)
                        uptime = getUptimeForDebuffSpellId("12969", buffsData, totalTimeElapsedRaw);
                      if (uptime == 0)
                        uptime = getUptimeForDebuffSpellId("12968", buffsData, totalTimeElapsedRaw);
                      if (uptime == 0)
                        uptime = getUptimeForDebuffSpellId("12967", buffsData, totalTimeElapsedRaw);
                      if (uptime == 0)
                        uptime = getUptimeForDebuffSpellId("12966", buffsData, totalTimeElapsedRaw);
                    } else if (singleTargetCast.indexOf("Shadow Vuln") > -1) {
                      uptime = getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedData, totalTimeElapsedRaw);
                    } else if (singleTargetCast.indexOf("efficiency") > -1) {
                      try {
                        if (VTManaGainData != null && VTManaGainData.resources != null && VTManaGainData.resources.length > 0) {
                          var rawManaGain = VTManaGainData.resources[0].wasted + VTManaGainData.resources[0].gains;
                          var shadowDamageTotal = shadowDamageDoneData.entries[0].total;
                          uptime = Math.round(rawManaGain * 2000 / shadowDamageTotal);
                        }
                      } catch { }
                    }
                    if (spellId == spell.guid.toString()) {
                      amount += spell.total;
                      if (!onlyTrash && singleTargetCast.indexOf("uptime") > -1 && singleTargetCast.indexOf("total") < 0 && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                        if (singleTargetCast.indexOf("Slice and Dice") > -1) {
                          playerDataBosses.entries.forEach(function (spellOnBoss, spellOnBossCount) {
                            if (spellOnBoss.guid == spell.guid)
                              uptime += Math.round(spellOnBoss.uptime * 100 / totalTimeElapsedBosses) * spell.total;
                          })
                        } else
                          uptime += getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedDataBosses, totalTimeElapsedBosses) * spell.total;
                      } else if (singleTargetCast.indexOf("Judgement") > -1) {
                        debuffsAppliedDataBossesJudgement.auras.forEach(function (debuffTotal, debuffTotalCount) {
                          uptime += Math.round(debuffTotal.totalUptime * 100 / totalTimeElapsedBosses);
                        })
                      } else if (singleTargetCast.indexOf("uptime") > -1) {
                        if (singleTargetCast.indexOf("Slice and Dice") > -1 || singleTargetCast.indexOf("Battle Shout") > -1 || singleTargetCast.indexOf("Commanding Shout") > -1 || singleTargetCast.indexOf("Earth Shield") > -1 || singleTargetCast.indexOf("Water Shield") > -1) {
                          uptime += Math.round(spell.uptime * 100 / totalTimeElapsedRaw) * spell.total;
                        } else if (singleTargetCast.indexOf("Hamstring") > -1 || singleTargetCast.indexOf("Shadow Vuln") > -1) {
                          uptime = uptime;
                        } else {
                          uptime += getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedData, totalTimeElapsedRaw) * spell.total;
                        }
                      }

                      if (singleTargetCast.indexOf("(Fire Vuln") > -1 && uptime > 0)
                        usedFireVulnerabilityOrBlastWave = true;

                      if (singleTargetCast.indexOf("overheal") > -1) {
                        healingData.entries.forEach(function (heal, healCount) {
                          if (spell.guid == heal.guid) {
                            var healed = 0;
                            if (heal.hitdetails != null) {
                              heal.hitdetails.forEach(function (hitData, hitDataCount) {
                                if (hitData.absorbOrOverheal != null)
                                  healed += hitData.absorbOrOverheal;
                              })
                            }
                            if (heal.overheal != null && heal.overheal > 0)
                              overheal += (heal.overheal * 100 / (heal.total + healed)) * spell.total;
                            else
                              amountOmitted += spell.total;
                          }
                        })
                      }
                    }
                  })
                }
              })
            }
            if (singleTargetCast.indexOf("twisted swings") > -1 && !onlyTrash) {
              var totalWindfuryAttacksOnTwistsDoneOnBosses = 0;
              var totalWindfuryAttacksOnBosses = 0;
              if (WindfuryAttacksOnTwistsDoneOnBosses != null && WindfuryAttacksOnTwistsDoneOnBosses.entries != null && WindfuryAttacksOnTwistsDoneOnBosses.entries.length == 1 && WindfuryAttacksOnTwistsDoneOnBosses.entries[0].hitCount != null && WindfuryAttacksOnTwistsDoneOnBosses.entries[0].hitCount >= 0) {
                totalWindfuryAttacksOnTwistsDoneOnBosses += WindfuryAttacksOnTwistsDoneOnBosses.entries[0].hitCount;
                totalWindfuryAttacksOnTwistsDoneOnBosses += WindfuryAttacksOnTwistsDoneOnBosses.entries[0].missCount;
              }
              if (WindfuryAttacksOnTwistsDoneOnBossesRank1 != null && WindfuryAttacksOnTwistsDoneOnBossesRank1.entries != null && WindfuryAttacksOnTwistsDoneOnBossesRank1.entries.length == 1 && WindfuryAttacksOnTwistsDoneOnBossesRank1.entries[0].hitCount != null && WindfuryAttacksOnTwistsDoneOnBossesRank1.entries[0].hitCount >= 0) {
                totalWindfuryAttacksOnTwistsDoneOnBosses += WindfuryAttacksOnTwistsDoneOnBossesRank1.entries[0].hitCount;
                totalWindfuryAttacksOnTwistsDoneOnBosses += WindfuryAttacksOnTwistsDoneOnBossesRank1.entries[0].missCount;
              }
              if (WindfuryAttacksOnBosses != null && WindfuryAttacksOnBosses.auras != null && WindfuryAttacksOnBosses.auras.length == 1 && WindfuryAttacksOnBosses.auras[0].totalUses != null && WindfuryAttacksOnBosses.auras[0].totalUses >= 0) {
                totalWindfuryAttacksOnBosses += WindfuryAttacksOnBosses.auras[0].totalUses;
              }
              if (WindfuryAttacksOnBossesRank1 != null && WindfuryAttacksOnBossesRank1.auras != null && WindfuryAttacksOnBossesRank1.auras.length == 1 && WindfuryAttacksOnBossesRank1.auras[0].totalUses != null && WindfuryAttacksOnBossesRank1.auras[0].totalUses >= 0) {
                totalWindfuryAttacksOnBosses += WindfuryAttacksOnBossesRank1.auras[0].totalUses;
              }
              if (TwistsDoneOnBosses != null && TwistsDoneOnBosses.entries != null && TwistsDoneOnBosses.entries.length == 1 && TwistsDoneOnBosses.entries[0].hitCount != null && TwistsDoneOnBosses.entries[0].hitCount >= 0 && DamageDoneOnBosses != null && DamageDoneOnBosses.entries != null && DamageDoneOnBosses.entries.length == 1 && DamageDoneOnBosses.entries[0].hitCount != null && DamageDoneOnBosses.entries[0].hitCount >= 0 && totalWindfuryAttacksOnTwistsDoneOnBosses > 0 && totalWindfuryAttacksOnBosses > 0) {
                uptime = Math.round((((TwistsDoneOnBosses.entries[0].hitCount + TwistsDoneOnBosses.entries[0].missCount) - ((totalWindfuryAttacksOnTwistsDoneOnBosses) / 2)) * 100) / ((DamageDoneOnBosses.entries[0].hitCount + DamageDoneOnBosses.entries[0].missCount) - totalWindfuryAttacksOnBosses));
              }
            }
            if (amount > 0 && Math.round(lowerRankUsed * 100 / amount) > 50) {
              sheet.getRange(singleTargetCasts.getRow() + singleTargetCastCount + 1, singleTargetCasts.getColumn() + playerDoneCount + classDoneCount + 1, 1, 1).setFontWeight("bold").setFontStyle("italic").setFontColor("#980000");
            }
            if (amount == 0 && !(singleTargetCast.indexOf("uptime") > -1 && uptime > 0)) {
              if (singleTargetCast.length > 0)
                singleTargetCastsArr[singleTargetCastCount + 1][playerDoneCount + classDoneCount + 1] = "0";
            }
            else {
              if (singleTargetCast.indexOf("uptime") > -1 || singleTargetCast.indexOf("Judgement") > -1 || singleTargetCast.indexOf("[99999]") > -1 || singleTargetCast.indexOf("[99998]") > -1 || singleTargetCast.indexOf("efficiency") > -1 || singleTargetCast.indexOf("twisted swings") > -1) {
                if (singleTargetCast.indexOf("Hamstring") > -1 || singleTargetCast.indexOf("Judgement") > -1 || singleTargetCast.indexOf("Shadow Vuln") > -1 || singleTargetCast.indexOf("Expose Weakness") > -1 || singleTargetCast.indexOf("Deep Wounds") > -1 || singleTargetCast.indexOf("efficiency") > -1 || singleTargetCast.indexOf("twisted swings") > -1) {
                  singleTargetCastsArr[singleTargetCastCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + uptime + "%)";
                } else {
                  singleTargetCastsArr[singleTargetCastCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + Math.round(uptime / amount) + "%)";
                }
              }
              else if (singleTargetCast.indexOf("overheal") > -1) {
                if ((amount - amountOmitted) > 0)
                  singleTargetCastsArr[singleTargetCastCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + Math.round(overheal / (amount - amountOmitted)) + "%)";
                else
                  singleTargetCastsArr[singleTargetCastCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (0%)";
              } else
                singleTargetCastsArr[singleTargetCastCount + 1][playerDoneCount + classDoneCount + 1] = amount;
              if (singleTargetCast.indexOf("Fireball") > -1 && usedFireVulnerabilityOrBlastWave == false) {
                singleTargetTotalTime = singleTargetTotalTime + amount * Number(3.5);
              } else if (singleTargetCast.indexOf("Frostbolt (rank 2+)") > -1 && usedIceblockOrIceBarrier == true) {
                singleTargetTotalTime = singleTargetTotalTime + amount * Number(2.5);
              } else
                singleTargetTotalTime = singleTargetTotalTime + amount * Number(singleTargetCast.split("{")[1].split("}")[0]);
            }
          })
        }

        //fill in aoe casts
        var aoeTotalTime = 0;
        if (aoeCastsToTrack != null) {
          var aoeCastsDone = 0;
          var totalHits = 0;
          var totalCasts = 0;
          var doneEntries = [];
          aoeCastsToTrack.forEach(function (aoeCast, aoeCastCount) {
            var amount = 0;
            var amountOmitted = 0;
            var overheal = 0;
            var lowerRankUsed = 0;
            var hitsThisSpell = 0;
            var castsThisSpell = 0;
            if (previousClass != playerByNameAsc.type) {
              if (aoeCastCount == 0) {
                var confColumnWidth = conf.getColumnWidth(confAoeCastsToTrack.getColumn());
                if (confColumnWidth > maxColumnWidth) {
                  maxColumnWidth = confColumnWidth;
                }
                copyRowStyles(conf, sheet, confAoeCastsToTrack, aoeCastsToTrack.length, aoeCasts.getRow() + 1, aoeCasts.getColumn() + playerDoneCount + classDoneCount, playersInThisClass, false, "left", darkMode);
              }
              if (aoeCast.indexOf("Cleave") > -1 || aoeCast.indexOf("Whirlwind") > -1)
                aoeCastsArr[aoeCastCount + 1][playerDoneCount + classDoneCount] = aoeCastsToTrackLang[aoeCastCount].split(" [")[0].split(" {")[0] + " (⌀)";
              else
                aoeCastsArr[aoeCastCount + 1][playerDoneCount + classDoneCount] = aoeCastsToTrackLang[aoeCastCount].split(" [")[0].split(" {")[0];
            }
            playerData.entries.forEach(function (spell, spellCount) {
              if (aoeCast.indexOf("[") > -1) {
                aoeCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                  if (spellId.indexOf("*") > 0 && spellId.substring(0, spellId.length - 1) == spell.guid.toString() && spell.total > 0) {
                    lowerRankUsed += spell.total;
                  }
                  spellId = spellId.replace("*", "");
                  if (spellId == spell.guid.toString()) {
                    amount += spell.total;

                    var dataToSearch = damageDoneData;
                    if (aoeCast.indexOf("overheal") > -1)
                      dataToSearch = healingData;

                    dataToSearch.entries.forEach(function (damageSpell, damageSpellCount) {
                      if (damageSpell.guid == spell.guid || (spell.guid == 13877 && damageSpell.guid == 22482)) {
                        if (!doneEntries.includes(playerByNameAsc.name + ": " + damageSpell.guid + " - " + damageSpell.hitCount + " - " + damageSpell.uses)) {
                          doneEntries.push(playerByNameAsc.name + ": " + damageSpell.guid + " - " + damageSpell.hitCount + " - " + damageSpell.uses);
                          totalHits += damageSpell.hitCount;
                          if (damageSpell.missCount != null && damageSpell.missCount > 0)
                            totalHits += damageSpell.missCount;
                          totalCasts += spell.total;
                          hitsThisSpell += damageSpell.hitCount;
                          if (damageSpell.missCount != null && damageSpell.missCount > 0)
                            hitsThisSpell += damageSpell.missCount;
                          castsThisSpell += spell.total;
                        }
                      }
                    })

                    if (aoeCast.indexOf("overheal") > -1) {
                      var spellIdHeal = spell.guid;
                      if (spell.guid == 15237)
                        spellIdHeal = 23455;
                      else if (spell.guid == 15430)
                        spellIdHeal = 23458;
                      else if (spell.guid == 15431)
                        spellIdHeal = 23459;
                      else if (spell.guid == 27799)
                        spellIdHeal = 27803;
                      else if (spell.guid == 27800)
                        spellIdHeal = 27804;
                      else if (spell.guid == 27801)
                        spellIdHeal = 27805;
                      else if (spell.guid == 25331)
                        spellIdHeal = 25329;
                      healingData.entries.forEach(function (heal, healCount) {
                        if (spellIdHeal == heal.guid) {
                          var healed = 0;
                          if (heal.hitdetails != null) {
                            heal.hitdetails.forEach(function (hitData, hitDataCount) {
                              if (hitData.absorbOrOverheal != null)
                                healed += hitData.absorbOrOverheal;
                            })
                          }
                          if (heal.overheal != null && heal.overheal > 0)
                            overheal += (heal.overheal * 100 / (heal.total + healed)) * spell.total;
                          else
                            amountOmitted += spell.total;
                        }
                      })
                    }
                  }
                })
              }
            })
            if (amount > 0 && Math.round(lowerRankUsed * 100 / amount) > 50) {
              sheet.getRange(aoeCasts.getRow() + aoeCastCount + 1, aoeCasts.getColumn() + playerDoneCount + classDoneCount + 1, 1, 1).setFontWeight("bold").setFontStyle("italic").setFontColor("#980000");
            }
            if (amount == 0) {
              if (aoeCast.length > 0)
                aoeCastsArr[aoeCastCount + 1][playerDoneCount + classDoneCount + 1] = "0";
            }
            else {
              if (aoeCast.indexOf("overheal") > -1)
                aoeCastsArr[aoeCastCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + Math.round(overheal / (amount - amountOmitted)) + "%)";
              else if (aoeCast.indexOf("Cleave") > -1 || aoeCast.indexOf("Whirlwind") > -1)
                aoeCastsArr[aoeCastCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + Math.round(hitsThisSpell * 100 / castsThisSpell) / 100 + ")";
              else
                aoeCastsArr[aoeCastCount + 1][playerDoneCount + classDoneCount + 1] = amount;
              aoeTotalTime = aoeTotalTime + amount * Number(aoeCast.split("{")[1].split("}")[0]);
            }
            aoeCastsDone++;
          })
          //hit count of pulsing spells can't be retrieved
          if (showAverageOfHitsPerAoeCast && playerByNameAsc.type != "Druid" && playerByNameAsc.type != "Paladin" && playerByNameAsc.type != "Warlock") {
            if (previousClass != playerByNameAsc.type) {
              aoeCastsArr[aoeCastsDone + 1][playerDoneCount + classDoneCount] = getStringForLang("hitsPerAoeCast", langKeys, langTrans, "", "", "", "");
              copyRangeStyle(confShowAverageOfHitsPerAoeCast, sheet.getRange(aoeCasts.getRow() + aoeCastsDone + 1, aoeCasts.getColumn() + playerDoneCount + classDoneCount, 1, 1), null, "left", null);
              if (darkMode)
                sheet.getRange(aoeCasts.getRow() + aoeCastsDone + 1, aoeCasts.getColumn() + playerDoneCount + classDoneCount, 1, 1).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
              copyRangeStyle(confShowAverageOfHitsPerAoeCast, sheet.getRange(aoeCasts.getRow() + aoeCastsDone + 1, aoeCasts.getColumn() + playerDoneCount + classDoneCount + 1, 1, playersInThisClass), null, "center", null);
              if (darkMode)
                sheet.getRange(aoeCasts.getRow() + aoeCastsDone + 1, aoeCasts.getColumn() + playerDoneCount + classDoneCount + 1, 1, playersInThisClass).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
            }
            if (totalCasts == 0)
              aoeCastsArr[aoeCastsDone + 1][playerDoneCount + classDoneCount + 1] = 0;
            else
              aoeCastsArr[aoeCastsDone + 1][playerDoneCount + classDoneCount + 1] = "=" + totalHits + "/" + totalCasts;
          }
        }

        //fill in seconds active
        if (secondsActive != null) {
          if (secondsActiveArr.length == 1) {
            addSingleEntryToMultiDimArray(secondsActiveArr, getStringForLang("secondsActiveST", langKeys, langTrans, "", "", "", "")); addSingleEntryToMultiDimArray(secondsActiveArr, getStringForLang("relativeActiveST", langKeys, langTrans, "", "", "", "")); addSingleEntryToMultiDimArray(secondsActiveArr, getStringForLang("relativeActiveTotal", langKeys, langTrans, "", "", "", ""));
            if (darkMode)
              sheet.getRange(secondsActive.getRow(), secondsActive.getColumn(), 1, maxColumns).setFontColor("#d9d9d9");
            else
              sheet.getRange(secondsActive.getRow(), secondsActive.getColumn(), 1, maxColumns).setFontColor("white");
            copyRangeStyle(confTotalAndInformationRowsDefaultTemplate, sheet.getRange(secondsActive.getRow() + 1, secondsActive.getColumn(), 1, maxColumns), null, "center", null);
            sheet.getRange(secondsActive.getRow() + 1, secondsActive.getColumn(), 1, maxColumns).setFontSize(confTotalAndInformationRowsDefaultTemplate.getFontSize()).setFontStyle(confTotalAndInformationRowsDefaultTemplate.getFontStyle()).setFontWeight(confTotalAndInformationRowsDefaultTemplate.getFontWeight()).setHorizontalAlignment("center");
            sheet.getRange(secondsActive.getRow() + 2, secondsActive.getColumn(), 1, maxColumns).setFontSize(confTotalAndInformationRowsDefaultTemplate.getFontSize()).setFontStyle(confTotalAndInformationRowsDefaultTemplate.getFontStyle()).setFontWeight(confTotalAndInformationRowsDefaultTemplate.getFontWeight()).setHorizontalAlignment("center");
            sheet.getRange(secondsActive.getRow() + 3, secondsActive.getColumn(), 1, maxColumns).setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID).setBackground("#b7b7b7");
            sheet.getRange(secondsActive.getRow() + 2, secondsActive.getColumn() + 1, 2, maxColumns).setNumberFormat("0%");
            addSingleEntryToMultiDimArray(secondsActiveArr, getStringForLang("relativeActiveAoe", langKeys, langTrans, "", "", "", "")); addSingleEntryToMultiDimArray(secondsActiveArr, getStringForLang("secondsActiveAoe", langKeys, langTrans, "", "", "", ""));
            copyRangeStyle(confTotalAndInformationRowsDefaultTemplate, sheet.getRange(secondsActive.getRow() + 5, secondsActive.getColumn(), 1, maxColumns), null, "center", null);
            sheet.getRange(secondsActive.getRow() + 4, secondsActive.getColumn() + 1, 1, maxColumns).setNumberFormat("0%");
            sheet.getRange(secondsActive.getRow() + 4, secondsActive.getColumn(), 2, maxColumns).setFontSize(confTotalAndInformationRowsDefaultTemplate.getFontSize()).setFontStyle(confTotalAndInformationRowsDefaultTemplate.getFontStyle()).setFontWeight(confTotalAndInformationRowsDefaultTemplate.getFontWeight()).setHorizontalAlignment("center");
            if (showWCLActivePercentage) {
              if (activePercentageTotalOnTrash > 0 && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                addSingleEntryToMultiDimArray(secondsActiveArr, getStringForLang("WCLactiveNoTrash", langKeys, langTrans, "", "", "", ""));
              else
                addSingleEntryToMultiDimArray(secondsActiveArr, getStringForLang("WCLactiveTrash", langKeys, langTrans, "", "", "", ""));
              addSingleEntryToMultiDimArray(secondsActiveArr, "");
              sheet.getRange(secondsActive.getRow() + 6, secondsActive.getColumn()).setHorizontalAlignment("center").setFontStyle("italic");
              sheet.getRange(secondsActive.getRow() + 6, secondsActive.getColumn() + 1, 1, maxColumns).setNumberFormat("0%").setFontStyle("italic").setHorizontalAlignment("center");
            }
          }
          if (previousClass != playerByNameAsc.type) {
            secondsActiveArr[0].push("");
            secondsActiveArr[1].push("");
            secondsActiveArr[2].push("");
            secondsActiveArr[3].push("");
            secondsActiveArr[4].push("");
            secondsActiveArr[5].push("");
            if (showWCLActivePercentage)
              secondsActiveArr[6].push("");
          }
          var secondsToSubstractFromHasteBuffsOnTrash = 0;
          var secondsToSubstractFromHasteBuffsOnBosses = 0;
          secondsToSubstractFromHasteBuffsOnBosses += BloodlustOnBosses * 9;
          secondsToSubstractFromHasteBuffsOnTrash += BloodlustOnTrash * 9;
          secondsToSubstractFromHasteBuffsOnBosses += IcyVeinsOnBosses * 3;
          secondsToSubstractFromHasteBuffsOnTrash += IcyVeinsOnTrash * 3;
          secondsToSubstractFromHasteBuffsOnBosses += SkullUsesOnBosses * 1.5;
          secondsToSubstractFromHasteBuffsOnTrash += SkullUsesOnTrash * 1.5;
          secondsToSubstractFromHasteBuffsOnBosses += BerserkingUsesOnBosses * 0.5;
          secondsToSubstractFromHasteBuffsOnTrash += BerserkingUsesOnTrash * 0.5;
          secondsToSubstractFromHasteBuffsOnBosses += ScrollsOfBlindingLightUsesOnBosses * 3.1;
          secondsToSubstractFromHasteBuffsOnTrash += ScrollsOfBlindingLightUsesOnTrash * 3.1;
          secondsToSubstractFromHasteBuffsOnBosses += QuagmirransEyeUsesOnBosses * 1;
          secondsToSubstractFromHasteBuffsOnTrash += QuagmirransEyeUsesOnTrash * 1;
          secondsToSubstractFromHasteBuffsOnBosses += ScarabOfTheInfiniteCycleUsesOnBosses * 1;
          secondsToSubstractFromHasteBuffsOnTrash += ScarabOfTheInfiniteCycleUsesOnTrash * 1;
          secondsToSubstractFromHasteBuffsOnBosses += MysticalSkyfireDiamondUsesOnBosses * 0.6;
          secondsToSubstractFromHasteBuffsOnTrash += MysticalSkyfireDiamondUsesOnTrash * 0.6;
          secondsToSubstractFromHasteBuffsOnBosses += MindQuickeningGemUsesOnBosses * 4.8;
          secondsToSubstractFromHasteBuffsOnTrash += MindQuickeningGemUsesOnTrash * 4.8;
          secondsToSubstractFromHasteBuffsOnBosses += BladeOfWizardryUsesOnBosses * 0.9;
          secondsToSubstractFromHasteBuffsOnTrash += BladeOfWizardryUsesOnTrash * 0.9;
          if ((playerByNameAsc.type != "Hunter" && playerByNameAsc.type != "Rogue" && playerByNameAsc.type != "Warrior") || gearSpellHaste > 0) {
            sheet.getRange(secondsActive.getRow() + 1, secondsActive.getColumn() + playerDoneCount + classDoneCount + 1).setNote(getStringForLang("spellhaste", langKeys, langTrans, "", "", "", "") + ": " + gearSpellHaste.toString());
            sheet.getRange(secondsActive.getRow() + 5, secondsActive.getColumn() + playerDoneCount + classDoneCount + 1).setNote(getStringForLang("spellhaste", langKeys, langTrans, "", "", "", "") + ": " + gearSpellHaste.toString());
          }
          sheet.getRange(secondsActive.getRow() + 3, secondsActive.getColumn() + playerDoneCount + classDoneCount + 1).setNote((Math.round(secondsToSubstractFromHasteBuffsOnBosses + secondsToSubstractFromHasteBuffsOnTrash)).toString() + " " + getStringForLang("minusSecondsSpellHaste", langKeys, langTrans, "", "", "", ""));
          secondsActiveArr[0].push(Math.round((singleTargetTotalTime + aoeTotalTime - secondsToSubstractFromHasteBuffsOnBosses - secondsToSubstractFromHasteBuffsOnTrash) / (1 + ((gearSpellHaste / 15.77) / 100))));
          secondsActiveArr[1].push(Math.round((singleTargetTotalTime) / (1 + ((gearSpellHaste / 15.77) / 100))) - Math.round((secondsToSubstractFromHasteBuffsOnBosses + secondsToSubstractFromHasteBuffsOnTrash) * (singleTargetTotalTime / (singleTargetTotalTime + aoeTotalTime))));
          secondsActiveArr[2].push("=" + (Math.round((singleTargetTotalTime) / (1 + ((gearSpellHaste / 15.77) / 100))) - Math.round((secondsToSubstractFromHasteBuffsOnBosses + secondsToSubstractFromHasteBuffsOnTrash) * (singleTargetTotalTime / (singleTargetTotalTime + aoeTotalTime)))) + "/MAX(" + sheet.getRange(secondsActive.getRow(), secondsActive.getColumn(), 1, maxColumns).getA1Notation() + ";1)");
          secondsActiveArr[3].push("=" + sheet.getRange(secondsActive.getRow() + 2, secondsActive.getColumn() + playerDoneCount + classDoneCount + 1).getA1Notation() + "+" + sheet.getRange(secondsActive.getRow() + 4, secondsActive.getColumn() + playerDoneCount + classDoneCount + 1).getA1Notation());
          secondsActiveArr[4].push("=" + (Math.round((aoeTotalTime) / (1 + ((gearSpellHaste / 15.77) / 100))) - Math.round((secondsToSubstractFromHasteBuffsOnBosses + secondsToSubstractFromHasteBuffsOnTrash) * (aoeTotalTime / (singleTargetTotalTime + aoeTotalTime)))) + "/MAX(" + sheet.getRange(secondsActive.getRow(), secondsActive.getColumn(), 1, maxColumns).getA1Notation() + ";1)");
          secondsActiveArr[5].push(Math.round((aoeTotalTime) / (1 + ((gearSpellHaste / 15.77) / 100))) - Math.round((secondsToSubstractFromHasteBuffsOnBosses + secondsToSubstractFromHasteBuffsOnTrash) * (aoeTotalTime / (singleTargetTotalTime + aoeTotalTime))));
          if (showWCLActivePercentage) {
            if (activePercentageTotalOnTrash > 0 && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
              secondsActiveArr[6].push(activePercentageTotal + "% (" + activePercentageTotalOnTrash + "%)");
            else
              secondsActiveArr[6].push(activePercentageTotal + "%");
          }
        }

        //fill in class cooldowns
        if (classCooldownsToTrack != null) {
          var totalAmount = 0;
          var line = 0;
          var classCooldownsDone = 0;
          classCooldownsToTrack.forEach(function (classCooldown, classCooldownCount) {
            if (!(factionName == "Alliance" && classCooldown.indexOf("Bloodlust") > -1) && !(factionName == "Horde" && classCooldown.indexOf("Heroism") > -1)) {
              var cdMultiplesOfActive = 1;
              var cooldownInSeconds = 0;
              if (classCooldown.indexOf("--") > -1 && classCooldown.indexOf("++") > -1) {
                cooldownInSeconds = classCooldown.split("--")[1];
                cdMultiplesOfActive = cooldownInSeconds / classCooldown.split("++")[1];
              }
              if (previousClass != playerByNameAsc.type) {
                if (classCooldown.indexOf("Power Infusion") > -1 || classCooldown.indexOf("Innervate") > -1 || classCooldown.indexOf("Blessing of Protection") > -1 || classCooldown.indexOf("Heroism") > -1 || classCooldown.indexOf("Bloodlust") > -1) {
                  if ((classCooldown.indexOf("Power Infusion") > -1 && playerByNameAsc.type == "Priest") || (classCooldown.indexOf("Innervate") > -1 && playerByNameAsc.type == "Druid") || (classCooldown.indexOf("Blessing of Protection") > -1 && playerByNameAsc.type == "Paladin")) {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("usedOrGainedTrash", langKeys, langTrans, "", "", "", "");
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("usedOrGainedBosses", langKeys, langTrans, "", "", "", "");
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("total", langKeys, langTrans, "", "", "", "");
                  } else {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("gainedOnTrash", langKeys, langTrans, "", "", "", "");
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("gainedOnBosses", langKeys, langTrans, "", "", "", "");
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("total", langKeys, langTrans, "", "", "", "");
                  }
                } else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("trashOn", langKeys, langTrans, "", "", "", "");
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("bossesOn", langKeys, langTrans, "", "", "", "");
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount] = classCooldownsToTrackLang[classCooldownCount].split(" [")[0] + " " + getStringForLang("total", langKeys, langTrans, "", "", "", "");
                }
                var confColumnWidth = conf.getColumnWidth(confClassCooldownsToTrack.getColumn());
                if (confColumnWidth > maxColumnWidth) {
                  maxColumnWidth = confColumnWidth;
                }
                var confRange = conf.createTextFinder(classCooldown).findNext();
                var rangeHeader = sheet.getRange(classCooldowns.getRow() + 1 + (classCooldownsDone * 3), classCooldowns.getColumn() + playerDoneCount + classDoneCount, 2, 1);
                sheet.setColumnWidth(rangeHeader.getColumn(), maxColumnWidth);
                copyRangeStyle(confRange, rangeHeader, null, "left", null);
                copyRangeStyle(confRange, sheet.getRange(classCooldowns.getRow() + 1 + (classCooldownsDone * 3), classCooldowns.getColumn() + 1 + playerDoneCount + classDoneCount, 2, playersInThisClass), null, "center", null);
                sheet.getRange(classCooldowns.getRow() + 3 + (classCooldownsDone * 3), classCooldowns.getColumn() + playerDoneCount + classDoneCount, 1, 1).setFontWeight("bold").setHorizontalAlignment("left");
                sheet.getRange(classCooldowns.getRow() + 3 + (classCooldownsDone * 3), classCooldowns.getColumn() + playerDoneCount + classDoneCount + 1, 1, playersInThisClass).setFontWeight("bold").setHorizontalAlignment("center");
              }

              if (classCooldown.indexOf("Power Infusion") > -1) {
                if (!gotAtLeastOnePI) {
                  if (playerByNameAsc.type == "Priest") {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                  }
                  else {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = getStringForLang("noAssign", langKeys, langTrans, "", "", "", "");
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = getStringForLang("noAssign", langKeys, langTrans, "", "", "", "");
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = getStringForLang("noAssign", langKeys, langTrans, "", "", "", "");
                  }
                }
                else {
                  if (playerByNameAsc.type == "Priest") {
                    if (PIoverwrittenWithCasts) {
                      if (PIonTrashSelf > 0)
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = PIonTrash + " (" + PIonTrashSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = PIonTrash;
                      if (PIOnBossesSelf > 0)
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = PIOnBosses + " (" + PIOnBossesSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = PIOnBosses;
                      if (PIonTrashSelf > 0 || PIOnBossesSelf > 0) {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (PIonTrash + PIOnBosses) + " (" + (PIonTrashSelf + PIOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")" + " (" + Math.round(PITotalUptime * cdMultiplesOfActive) + "%)";
                          sheet.getRange(classCooldowns.getRow() + 3 + (classCooldownsDone * 3), classCooldowns.getColumn() + 1 + playerDoneCount + classDoneCount).setFontSize(7);
                        } else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (PIonTrash + PIOnBosses) + " (" + (PIonTrashSelf + PIOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      }
                      else {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PIonTrash + PIOnBosses + " (" + Math.round(PITotalUptime * cdMultiplesOfActive) + "%)";
                        else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PIonTrash + PIOnBosses;
                      }
                    } else {
                      if (PIonTrashSelf > 0)
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = PIonTrash + "* (" + PIonTrashSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = PIonTrash + "*";
                      if (PIOnBossesSelf > 0)
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = PIOnBosses + "* (" + PIOnBossesSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = PIOnBosses + "*";
                      if (PIonTrashSelf > 0 || PIOnBossesSelf > 0) {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (PIonTrash + PIOnBosses) + "* (" + (PIonTrashSelf + PIOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")" + " (" + Math.round(PITotalUptime * cdMultiplesOfActive) + "%)";
                          sheet.getRange(classCooldowns.getRow() + 3 + (classCooldownsDone * 3), classCooldowns.getColumn() + 1 + playerDoneCount + classDoneCount).setFontSize(7);
                        } else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (PIonTrash + PIOnBosses) + "* (" + (PIonTrashSelf + PIOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      }
                      else {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PIonTrash + PIOnBosses + "*" + " (" + Math.round(PITotalUptime * cdMultiplesOfActive) + "%)";
                        else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PIonTrash + PIOnBosses + "*";
                      }
                    }
                  } else {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = PIonTrash;
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = PIOnBosses;
                    if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                      classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PIonTrash + PIOnBosses + " (" + Math.round(PITotalUptime * cdMultiplesOfActive) + "%)";
                    else
                      classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PIonTrash + PIOnBosses;
                  }
                  totalAmount += PIonTrash + PIOnBosses;
                }
              }
              else if (classCooldown.indexOf("Arcane Power") > -1) {
                if (!gotAtLeastOneAPComb) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = getStringForLang("winterChill", langKeys, langTrans, "", "", "", "")
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = getStringForLang("winterChill", langKeys, langTrans, "", "", "", "")
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = getStringForLang("winterChill", langKeys, langTrans, "", "", "", "")
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = APCombonTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = APCombOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = APCombonTrash + APCombOnBosses + " (" + Math.round((APCombonTrash + APCombOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = APCombonTrash + APCombOnBosses;
                  totalAmount += APCombonTrash + APCombOnBosses;
                }
              }
              else if (classCooldown.indexOf("Evocation") > -1) {
                if (!gotAtLeastOneEvocation) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = EvocationOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = EvocationOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = EvocationOnTrash + EvocationOnBosses + " (" + Math.round((EvocationOnTrash + EvocationOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = EvocationOnTrash + EvocationOnBosses;
                  totalAmount += EvocationOnTrash + EvocationOnBosses;
                }
              }
              else if (classCooldown.indexOf("Evasion") > -1) {
                if (!gotAtLeastOneEvasion) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = EvasionOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = EvasionOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = EvasionOnTrash + EvasionOnBosses + " (" + Math.round(EvasionTotalUptime * cdMultiplesOfActive) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = EvasionOnTrash + EvasionOnBosses;
                  totalAmount += EvasionOnTrash + EvasionOnBosses;
                }
              }
              else if (classCooldown.indexOf("Adrenaline Rush") > -1) {
                if (!gotAtLeastOneAdrenalineRush) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = AdrenalineRushOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = AdrenalineRushOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = AdrenalineRushOnTrash + AdrenalineRushOnBosses + " (" + Math.round(AdrenalineRushTotalUptime * cdMultiplesOfActive) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = AdrenalineRushOnTrash + AdrenalineRushOnBosses;
                  totalAmount += AdrenalineRushOnTrash + AdrenalineRushOnBosses;
                }
              }
              else if (classCooldown.indexOf("Berserker Rage") > -1) {
                if (!gotAtLeastOneBerserkerRage) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BerserkerRageOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BerserkerRageOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BerserkerRageOnTrash + BerserkerRageOnBosses + " (" + Math.round(BerserkerRageTotalUptime * cdMultiplesOfActive) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BerserkerRageOnTrash + BerserkerRageOnBosses;
                  totalAmount += BerserkerRageOnTrash + BerserkerRageOnBosses;
                }
              }
              else if (classCooldown.indexOf("Bloodrage") > -1) {
                if (!gotAtLeastOneBloodrage) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BloodrageOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BloodrageOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BloodrageOnTrash + BloodrageOnBosses + " (" + Math.round((BloodrageOnTrash + BloodrageOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BloodrageOnTrash + BloodrageOnBosses;
                  totalAmount += BloodrageOnTrash + BloodrageOnBosses;
                }
              }
              else if (classCooldown.indexOf("Challenging Shout") > -1) {
                if (!gotAtLeastOneChallengingShout) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ChallengingShoutOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ChallengingShoutOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ChallengingShoutOnTrash + ChallengingShoutOnBosses + " (" + Math.round((ChallengingShoutOnTrash + ChallengingShoutOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ChallengingShoutOnTrash + ChallengingShoutOnBosses;
                  totalAmount += ChallengingShoutOnTrash + ChallengingShoutOnBosses;
                }
              }
              else if (classCooldown.indexOf("Death Wish") > -1) {
                if (!gotAtLeastOneDeathWish) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DeathWishOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DeathWishOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DeathWishOnTrash + DeathWishOnBosses + " (" + Math.round(DeathWishTotalUptime * cdMultiplesOfActive) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DeathWishOnTrash + DeathWishOnBosses;
                  totalAmount += DeathWishOnTrash + DeathWishOnBosses;
                }
              }
              else if (classCooldown.indexOf("Last Stand") > -1) {
                if (!gotAtLeastOneLastStand) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = LastStandOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = LastStandOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = LastStandOnTrash + LastStandOnBosses + " (" + Math.round((LastStandOnTrash + LastStandOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = LastStandOnTrash + LastStandOnBosses;
                  totalAmount += LastStandOnTrash + LastStandOnBosses;
                }
              }
              else if (classCooldown.indexOf("Recklessness") > -1) {
                if (!gotAtLeastOneRecklessness) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = RecklessnessOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = RecklessnessOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = RecklessnessOnTrash + RecklessnessOnBosses + " (" + Math.round((RecklessnessOnTrash + RecklessnessOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = RecklessnessOnTrash + RecklessnessOnBosses;
                  totalAmount += RecklessnessOnTrash + RecklessnessOnBosses;
                }
              }
              else if (classCooldown.indexOf("Retaliation") > -1) {
                if (!gotAtLeastOneRetaliation) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = RetaliationOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = RetaliationOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = RetaliationOnTrash + RetaliationOnBosses + " (" + Math.round((RetaliationOnTrash + RetaliationOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = RetaliationOnTrash + RetaliationOnBosses;
                  totalAmount += RetaliationOnTrash + RetaliationOnBosses;
                }
              }
              else if (classCooldown.indexOf("Shield Wall") > -1) {
                if (!gotAtLeastOneShieldWall) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ShieldWallOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ShieldWallOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ShieldWallOnTrash + ShieldWallOnBosses + " (" + Math.round((ShieldWallOnTrash + ShieldWallOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ShieldWallOnTrash + ShieldWallOnBosses;
                  totalAmount += ShieldWallOnTrash + ShieldWallOnBosses;
                }
              }
              else if (classCooldown.indexOf("Nature's Swiftness") > -1) {
                if (!gotAtLeastOneNaturesSwiftness) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = NaturesSwiftnessOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = NaturesSwiftnessOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = NaturesSwiftnessOnTrash + NaturesSwiftnessOnBosses + " (" + Math.round((NaturesSwiftnessOnTrash + NaturesSwiftnessOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = NaturesSwiftnessOnTrash + NaturesSwiftnessOnBosses;
                  totalAmount += NaturesSwiftnessOnTrash + NaturesSwiftnessOnBosses;
                }
              }
              else if (classCooldown.indexOf("Elemental Mastery") > -1) {
                if (!gotAtLeastOneElementalMastery) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ElementalMasteryOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ElementalMasteryOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ElementalMasteryOnTrash + ElementalMasteryOnBosses + " (" + Math.round((ElementalMasteryOnTrash + ElementalMasteryOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ElementalMasteryOnTrash + ElementalMasteryOnBosses;
                  totalAmount += ElementalMasteryOnTrash + ElementalMasteryOnBosses;
                }
              }
              else if (classCooldown.indexOf("Rebirth") > -1) {
                if (!gotAtLeastOneRebirth) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = RebirthOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = RebirthOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = RebirthOnTrash + RebirthOnBosses + " (" + Math.round((RebirthOnTrash + RebirthOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = RebirthOnTrash + RebirthOnBosses;
                  totalAmount += RebirthOnTrash + RebirthOnBosses;
                }
              }
              else if (classCooldown.indexOf("Challenging Roar") > -1) {
                if (!gotAtLeastOneChallengingRoar) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ChallengingRoarOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ChallengingRoarOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ChallengingRoarOnTrash + ChallengingRoarOnBosses + " (" + Math.round((ChallengingRoarOnTrash + ChallengingRoarOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ChallengingRoarOnTrash + ChallengingRoarOnBosses;
                  totalAmount += ChallengingRoarOnTrash + ChallengingRoarOnBosses;
                }
              }
              else if (classCooldown.indexOf("Dash") > -1) {
                if (!gotAtLeastOneDash) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DashOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DashOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DashOnTrash + DashOnBosses + " (" + Math.round((DashOnTrash + DashOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DashOnTrash + DashOnBosses;
                  totalAmount += DashOnTrash + DashOnBosses;
                }
              }
              else if (classCooldown.indexOf("Frenzied Regeneration") > -1) {
                if (!gotAtLeastOneFrenziedRegeneration) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = FrenziedRegenerationOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = FrenziedRegenerationOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = FrenziedRegenerationOnTrash + FrenziedRegenerationOnBosses + " (" + Math.round((FrenziedRegenerationOnTrash + FrenziedRegenerationOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  } else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = FrenziedRegenerationOnTrash + FrenziedRegenerationOnBosses;
                  totalAmount += FrenziedRegenerationOnTrash + FrenziedRegenerationOnBosses;
                }
              }
              else if (classCooldown.indexOf("Innervate") > -1) {
                if (!gotAtLeastOneInnervate) {
                  if (playerByNameAsc.type == "Druid") {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                  }
                  else {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                  }
                }
                else {
                  if (playerByNameAsc.type == "Druid") {
                    if (InnervateUsedBySelf) {
                      if (InnervateOnTrashSelf > 0)
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = InnervateOnTrash + " (" + InnervateOnTrashSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = InnervateOnTrash;
                      if (InnervateOnBossesSelf > 0)
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = InnervateOnBosses + " (" + InnervateOnBossesSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = InnervateOnBosses;
                      if (InnervateOnTrashSelf > 0 || InnervateOnBossesSelf > 0) {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (InnervateOnTrash + InnervateOnBosses) + " (" + (InnervateOnTrashSelf + InnervateOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")" + " (" + Math.round((InnervateOnTrash + InnervateOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                          sheet.getRange(classCooldowns.getRow() + 3 + (classCooldownsDone * 3), classCooldowns.getColumn() + 1 + playerCount + classDoneCount).setFontSize(7);
                        } else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (InnervateOnTrash + InnervateOnBosses) + " (" + (InnervateOnTrashSelf + InnervateOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      }
                      else {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = InnervateOnTrash + InnervateOnBosses + " (" + Math.round((InnervateOnTrash + InnervateOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                        else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = InnervateOnTrash + InnervateOnBosses;
                      }
                    } else {
                      if (InnervateOnTrashSelf > 0)
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = InnervateOnTrash + "* (" + InnervateOnTrashSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = InnervateOnTrash + "*";
                      if (InnervateOnBossesSelf > 0)
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = InnervateOnBosses + "* (" + InnervateOnBossesSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = InnervateOnBosses + "*";
                      if (InnervateOnTrashSelf > 0 || InnervateOnBossesSelf > 0) {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (InnervateOnTrash + InnervateOnBosses) + "* (" + (InnervateOnTrashSelf + InnervateOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")" + " (" + Math.round((InnervateOnTrash + InnervateOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                          sheet.getRange(classCooldowns.getRow() + 3 + (classCooldownsDone * 3), classCooldowns.getColumn() + 1 + playerCount + classDoneCount).setFontSize(7);
                        } else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (InnervateOnTrash + InnervateOnBosses) + "* (" + (InnervateOnTrashSelf + InnervateOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      }
                      else {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (InnervateOnTrash + InnervateOnBosses) + "*" + " (" + Math.round((InnervateOnTrash + InnervateOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                        else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (InnervateOnTrash + InnervateOnBosses) + "*";
                      }
                    }
                  } else {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = InnervateOnTrash;
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = InnervateOnBosses;
                    if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                      classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (InnervateOnTrash + InnervateOnBosses) + " (" + Math.round((InnervateOnTrash + InnervateOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                    else
                      classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = InnervateOnTrash + InnervateOnBosses;
                  }
                  totalAmount += InnervateOnTrash + InnervateOnBosses;
                }
              }
              else if (classCooldown.indexOf("Blessing of Protection") > -1) {
                if (!gotAtLeastOneBoP) {
                  if (playerByNameAsc.type == "Paladin") {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                  }
                  else {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                  }
                }
                else {
                  if (playerByNameAsc.type == "Paladin") {
                    if (BoPUsedBySelf) {
                      if (BoPOnTrashSelf > 0)
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BoPOnTrash + " (" + BoPOnTrashSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BoPOnTrash;
                      if (BoPOnBossesSelf > 0)
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BoPOnBosses + " (" + BoPOnBossesSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BoPOnBosses;
                      if (BoPOnTrashSelf > 0 || BoPOnBossesSelf > 0) {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (BoPOnTrash + BoPOnBosses) + " (" + (BoPOnTrashSelf + BoPOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")" + " (" + Math.round((BoPOnTrash + BoPOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                          sheet.getRange(classCooldowns.getRow() + 3 + (classCooldownsDone * 3), classCooldowns.getColumn() + 1 + playerCount + classDoneCount).setFontSize(7);
                        } else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (BoPOnTrash + BoPOnBosses) + " (" + (BoPOnTrashSelf + BoPOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      }
                      else {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BoPOnTrash + BoPOnBosses + " (" + Math.round((BoPOnTrash + BoPOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                        else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BoPOnTrash + BoPOnBosses;
                      }
                    } else {
                      if (BoPOnTrashSelf > 0)
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BoPOnTrash + "* (" + BoPOnTrashSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BoPOnTrash + "*";
                      if (BoPOnBossesSelf > 0)
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BoPOnBosses + "* (" + BoPOnBossesSelf + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      else
                        classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BoPOnBosses + "*";
                      if (BoPOnTrashSelf > 0 || BoPOnBossesSelf > 0) {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (BoPOnTrash + BoPOnBosses) + "* (" + (BoPOnTrashSelf + BoPOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")" + " (" + Math.round((BoPOnTrash + BoPOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                          sheet.getRange(classCooldowns.getRow() + 3 + (classCooldownsDone * 3), classCooldowns.getColumn() + 1 + playerCount + classDoneCount).setFontSize(7);
                        } else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (BoPOnTrash + BoPOnBosses) + "* (" + (BoPOnTrashSelf + BoPOnBossesSelf) + " " + getStringForLang("self", langKeys, langTrans, "", "", "", "") + ")";
                      }
                      else {
                        if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (BoPOnTrash + BoPOnBosses) + "*" + " (" + Math.round((BoPOnTrash + BoPOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                        else
                          classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (BoPOnTrash + BoPOnBosses) + "*";
                      }
                    }
                  } else {
                    classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BoPOnTrash;
                    classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BoPOnBosses;
                    if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                      classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = (BoPOnTrash + BoPOnBosses) + " (" + Math.round((BoPOnTrash + BoPOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                    else
                      classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BoPOnTrash + BoPOnBosses;
                  }
                  totalAmount += BoPOnTrash + BoPOnBosses;
                }
              }
              else if (classCooldown.indexOf("Tranquility") > -1) {
                if (!gotAtLeastOneTranquility) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = TranquilityOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = TranquilityOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = TranquilityOnTrash + TranquilityOnBosses + " (" + Math.round((TranquilityOnTrash + TranquilityOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = TranquilityOnTrash + TranquilityOnBosses;
                  totalAmount += TranquilityOnTrash + TranquilityOnBosses;
                }
              }
              else if (classCooldown.indexOf("Rapid Fire") > -1) {
                if (!gotAtLeastOneRapidFire) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = RapidFireOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = RapidFireOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = RapidFireOnTrash + RapidFireOnBosses + " (" + Math.round(RapidFireTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = RapidFireOnTrash + RapidFireOnBosses;
                  totalAmount += RapidFireOnTrash + RapidFireOnBosses;
                }
              }
              else if (classCooldown.indexOf("Bestial Wrath") > -1) {
                if (!gotAtLeastOneBestialWrath) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BestialWrathOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BestialWrathOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BestialWrathOnTrash + BestialWrathOnBosses + " (" + Math.round(BestialWrathTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BestialWrathOnTrash + BestialWrathOnBosses;
                  totalAmount += BestialWrathOnTrash + BestialWrathOnBosses;
                }
              }
              else if (classCooldown.indexOf("Readiness") > -1) {
                if (!gotAtLeastOneReadiness) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ReadinessOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ReadinessOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ReadinessOnTrash + ReadinessOnBosses + " (" + Math.round((ReadinessOnTrash + ReadinessOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ReadinessOnTrash + ReadinessOnBosses;
                  totalAmount += ReadinessOnTrash + ReadinessOnBosses;
                }
              }
              else if (classCooldown.indexOf("Deterrence") > -1) {
                if (!gotAtLeastOneDeterrence) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DeterrenceOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DeterrenceOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DeterrenceOnTrash + DeterrenceOnBosses + " (" + Math.round((DeterrenceOnTrash + DeterrenceOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DeterrenceOnTrash + DeterrenceOnBosses;
                  totalAmount += DeterrenceOnTrash + DeterrenceOnBosses;
                }
              }
              else if (classCooldown.indexOf("Inner Focus") > -1) {
                if (!gotAtLeastOneInnerFocus) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = InnerFocusOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = InnerFocusOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = InnerFocusOnTrash + InnerFocusOnBosses + " (" + Math.round((InnerFocusOnTrash + InnerFocusOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = InnerFocusOnTrash + InnerFocusOnBosses;
                  totalAmount += InnerFocusOnTrash + InnerFocusOnBosses;
                }
              }
              else if (classCooldown.indexOf("Shadowfiend") > -1) {
                if (!gotAtLeastOneShadowfiend) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ShadowfiendOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ShadowfiendOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ShadowfiendOnTrash + ShadowfiendOnBosses + " (" + Math.round((ShadowfiendOnTrash + ShadowfiendOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ShadowfiendOnTrash + ShadowfiendOnBosses;
                  totalAmount += ShadowfiendOnTrash + ShadowfiendOnBosses;
                }
              }
              else if (classCooldown.indexOf("Blessing of Protection") > -1) {
                if (!gotAtLeastOneBoP) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BoPOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BoPOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BoPOnTrash + BoPOnBosses + " (" + Math.round((BoPOnTrash + BoPOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BoPOnTrash + BoPOnBosses;
                  totalAmount += BoPOnTrash + BoPOnBosses;
                }
              }
              else if (classCooldown.indexOf("Divine Favor") > -1) {
                if (!gotAtLeastOneDivineFavor) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DivineFavorOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DivineFavorOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineFavorOnTrash + DivineFavorOnBosses + " (" + Math.round((DivineFavorOnTrash + DivineFavorOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineFavorOnTrash + DivineFavorOnBosses;
                  totalAmount += DivineFavorOnTrash + DivineFavorOnBosses;
                }
              }
              else if (classCooldown.indexOf("Divine Intervention") > -1) {
                if (!gotAtLeastOneDivineIntervention) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DivineInterventionOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DivineInterventionOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineInterventionOnTrash + DivineInterventionOnBosses + " (" + Math.round((DivineInterventionOnTrash + DivineInterventionOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineInterventionOnTrash + DivineInterventionOnBosses;
                  totalAmount += DivineInterventionOnTrash + DivineInterventionOnBosses;
                }
              }
              else if (classCooldown.indexOf("Divine Protection") > -1) {
                if (!gotAtLeastOneDivineProtection) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DivineProtectionOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DivineProtectionOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineProtectionOnTrash + DivineProtectionOnBosses + " (" + Math.round((DivineProtectionOnTrash + DivineProtectionOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineProtectionOnTrash + DivineProtectionOnBosses;
                  totalAmount += DivineProtectionOnTrash + DivineProtectionOnBosses;
                }
              }
              else if (classCooldown.indexOf("Divine Shield") > -1) {
                if (!gotAtLeastOneDivineShield) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DivineShieldOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DivineShieldOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineShieldOnTrash + DivineShieldOnBosses + " (" + Math.round((DivineShieldOnTrash + DivineShieldOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineShieldOnTrash + DivineShieldOnBosses;
                  totalAmount += DivineShieldOnTrash + DivineShieldOnBosses;
                }
              }
              else if (classCooldown.indexOf("Lay on Hands") > -1) {
                if (!gotAtLeastOneLayOnHands) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = LayOnHandsOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = LayOnHandsOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = LayOnHandsOnTrash + LayOnHandsOnBosses + " (" + Math.round((LayOnHandsOnTrash + LayOnHandsOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = LayOnHandsOnTrash + LayOnHandsOnBosses;
                  totalAmount += LayOnHandsOnTrash + LayOnHandsOnBosses;
                }
              }
              else if (classCooldown.indexOf("Mana Tide Totem") > -1) {
                if (!gotAtLeastOneManaTideTotem) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ManaTideTotemOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ManaTideTotemOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ManaTideTotemOnTrash + ManaTideTotemOnBosses + " (" + Math.round((ManaTideTotemOnTrash + ManaTideTotemOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ManaTideTotemOnTrash + ManaTideTotemOnBosses;
                  totalAmount += ManaTideTotemOnTrash + ManaTideTotemOnBosses;
                }
              }
              else if (classCooldown.indexOf("Vanish") > -1) {
                if (!gotAtLeastOneVanish) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = VanishOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = VanishOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = VanishOnTrash + VanishOnBosses + " (" + Math.round((VanishOnTrash + VanishOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = VanishOnTrash + VanishOnBosses;
                  totalAmount += VanishOnTrash + VanishOnBosses;
                }
              }
              else if (classCooldown.indexOf("Barkskin") > -1) {
                if (!gotAtLeastOneBarkskin) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BarkskinOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BarkskinOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BarkskinOnTrash + BarkskinOnBosses + " (" + Math.round(BarkskinTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BarkskinOnTrash + BarkskinOnBosses;
                  totalAmount += BarkskinOnTrash + BarkskinOnBosses;
                }
              }
              else if (classCooldown.indexOf("Force of Nature") > -1) {
                if (!gotAtLeastOneForceOfNature) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ForceOfNatureOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ForceOfNatureOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ForceOfNatureOnTrash + ForceOfNatureOnBosses + " (" + Math.round((ForceOfNatureOnTrash + ForceOfNatureOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ForceOfNatureOnTrash + ForceOfNatureOnBosses;
                  totalAmount += ForceOfNatureOnTrash + ForceOfNatureOnBosses;
                }
              }
              else if (classCooldown.indexOf("Summon Water Elemental") > -1) {
                if (!gotAtLeastOneSummonWaterElemental) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = SummonWaterElementalOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = SummonWaterElementalOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = SummonWaterElementalOnTrash + SummonWaterElementalOnBosses + " (" + Math.round((SummonWaterElementalOnTrash + SummonWaterElementalOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = SummonWaterElementalOnTrash + SummonWaterElementalOnBosses;
                  totalAmount += SummonWaterElementalOnTrash + SummonWaterElementalOnBosses;
                }
              }
              else if (classCooldown.indexOf("Invisibility") > -1) {
                if (!gotAtLeastOneInvisibility) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = InvisibilityOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = InvisibilityOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = InvisibilityOnTrash + InvisibilityOnBosses + " (" + Math.round((InvisibilityOnTrash + InvisibilityOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = InvisibilityOnTrash + InvisibilityOnBosses;
                  totalAmount += InvisibilityOnTrash + InvisibilityOnBosses;
                }
              }
              else if (classCooldown.indexOf("Cold Snap") > -1) {
                if (!gotAtLeastOneColdSnap) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ColdSnapOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ColdSnapOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ColdSnapOnTrash + ColdSnapOnBosses + " (" + Math.round((ColdSnapOnTrash + ColdSnapOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ColdSnapOnTrash + ColdSnapOnBosses;
                  totalAmount += ColdSnapOnTrash + ColdSnapOnBosses;
                }
              }
              else if (classCooldown.indexOf("Desperate Prayer") > -1) {
                if (!gotAtLeastOneDesperatePrayer) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DesperatePrayerOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DesperatePrayerOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DesperatePrayerOnTrash + DesperatePrayerOnBosses + " (" + Math.round((DesperatePrayerOnTrash + DesperatePrayerOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DesperatePrayerOnTrash + DesperatePrayerOnBosses;
                  totalAmount += DesperatePrayerOnTrash + DesperatePrayerOnBosses;
                }
              }
              else if (classCooldown.indexOf("Gift of the Naaru") > -1) {
                if (!gotAtLeastOneGiftOfTheNaaru) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = GiftOfTheNaaruOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = GiftOfTheNaaruOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = GiftOfTheNaaruOnTrash + GiftOfTheNaaruOnBosses + " (" + Math.round((GiftOfTheNaaruOnTrash + GiftOfTheNaaruOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = GiftOfTheNaaruOnTrash + GiftOfTheNaaruOnBosses;
                  totalAmount += GiftOfTheNaaruOnTrash + GiftOfTheNaaruOnBosses;
                }
              }
              else if (classCooldown.indexOf("Soulshatter") > -1) {
                if (!gotAtLeastOneSoulshatter) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = SoulshatterOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = SoulshatterOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = SoulshatterOnTrash + SoulshatterOnBosses + " (" + Math.round((SoulshatterOnTrash + SoulshatterOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = SoulshatterOnTrash + SoulshatterOnBosses;
                  totalAmount += SoulshatterOnTrash + SoulshatterOnBosses;
                }
              }
              else if (classCooldown.indexOf("Icy Veins") > -1) {
                if (!gotAtLeastOneIcyVeins) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = IcyVeinsOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = IcyVeinsOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = IcyVeinsOnTrash + IcyVeinsOnBosses + " (" + Math.round(IcyVeinsTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = IcyVeinsOnTrash + IcyVeinsOnBosses;
                  totalAmount += IcyVeinsOnTrash + IcyVeinsOnBosses;
                }
              }
              else if (classCooldown.indexOf("Presence of Mind") > -1) {
                if (!gotAtLeastOnePoM) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = PoMOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = PoMOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PoMOnTrash + PoMOnBosses + " (" + Math.round((PoMOnTrash + PoMOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PoMOnTrash + PoMOnBosses;
                  totalAmount += PoMOnTrash + PoMOnBosses;
                }
              }
              else if (classCooldown.indexOf("Avenging Wrath") > -1) {
                if (!gotAtLeastOneAvengingWrath) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = AvengingWrathOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = AvengingWrathOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = AvengingWrathOnTrash + AvengingWrathOnBosses + " (" + Math.round(AvengingWrathTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = AvengingWrathOnTrash + AvengingWrathOnBosses;
                  totalAmount += AvengingWrathOnTrash + AvengingWrathOnBosses;
                }
              }
              else if (classCooldown.indexOf("Divine Illumination") > -1) {
                if (!gotAtLeastOneDivineIllumination) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DivineIlluminationOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DivineIlluminationOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineIlluminationOnTrash + DivineIlluminationOnBosses + " (" + Math.round(DivineIlluminationTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DivineIlluminationOnTrash + DivineIlluminationOnBosses;
                  totalAmount += DivineIlluminationOnTrash + DivineIlluminationOnBosses;
                }
              }
              else if (classCooldown.indexOf("Devouring Plague") > -1) {
                if (!gotAtLeastOneDevouringPlague) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = DevouringPlagueOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = DevouringPlagueOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DevouringPlagueOnTrash + DevouringPlagueOnBosses + " (" + Math.round((DevouringPlagueOnTrash + DevouringPlagueOnBosses) * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / (Number(classCooldown.split("--")[1]))) + 1)) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = DevouringPlagueOnTrash + DevouringPlagueOnBosses;
                  totalAmount += DevouringPlagueOnTrash + DevouringPlagueOnBosses;
                }
              }
              else if (classCooldown.indexOf("Pain Suppression") > -1) {
                if (!gotAtLeastOnePainSuppression) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = PainSuppressionOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = PainSuppressionOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PainSuppressionOnTrash + PainSuppressionOnBosses + " (" + Math.round(PainSuppressionTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = PainSuppressionOnTrash + PainSuppressionOnBosses;
                  totalAmount += PainSuppressionOnTrash + PainSuppressionOnBosses;
                }
              }
              else if (classCooldown.indexOf("Cloak of Shadows") > -1) {
                if (!gotAtLeastOneCloakOfShadows) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = CloakOfShadowsOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = CloakOfShadowsOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = CloakOfShadowsOnTrash + CloakOfShadowsOnBosses + " (" + Math.round(CloakOfShadowsTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = CloakOfShadowsOnTrash + CloakOfShadowsOnBosses;
                  totalAmount += CloakOfShadowsOnTrash + CloakOfShadowsOnBosses;
                }
              }
              else if (classCooldown.indexOf("Bloodlust") > -1) {
                if (!gotAtLeastOneBloodlust) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = BloodlustOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = BloodlustOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BloodlustOnTrash + BloodlustOnBosses + " (" + Math.round(BloodlustTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = BloodlustOnTrash + BloodlustOnBosses;
                  totalAmount += BloodlustOnTrash + BloodlustOnBosses;
                }
              }
              else if (classCooldown.indexOf("Amplify Curse") > -1) {
                if (!gotAtLeastOneAmplifyCurse) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = AmplifyCurseOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = AmplifyCurseOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = AmplifyCurseOnTrash + AmplifyCurseOnBosses + " (" + Math.round(AmplifyCurseTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = AmplifyCurseOnTrash + AmplifyCurseOnBosses;
                  totalAmount += AmplifyCurseOnTrash + AmplifyCurseOnBosses;
                }
              }
              else if (classCooldown.indexOf("Heroism") > -1) {
                if (!gotAtLeastOneHeroism) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = HeroismOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = HeroismOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = HeroismOnTrash + HeroismOnBosses + " (" + Math.round(HeroismTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = HeroismOnTrash + HeroismOnBosses;
                  totalAmount += HeroismOnTrash + HeroismOnBosses;
                }
              }
              else if (classCooldown.indexOf("Shamanistic Rage") > -1) {
                if (!gotAtLeastOneShamanisticRage) {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = "0";
                  classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = "0";
                }
                else {
                  classCooldownsArr[line + 1][playerDoneCount + classDoneCount + 1] = ShamanisticRageOnTrash;
                  classCooldownsArr[line + 2][playerDoneCount + classDoneCount + 1] = ShamanisticRageOnBosses;
                  if (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0))
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ShamanisticRageOnTrash + ShamanisticRageOnBosses + " (" + Math.round(ShamanisticRageTotalUptime * cdMultiplesOfActive) + "%)";
                  else
                    classCooldownsArr[line + 3][playerDoneCount + classDoneCount + 1] = ShamanisticRageOnTrash + ShamanisticRageOnBosses;
                  totalAmount += ShamanisticRageOnTrash + ShamanisticRageOnBosses;
                }
              }
              line += 3;
              classCooldownsDone++;
            }
          })
        }

        //fill in damage taken
        var totalAmount = 0;
        var damageTakenToTrackSpells = sheet.getRange(damageTaken.getRow() + 1, damageTaken.getColumn(), damageTakenMaxEntries, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
        damageTakenToTrackSpells.forEach(function (damageTakenToTrackSpell, damageTakenToTrackSpellCount) {
          if (damageTakenArr[damageTakenToTrackSpellCount + 1] == null || damageTakenArr[damageTakenToTrackSpellCount + 1].length == 0) {
            damageTakenArr[damageTakenToTrackSpellCount + 1] = [];
            damageTakenArr[damageTakenToTrackSpellCount + 1].push(damageTakenToTrackSpell.split(" [")[0] + " " + damageTakenToTrackSpell.split("] ")[1]);
          }
          var amount = 0;
          var spellIdsToLink = "";
          damageTakenTotalData.entries.forEach(function (damageTakenEntry, damageTakenEntryCount) {
            if (damageTakenToTrackSpell.indexOf("[") > -1) {
              damageTakenToTrackSpell.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                if (spellId == damageTakenEntry.guid.toString()) {
                  amount += damageTakenEntry.total;
                  if (spellIdsToLink.indexOf(spellId) < 0)
                    spellIdsToLink += spellId + ",";
                }
              })
            }
          })
          var urlToLink = urlDamageTakenTop.replace(baseUrl + "report/tables/damage-taken/", baseUrlFrontEnd).replace(logId, logId + "#type=damage-taken").replace(apiKeyString, "").replace(startEndStringNoFilter, "") + "&translate=true&boss=" + bossString + "&difficulty=0&source=" + playerByNameAsc.id + "&view=events&pins=2%24Off%24%23244F4B%24expression%24ability.id%20IN%20(" + spellIdsToLink.substr(0, spellIdsToLink.length - 1) + ")";
          if (previousClass != playerByNameAsc.type)
            damageTakenArr[damageTakenToTrackSpellCount + 1].push("");

          if (amount == 0)
            damageTakenArr[damageTakenToTrackSpellCount + 1].push("");
          else
            damageTakenArr[damageTakenToTrackSpellCount + 1].push("=HYPERLINK(\"" + urlToLink + "\"," + amount + ")");
          totalAmount += amount;
        })
        for (var i = damageTakenArr.length; i <= damageTakenMaxEntries; i++) {
          addSingleEntryToMultiDimArray(damageTakenArr, "");
        }
        var reflectedTotal = 0;
        if (damageReflectedData != null && damageReflectedData.entries != null) {
          damageReflectedData.entries.forEach(function (reflectedSpell, reflectedSpellCount) {
            if (reflectedSpell.id == playerByNameAsc.id)
              if (reflectedSpell.total > 0) {
                reflectedTotal += reflectedSpell.total;
              }
          })
        }
        if (damageTakenArr.length == damageTakenMaxEntries + 1) {
          if (showDamageReflectRow) {
            addSingleEntryToMultiDimArray(damageTakenArr, getStringForLang("damageReflected", langKeys, langTrans, "", "", "", ""));
          }
          if (showFriendlyFireRow) {
            addSingleEntryToMultiDimArray(damageTakenArr, getStringForLang("damageHostile", langKeys, langTrans, "", "", "", ""));
            addSingleEntryToMultiDimArray(damageTakenArr, getStringForLang("friendlyFire", langKeys, langTrans, "", "", "", ""));
          }
          if (showDeathCountRow) {
            if (onlyBosses || onlyTrash || (onlyFightNr != null && onlyFightNr.toString().length > 0))
              addSingleEntryToMultiDimArray(damageTakenArr, getStringForLang("totalDeaths", langKeys, langTrans, "", "", "", ""));
            else
              addSingleEntryToMultiDimArray(damageTakenArr, getStringForLang("totalDeaths", langKeys, langTrans, "", "", "", "") + " (" + getStringForLang("justOnTrash", langKeys, langTrans, "", "", "", "") + ")");
          }
          addSingleEntryToMultiDimArray(damageTakenArr, getStringForLang("avoidableDamageTaken", langKeys, langTrans, "", "", "", ""));
          addSingleEntryToMultiDimArray(damageTakenArr, "");
        }
        if (showDamageReflectRow && reflectedTotal > 0) {
          if (showDeathCountRow) {
            if (showFriendlyFireRow) {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 6].push("");
              damageTakenArr[damageTakenArr.length - 6].push("=HYPERLINK(\"" + urlDamageReflectedLinkPlayer + "\"," + reflectedTotal + ")");
            } else {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 4].push("");
              damageTakenArr[damageTakenArr.length - 4].push("=HYPERLINK(\"" + urlDamageReflectedLinkPlayer + "\"," + reflectedTotal + ")");
            }
          } else {
            if (showFriendlyFireRow) {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 5].push("");
              damageTakenArr[damageTakenArr.length - 5].push("=HYPERLINK(\"" + urlDamageReflectedLinkPlayer + "\"," + reflectedTotal + ")");
            } else {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 3].push("");
              damageTakenArr[damageTakenArr.length - 3].push("=HYPERLINK(\"" + urlDamageReflectedLinkPlayer + "\"," + reflectedTotal + ")");
            }
          }
          totalAmount += reflectedTotal;
        }
        else if (showDamageReflectRow) {
          if (showDeathCountRow) {
            if (showFriendlyFireRow) {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 6].push("");
              damageTakenArr[damageTakenArr.length - 6].push("");
            } else {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 4].push("");
              damageTakenArr[damageTakenArr.length - 4].push("");
            }
          } else {
            if (showFriendlyFireRow)
              damageTakenArr[damageTakenArr.length - 3].push("");
            else
              damageTakenArr[damageTakenArr.length - 3].push("");
          }
        }

        if (showFriendlyFireRow) {
          if (showDeathCountRow) {
            if (hostilePlayersTotal > 0) {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 5].push("");
              damageTakenArr[damageTakenArr.length - 5].push("=HYPERLINK(\"" + urlHostilePlayersLinkPlayer + "\"," + hostilePlayersTotal + ")");
              totalAmount += hostilePlayersTotal;
            }
            else {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 5].push("");
              damageTakenArr[damageTakenArr.length - 5].push("");
            }

            if (friendlyFireTotal > 0) {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 4].push("");
              damageTakenArr[damageTakenArr.length - 4].push("=HYPERLINK(\"" + urlFriendlyFireLinkPlayer + "\"," + friendlyFireTotal + ")");
              totalAmount += friendlyFireTotal;
            }
            else {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 4].push("");
              damageTakenArr[damageTakenArr.length - 4].push("");
            }
          } else {
            if (hostilePlayersTotal > 0) {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 4].push("");
              damageTakenArr[damageTakenArr.length - 4].push("=HYPERLINK(\"" + urlHostilePlayersLinkPlayer + "\"," + hostilePlayersTotal + ")");
              totalAmount += hostilePlayersTotal;
            }
            else {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 4].push("");
              damageTakenArr[damageTakenArr.length - 4].push("");
            }

            if (friendlyFireTotal > 0) {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 3].push("");
              damageTakenArr[damageTakenArr.length - 3].push("=HYPERLINK('" + urlFriendlyFireLinkPlayer + "'," + friendlyFireTotal + ")");
              totalAmount += friendlyFireTotal;
            }
            else {
              if (previousClass != playerByNameAsc.type)
                damageTakenArr[damageTakenArr.length - 3].push("");
              damageTakenArr[damageTakenArr.length - 3].push("");
            }
          }
        }

        if (showDeathCountRow) {
          if (previousClass != playerByNameAsc.type)
            damageTakenArr[damageTakenArr.length - 3].push("");
          if (deathsData.entries.length > 0) {
            var deathsTotal = 0;
            deathsData.entries.forEach(function (deathDataTotal, deathDataTotalCount) {
              if (deathDataTotal.id == playerByNameAsc.id)
                deathsTotal++;
            })
            if ((!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)) && deathsDataTrash.entries.length > 0) {
              var deathsOnTrash = 0;
              deathsDataTrash.entries.forEach(function (deathDataTrash, deathDataTrashCount) {
                if (deathDataTrash.id == playerByNameAsc.id)
                  deathsOnTrash++;
              })
              damageTakenArr[damageTakenArr.length - 3].push(deathsTotal + " (" + deathsOnTrash + ")");
            } else {
              if (onlyBosses || onlyTrash || (onlyFightNr != null && onlyFightNr.toString().length > 0))
                damageTakenArr[damageTakenArr.length - 3].push(deathsTotal);
              else
                damageTakenArr[damageTakenArr.length - 3].push(deathsTotal + " (0)");
            }
          }
          else
            damageTakenArr[damageTakenArr.length - 3].push("");
        }

        if (previousClass != playerByNameAsc.type)
          damageTakenArr[damageTakenArr.length - 2].push("");
        damageTakenArr[damageTakenArr.length - 2].push(totalAmount);
        if (previousClass != playerByNameAsc.type)
          damageTakenArr[damageTakenArr.length - 1].push("");
        damageTakenArr[damageTakenArr.length - 1].push("=" + totalAmount.toString() + "/MAX(" + sheet.getRange(damageTaken.getRow() + damageTakenArr.length - 2, damageTaken.getColumn() + 1, 1, maxColumns).getA1Notation() + ";1)");

        //fill in debuffs applied
        var totalAmount = 0;
        var debuffsToTrackSpells = sheet.getRange(debuffs.getRow() + 1, debuffs.getColumn(), debuffsMaxEntries, 1).getValues().reduce(function (ar, e) { if (e[0]) ar.push(e[0]); return ar; }, []);
        debuffsToTrackSpells.forEach(function (debuffsToTrackSpell, debuffsToTrackSpellCount) {
          if (debuffsArr[debuffsToTrackSpellCount + 1] == null || debuffsArr[debuffsToTrackSpellCount + 1].length == 0) {
            debuffsArr[debuffsToTrackSpellCount + 1] = [];
            debuffsArr[debuffsToTrackSpellCount + 1].push(debuffsToTrackSpell.split(" [")[0] + " " + debuffsToTrackSpell.split("] ")[1]);
          }
          var amount = 0;
          debuffsData.auras.forEach(function (debuff, debuffCount) {
            if (debuffsToTrackSpell.indexOf("[") > -1) {
              debuffsToTrackSpell.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                if (spellId == debuff.guid.toString()) {
                  amount += debuff.totalUses;
                }
              })
            }
          })
          if (previousClass != playerByNameAsc.type)
            debuffsArr[debuffsToTrackSpellCount + 1].push("");
          if (amount == 0)
            debuffsArr[debuffsToTrackSpellCount + 1].push("");
          else
            debuffsArr[debuffsToTrackSpellCount + 1].push(amount);
          totalAmount += amount;
        })
        for (var i = debuffsArr.length; i <= debuffsMaxEntries; i++) {
          addSingleEntryToMultiDimArray(debuffsArr, "");
        }

        //fill in stats and miscellaneous
        if (playerDoneCount == 0) {
          copyRowStyles(conf, sheet, confStatsAndMiscToTrack, statsAndMiscToTrack.length, statsAndMisc.getRow() + 1, statsAndMisc.getColumn(), maxColumns, true, "right", darkMode);

          var confColumnWidth = conf.getColumnWidth(confStatsAndMiscToTrack.getColumn());
          if (confColumnWidth > maxColumnWidth) {
            maxColumnWidth = confColumnWidth;
          }
        }
        statsAndMiscToTrack.forEach(function (statOrMisc, statOrMiscCount) {
          var amount = 0;
          var uptime = 0;
          var totalDamageDoneUses = 0;
          var totalHealingDoneUses = 0;
          var totalDamageTakenUses = 0;
          if (statsAndMiscArr[statOrMiscCount + 1] == null || statsAndMiscArr[statOrMiscCount + 1].length == 0) {
            statsAndMiscArr[statOrMiscCount + 1] = [];
            statsAndMiscArr[statOrMiscCount + 1].push(statsAndMiscToTrackLang[statOrMiscCount].split(" [")[0].split(" {")[0]);
          }
          if (previousClass != playerByNameAsc.type)
            statsAndMiscArr[statOrMiscCount + 1].push("");
          if (statOrMisc.indexOf("Battle Shout uptime on you") > -1) {
            uptime += getUptimeForDebuffSpellId("2048", buffsData, totalTimeElapsedRaw);
            if (uptime > 0)
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = uptime + "%";
            else
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = "";
          } else if (statOrMisc.indexOf("Commanding Shout uptime on you") > -1) {
            uptime += getUptimeForDebuffSpellId("469", buffsData, totalTimeElapsedRaw);
            if (uptime > 0)
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = uptime + "%";
            else
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = "";
          } else if (statOrMisc.indexOf("Tinnitus uptime on you") > -1) {
            uptime += getUptimeForDebuffSpellId("51120", debuffsData, totalTimeElapsedRaw);
            if (uptime > 0)
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = uptime + "%";
            else
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = "";
          } else if (statOrMisc.indexOf("# of extra Windfury Attacks") > -1) {
            var lowerlevelwindfuryprocs = getUsesForDebuffSpellId("10610", buffsData) + getUsesForDebuffSpellId("8516", buffsData) + getUsesForDebuffSpellId("10608", buffsData) + getUsesForDebuffSpellId("25583", buffsData);
            var maxlevelwindfuryprocs = getUsesForDebuffSpellId("25584", buffsData);
            if ((maxlevelwindfuryprocs + lowerlevelwindfuryprocs) > 0 && Math.round(lowerlevelwindfuryprocs * 100 / (maxlevelwindfuryprocs + lowerlevelwindfuryprocs)) > 50)
              sheet.getRange(statsAndMisc.getRow() + statOrMiscCount + 1, statsAndMisc.getColumn() + playerDoneCount + classDoneCount + 1, 1, 1).setFontWeight("bold").setFontStyle("italic").setFontColor("#980000");
            if (playerByNameAsc.type == "Shaman" && lowerlevelwindfuryprocs + maxlevelwindfuryprocs > 0)
              sheet.getRange(statsAndMisc.getRow() + statOrMiscCount + 1, statsAndMisc.getColumn() + playerDoneCount + classDoneCount + 1, 1, 1).setFontWeight("bold").setFontStyle("italic").setFontColor("#986200");
            statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = lowerlevelwindfuryprocs + maxlevelwindfuryprocs;
          } else if (statOrMisc.indexOf("# of Battle Squawk buffs on bosses") > -1) {
            statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = getUsesForDebuffSpellId("23060", buffsData) - getUsesForDebuffSpellId("23060", buffsDataTrash);
          } else if (statOrMisc.indexOf("Critical heals done") > -1) {
            healingData.entries.forEach(function (healing, healingCount) {
              var increased = false;
              if (healing.targets != null && healing.targets.length > 0 && healing.actor != null && healing.actor == playerByNameAsc.id) {
                if (healing.critHitCount != null && healing.critHitCount > 0) {
                  amount += healing.critHitCount;
                  increased = true;
                }
              } else if (healing.subentries != null) {
                healing.subentries.forEach(function (subentry, subentryCount) {
                  if (subentry.targets != null && subentry.targets.length > 0 && subentry.actor != null && subentry.actor == playerByNameAsc.id) {
                    if (subentry.critHitCount != null && subentry.critHitCount > 0) {
                      amount += subentry.critHitCount;
                      increased = true;
                    }
                  }
                })
              }
              if (increased)
                totalHealingDoneUses += healing.hitCount;
            })
            if (amount > 0)
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + Math.round(amount * 1000 / totalHealingDoneUses) / 10 + "%)";
            else
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = "";
          } else if (statOrMisc.indexOf("outgoing") > -1) {
            var searchType = statOrMisc.split(" outgoing")[0];
            damageDoneData.entries.forEach(function (damage, damageCount) {
              if (damage.uses != null && damage.uses > 0) {
                var increased = false;
                if (damage.targets != null && damage.targets.length > 0 && damage.actor != null && damage.actor == playerByNameAsc.id) {
                  if (statOrMisc.indexOf("Critical") > -1) {
                    if (damage.critHitCount != null && damage.critHitCount > 0) {
                      amount += damage.critHitCount;
                      increased = true;
                    }
                  } else if (damage.missdetails != null) {
                    damage.missdetails.forEach(function (missdetail, missdetailCount) {
                      if (missdetail.type != null && missdetail.type.indexOf(searchType) > -1) {
                        amount += missdetail.count;
                        increased = true;
                      }
                    })
                  }
                } else if (damage.subentries != null) {
                  damage.subentries.forEach(function (subentry, subentryCount) {
                    if (subentry.targets != null && subentry.targets.length > 0 && subentry.actor != null && subentry.actor == playerByNameAsc.id) {
                      if (statOrMisc.indexOf("Critical") > -1) {
                        if (subentry.critHitCount != null && subentry.critHitCount > 0) {
                          amount += subentry.critHitCount;
                          increased = true;
                        }
                      } else if (subentry.missdetails != null) {
                        subentry.missdetails.forEach(function (missdetail, missdetailCount) {
                          if (missdetail.type != null && missdetail.type.indexOf(searchType) > -1) {
                            amount += missdetail.count;
                            increased = true;
                          }
                        })
                      }
                    }
                  })
                }
                if (increased) {
                  if (damage.hitCount == 0 && (damage.missCount != null || statOrMisc.indexOf("Critical") > -1) && damage.uses >= damage.missCount)
                    totalDamageDoneUses += damage.uses;
                  else
                    totalDamageDoneUses += damage.hitCount;
                  if (damage.missCount != null && damage.missCount > 0)
                    totalDamageDoneUses += damage.missCount;
                }
              }
            })
            if (amount > 0)
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + Math.round(amount * 1000 / totalDamageDoneUses) / 10 + "%)";
            else
              statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = "";
          } else if (statOrMisc.indexOf("incoming") > -1) {
            var searchType = statOrMisc.split(" incoming")[0];
            if (statOrMisc.indexOf("Dodge") > -1 || statOrMisc.indexOf("Immune") > -1 || statOrMisc.indexOf("Miss") > -1 || statOrMisc.indexOf("Parry") > -1) {
              damageTakenTotalData.entries.forEach(function (damage, damageCount) {
                if (damage.guid == 1) {
                  if (damage.sources != null && damage.sources.length > 0) {
                    if (damage.missdetails != null) {
                      damage.missdetails.forEach(function (missdetail, missdetailCount) {
                        if (missdetail.type != null && missdetail.type.indexOf(searchType) > -1)
                          amount += missdetail.count;
                      })
                    }
                  } else if (damage.subentries != null) {
                    damage.subentries.forEach(function (subentry, subentryCount) {
                      if (subentry.sources != null && subentry.sources.length > 0) {
                        if (subentry.missdetails != null) {
                          subentry.missdetails.forEach(function (missdetail, missdetailCount) {
                            if (missdetail.type != null && missdetail.type.indexOf(searchType) > -1)
                              amount += missdetail.count;
                          })
                        }
                      }
                    })
                  }
                  totalDamageTakenUses += damage.hitCount;
                  if (damage.missCount != null && damage.missCount > 0)
                    totalDamageTakenUses += damage.missCount;
                }
              })
              if (amount > 0)
                statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + Math.round(amount * 1000 / totalDamageTakenUses) / 10 + "%)";
              else
                statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = "";
            }
            if (statOrMisc.indexOf("Crushing") > -1 || statOrMisc.indexOf("Critical") > -1 || statOrMisc.indexOf("Blocked") > -1 || statOrMisc.indexOf("Resist") > -1) {
              damageTakenTotalData.entries.forEach(function (damage, damageCount) {
                if (damage.guid == 1) {
                  if (damage.sources != null && damage.sources.length > 0) {
                    if (damage.hitdetails != null) {
                      damage.hitdetails.forEach(function (hitdetail, missdetailCount) {
                        if (hitdetail.type != null && hitdetail.type.indexOf(searchType) > -1)
                          amount += hitdetail.count;
                      })
                    }
                  } else if (damage.subentries != null) {
                    damage.subentries.forEach(function (subentry, subentryCount) {
                      if (subentry.sources != null && subentry.sources.length > 0) {
                        if (subentry.hitdetails != null) {
                          subentry.hitdetails.forEach(function (hitdetail, missdetailCount) {
                            if (hitdetail.type != null && hitdetail.type.indexOf(searchType) > -1)
                              amount += hitdetail.count;
                          })
                        }
                      }
                    })
                  }
                  totalDamageTakenUses += damage.hitCount;
                  if (damage.missCount != null && damage.missCount > 0)
                    totalDamageTakenUses += damage.missCount;
                }
              })
              if (amount > 0)
                statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = amount + " (" + Math.round(amount * 1000 / totalDamageTakenUses) / 10 + "%)";
              else
                statsAndMiscArr[statOrMiscCount + 1][playerDoneCount + classDoneCount + 1] = "";
            }
          }
        })

        //fill in trinkets and racials
        if (playerDoneCount == 0) {
          copyRowStyles(conf, sheet, confTrinketsAndRacialsToTrack, trinketsAndRacialsToTrack.length, trinketsAndRacials.getRow() + 1, trinketsAndRacials.getColumn(), maxColumns, true, "right", darkMode);

          var confColumnWidth = conf.getColumnWidth(confTrinketsAndRacialsToTrack.getColumn());
          if (confColumnWidth > maxColumnWidth) {
            maxColumnWidth = confColumnWidth;
          }
        }
        if (!onlyTrash) {
          sheet.getRange(trinketsAndRacials.getRow(), trinketsAndRacials.getColumn() + playerDoneCount + classDoneCount + 1, 1, 1).setNote(getStringForLang("trinketsUsed", langKeys, langTrans, "", "", "", "") + ":\r\n" + trinketsUsed.substr(0, trinketsUsed.length - 2));
        }
        trinketsAndRacialsToTrack.forEach(function (trinketOrRacial, trinketOrRacialCount) {
          var amount = 0;
          var uptime = 0;
          var cdMultiplesOfActive = 0;
          if (trinketOrRacial.indexOf("--") > -1 && trinketOrRacial.indexOf("++") > -1 && (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)))
            cdMultiplesOfActive = trinketOrRacial.split("--")[1] * 100 / trinketOrRacial.split("++")[1];
          if (trinketsAndRacialsArr[trinketOrRacialCount + 1] == null || trinketsAndRacialsArr[trinketOrRacialCount + 1].length == 0) {
            trinketsAndRacialsArr[trinketOrRacialCount + 1] = [];
            trinketsAndRacialsArr[trinketOrRacialCount + 1].push(trinketsAndRacialsToTrackLang[trinketOrRacialCount].split(" [")[0].split(" {")[0]);
          }
          if (trinketOrRacial.indexOf("Arcanite Dragonling") > -1) {
            amount += usedArcaniteDragonling;
          } else if (trinketOrRacial.indexOf("Spell Vulnerability") > -1 && !onlyTrash) {
            debuffsAppliedDataBosses.auras.forEach(function (debuff, debuffCount) {
              if (debuff.guid.toString() == trinketOrRacial.split("[")[1].split("]")[0]) {
                amount += debuff.bands.length;
                uptime = Math.round(debuff.totalUptime * 100 / totalTimeElapsedRaw);
              }
            })
          } else {
            var checkAura = false;
            if (trinketOrRacial.indexOf("{") > -1 && trinketOrRacial.split("{")[1].split("}")[0] == "true") {
              checkAura = true;
            }
            if (checkAura) {
              buffsData.auras.forEach(function (spell, spellCount) {
                if (trinketOrRacial.indexOf("[") > -1) {
                  trinketOrRacial.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                    if (spellId == spell.guid.toString()) {
                      if (cdMultiplesOfActive > 0) {
                        if (Number(trinketOrRacial.split("++")[1]) > 0) {
                          var lastSpellBandEnd = 0;
                          spell.bands.forEach(function (spellBand, spellBandCount) {
                            if (lastSpellBandEnd == 0 || ((Number(lastSpellBandEnd) + Number(trinketOrRacial.split("++")[1]) * 1000) < spellBand.endTime))
                              amount += 1;
                            lastSpellBandEnd = spellBand.endTime;
                          })
                        }
                        else
                          amount += spell.bands.length;
                        uptime = Math.round((spell.totalUptime * 100 / Math.abs(raidDuration) * cdMultiplesOfActive) / 100);
                      } else {
                        amount += spell.bands.length;
                        uptime = Math.round(spell.totalUptime * 100 / totalTimeElapsedRaw);
                      }
                    }
                  })
                }
              })
            } else {
              playerData.entries.forEach(function (spell, spellCount) {
                if (trinketOrRacial.indexOf("[") > -1) {
                  trinketOrRacial.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                    if (spellId == spell.guid.toString()) {
                      amount += spell.total;
                      if (cdMultiplesOfActive > 0) {
                        uptime = Math.round((spell.uptime * 100 / Math.abs(raidDuration) * cdMultiplesOfActive) / 100);
                      }
                      else
                        uptime = Math.round(spell.uptime * 100 / totalTimeElapsedRaw);
                    }
                  })
                }
              })
            }
          }
          if (previousClass != playerByNameAsc.type)
            trinketsAndRacialsArr[trinketOrRacialCount + 1].push("");
          if (amount == 0) {
            if (trinketOrRacial.indexOf("Berserking") > -1)
              trinketsAndRacialsArr[trinketOrRacialCount + 1].push("");
            else if (trinketOrRacial.indexOf("Blood Fury") > -1)
              trinketsAndRacialsArr[trinketOrRacialCount + 1].push("");
            else if (trinketOrRacial.indexOf("Arcane Torrent") > -1)
              trinketsAndRacialsArr[trinketOrRacialCount + 1].push("");
            else if (trinketOrRacial.indexOf("War Stomp") > -1)
              trinketsAndRacialsArr[trinketOrRacialCount + 1].push("");
            else if (trinketOrRacial.indexOf("Will of the Forsaken") > -1)
              trinketsAndRacialsArr[trinketOrRacialCount + 1].push("");
            else if (trinketOrRacial.indexOf("Arcanite Dragonling") > -1)
              trinketsAndRacialsArr[trinketOrRacialCount + 1].push("----------");
            else
              trinketsAndRacialsArr[trinketOrRacialCount + 1].push("");
          }
          else {
            if (trinketOrRacial.indexOf("Arcanite Dragonling") > -1)
              trinketsAndRacialsArr[trinketOrRacialCount + 1].push("dmg done");
            else {
              if (uptime > 0 && (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)))
                trinketsAndRacialsArr[trinketOrRacialCount + 1].push(amount + " (" + uptime + "%)");
              else {
                if (trinketOrRacial.indexOf("--") > -1 && (!onlyBosses && !onlyTrash && (onlyFightNr == null || onlyFightNr.toString().length == 0)))
                  trinketsAndRacialsArr[trinketOrRacialCount + 1].push(amount + " (" + Math.round(amount * 100 / (Math.floor(Math.abs(raidDuration) / 1000 / trinketOrRacial.split("--")[1]) + 1)) + "%)");
                else
                  trinketsAndRacialsArr[trinketOrRacialCount + 1].push(amount);
              }
            }
          }
        })

        //fill in engineering stuff
        var oilOfImmolationCount = 0;
        if (playerDoneCount == 0) {
          copyRowStyles(conf, sheet, confEngineeringToTrack, engineeringToTrack.length, engineering.getRow() + 1, engineering.getColumn(), maxColumns, true, "right", darkMode);

          var confColumnWidth = conf.getColumnWidth(confEngineeringToTrack.getColumn());
          if (confColumnWidth > maxColumnWidth) {
            maxColumnWidth = confColumnWidth;
          }
        }
        engineeringToTrack.forEach(function (engineeringCast, engineeringCastCount) {
          var amount = 0;
          var hitCount = 0;
          if (engineeringArr[engineeringCastCount + 1] == null || engineeringArr[engineeringCastCount + 1].length == 0) {
            engineeringArr[engineeringCastCount + 1] = [];
            engineeringArr[engineeringCastCount + 1].push(engineeringToTrackLang[engineeringCastCount].split(" [")[0]);
          }
          playerData.entries.forEach(function (spell, spellCount) {
            if (engineeringCast.indexOf("[") > -1) {
              engineeringCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                if (spellId == spell.guid.toString()) {
                  amount += spell.total;
                  damageDoneData.entries.forEach(function (damDone, damDoneCount) {
                    if (damDone != null && damDone.guid != null && damDone.guid == spell.guid) {
                      hitCount += damDone.hitCount + damDone.missCount;
                      if (damDone.missCount != null && damDone.missCount > 0)
                        hitCount += damDone.missCount;
                    }
                  })
                }
              })
            }
          })
          if (previousClass != playerByNameAsc.type)
            engineeringArr[engineeringCastCount + 1].push("");
          if (amount == 0)
            engineeringArr[engineeringCastCount + 1].push("");
          else if (engineeringCast.indexOf("Immolation") < 0 && engineeringCast.indexOf("Dummy") < 0)
            engineeringArr[engineeringCastCount + 1].push(amount.toString() + " (⌀" + Math.round(hitCount / amount) + ")");
          else {
            if (engineeringCast.indexOf("Immolation") > -1)
              oilOfImmolationCount = amount;
            engineeringArr[engineeringCastCount + 1].push(amount);
          }
        })
        var totalEngineeringDamage = 0;
        if (showEngineeringDmg) {
          if (engineeringArr.length == engineeringToTrack.length + 1) {
            addSingleEntryToMultiDimArray(engineeringArr, getStringForLang("damageDoneEngi", langKeys, langTrans, "", "", "", ""));
            copyRangeStyle(confShowEngineeringDmg, sheet.getRange(engineering.getRow() + engineeringArr.length - 1, engineering.getColumn() + 1, 1, maxColumns), null, "center", null);
            copyRangeStyle(confShowOilOfImmolationDmg, sheet.getRange(engineering.getRow() + engineeringArr.length - 1, engineering.getColumn(), 1, 1), null, "right", null);
            if (darkMode)
              sheet.getRange(engineering.getRow() + engineeringArr.length - 1, engineering.getColumn(), 1, 1).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
          }
          if (previousClass != playerByNameAsc.type) {
            engineeringArr[engineeringToTrack.length + 1].push("");
          }
          if (damageTakenEngineeringData.entries.length > 0) {
            damageTakenEngineeringData.entries.forEach(function (damageTakenEngineeringDataEntry, damageTakenEngineeringDataEntryCount) {
              if (damageTakenEngineeringDataEntry.id == playerByNameAsc.id)
                totalEngineeringDamage = totalEngineeringDamage + damageTakenEngineeringDataEntry.total;
            })
            if (totalEngineeringDamage > 0) {
              engineeringArr[engineeringToTrack.length + 1].push(totalEngineeringDamage);
              sheet.getRange(engineering.getRow() + engineeringToTrack.length + 1, engineering.getColumn() + playerCount + classDoneCount + 1, 1, 1).setNote(getStringForLang("dps", langKeys, langTrans, "", "", "", "") + ": " + Math.round(totalEngineeringDamage * 10000 / raidDuration) / 10);
            }
            else
              engineeringArr[engineeringToTrack.length + 1].push("");
          } else
            engineeringArr[engineeringToTrack.length + 1].push("");
        }

        var totalOilOfImmolationDamage = 0;
        if (showOilOfImmolationDmg) {
          if (engineeringArr.length == engineeringToTrack.length + 1 + bonusEngi) {
            addSingleEntryToMultiDimArray(engineeringArr, getStringForLang("damageDoneImmolation", langKeys, langTrans, "", "", "", ""));
            copyRangeStyle(confShowOilOfImmolationDmg, sheet.getRange(engineering.getRow() + engineeringArr.length - 1, engineering.getColumn() + 1, 1, maxColumns), null, "center", null);
            if (darkMode)
              sheet.getRange(engineering.getRow() + engineeringArr.length - 1, engineering.getColumn() + 1, 1, maxColumns).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID)
            copyRangeStyle(confShowOilOfImmolationDmg, sheet.getRange(engineering.getRow() + engineeringArr.length - 1, engineering.getColumn(), 1, 1), null, "right", null);
            if (darkMode)
              sheet.getRange(engineering.getRow() + engineeringArr.length - 1, engineering.getColumn(), 1, 1).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
          }
          if (previousClass != playerByNameAsc.type)
            engineeringArr[engineeringArr.length - 1].push("");
          if (damageTakenOilOfImmoData.entries.length > 0) {
            damageTakenOilOfImmoData.entries.forEach(function (damageTakenOilOfImmoDataEntry, damageTakenOilOfImmoDataEntryCount) {
              if (damageTakenOilOfImmoDataEntry.id == playerByNameAsc.id)
                totalOilOfImmolationDamage = totalOilOfImmolationDamage + damageTakenOilOfImmoDataEntry.total;
            })
            if (totalOilOfImmolationDamage > 0) {
              engineeringArr[engineeringArr.length - 1].push(totalOilOfImmolationDamage);
              sheet.getRange(engineering.getRow() + engineeringToTrack.length + 2, engineering.getColumn() + classDoneCount + playerCount + 1, 1, 1).setNote(getStringForLang("dps", langKeys, langTrans, "", "", "", "") + ": " + Math.round(totalOilOfImmolationDamage * 10000 / raidDuration) / 10);
            } else {
              if (oilOfImmolationCount > 0)
                engineeringArr[engineeringArr.length - 1].push(0);
              else
                engineeringArr[engineeringArr.length - 1].push("");
            }
          } else {
            if (oilOfImmolationCount > 0)
              engineeringArr[engineeringArr.length - 1].push(0);
            else
              engineeringArr[engineeringArr.length - 1].push("");
          }
        }
        if (showEngineeringDmg) {

        }

        //fill in other casts
        if (playerDoneCount == 0) {
          copyRowStyles(conf, sheet, confOtherCastsToTrack, otherCastsToTrack.length, otherCasts.getRow() + 1, otherCasts.getColumn(), maxColumns, true, "right", darkMode);

          var confColumnWidth = conf.getColumnWidth(confOtherCastsToTrack.getColumn());
          if (confColumnWidth > maxColumnWidth) {
            maxColumnWidth = confColumnWidth;
          }
        }
        otherCastsToTrack.forEach(function (otherCast, otherCastCount) {
          var amount = 0;
          var debuffIdString = otherCast.split("[")[1].split("]")[0];
          var otherCastString = "";
          if (otherCastsArr[otherCastCount + 1] == null || otherCastsArr[otherCastCount + 1].length == 0) {
            otherCastsArr[otherCastCount + 1] = [];
            otherCastString = otherCastsToTrackLang[otherCastCount].split(" [")[0].split(" {")[0];
            if (!onlyTrash && otherCast.indexOf("uptime") > -1 && otherCast.indexOf("total") < 0 && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
              otherCastString = otherCastString.replace("%)", "% - " + getStringForLang("overall", langKeys, langTrans, "", "", "", "") + ": " + getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedDataBossesTotal, totalTimeElapsedBosses) + "%)");
            } else if (otherCast.indexOf("uptime") > -1) {
              otherCastString = otherCastString.replace("%)", "% - " + getStringForLang("overall", langKeys, langTrans, "", "", "", "") + ": " + getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedDataTotal, totalTimeElapsedRaw) + "%)");
            }
            otherCastsArr[otherCastCount + 1].push(otherCastString);
          }
          if (Number(otherCast.split("++")[1]) > 0) {
            buffsData.auras.forEach(function (spell, spellCount) {
              if (otherCast.indexOf("[") > -1) {
                otherCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                  if (spellId == spell.guid.toString()) {
                    var lastSpellBandEnd = 0;
                    spell.bands.forEach(function (spellBand, spellBandCount) {
                      if (lastSpellBandEnd == 0 || ((Number(lastSpellBandEnd) + (Number(otherCast.split("++")[1]) * 1000)) < spellBand.endTime))
                        amount += 1;
                      lastSpellBandEnd = spellBand.endTime;
                    })
                  }
                })
              }
            })
          }
          var castAmount = 0;
          playerData.entries.forEach(function (spell, spellCount) {
            if (otherCast.indexOf("[") > -1) {
              otherCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                if (spellId == spell.guid.toString()) {
                  castAmount += spell.total;
                }
              })
            }
          })
          if (castAmount > amount)
            amount = castAmount;
          if (previousClass != playerByNameAsc.type)
            otherCastsArr[otherCastCount + 1].push("");
          if (otherCast.indexOf("Gift of Arthas") > -1) {
            if (!onlyTrash && otherCast.indexOf("uptime") > -1 && otherCast.indexOf("total") < 0 && (onlyFightNr == null || onlyFightNr.toString().length == 0)) {
              var amount = getAmountForDebuffSpellId(debuffIdString, debuffsAppliedDataBosses, totalTimeElapsedBosses);
              if (amount > 0)
                otherCastsArr[otherCastCount + 1].push(amount + " (" + getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedDataBosses, totalTimeElapsedBosses) + "%)");
              else
                otherCastsArr[otherCastCount + 1].push("");
            } else {
              var amount = getAmountForDebuffSpellId(debuffIdString, debuffsAppliedData, totalTimeElapsedRaw);
              if (amount > 0)
                otherCastsArr[otherCastCount + 1].push(amount + " (" + getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedData, totalTimeElapsedRaw) + "%)");
              else
                otherCastsArr[otherCastCount + 1].push("");
            }
          } else {
            if (amount == 0)
              otherCastsArr[otherCastCount + 1].push("");
            else
              otherCastsArr[otherCastCount + 1].push(amount);
          }
        })
        if (showUsedTemporaryWeaponEnchant) {
          if (otherCastsArr.length == otherCastsToTrack.length + 1) {
            addSingleEntryToMultiDimArray(otherCastsArr, getStringForLang("temporaryWeaponEnhancement", langKeys, langTrans, "", "", "", ""));
            copyRangeStyle(confShowOilOfImmolationDmg, sheet.getRange(otherCasts.getRow() + otherCastsArr.length - 1, otherCasts.getColumn() + 1, 1, maxColumns), null, "center", null);
            if (darkMode)
              sheet.getRange(otherCasts.getRow() + otherCastsArr.length - 1, otherCasts.getColumn() + 1, 1, maxColumns).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
            copyRangeStyle(confShowOilOfImmolationDmg, sheet.getRange(otherCasts.getRow() + otherCastsArr.length - 1, otherCasts.getColumn(), 1, 1), null, "right", null);
            if (darkMode)
              sheet.getRange(otherCasts.getRow() + otherCastsArr.length - 1, otherCasts.getColumn(), 1, 1).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
          }
          if (previousClass != playerByNameAsc.type)
            otherCastsArr[otherCastsArr.length - 1].push("");
          otherCastsArr[otherCastsArr.length - 1].push(usedTemporaryWeaponEnchant);
        }

        //fill in absorbs
        var totalAmount = 0;
        if (playerDoneCount == 0) {
          copyRowStyles(conf, sheet, confAbsorbsToTrack, absorbsToTrack.length, absorbs.getRow() + 1, absorbs.getColumn(), maxColumns, true, "right", darkMode);

          var confColumnWidth = conf.getColumnWidth(confAbsorbsToTrack.getColumn());
          if (confColumnWidth > maxColumnWidth) {
            maxColumnWidth = confColumnWidth;
          }
        }
        absorbsToTrack.forEach(function (absorbCast, absorbCastCount) {
          var amount = 0;
          if (absorbsArr[absorbCastCount + 1] == null || absorbsArr[absorbCastCount + 1].length == 0) {
            absorbsArr[absorbCastCount + 1] = [];
            if (absorbCast.indexOf("Power Word: Shield") > -1)
              absorbsArr[absorbCastCount + 1].push(absorbsToTrackLang[absorbCastCount].split(" [")[0] + " (" + getStringForLang("excludedFromAbsorbed", langKeys, langTrans, "", "", "", "") + ")");
            else
              absorbsArr[absorbCastCount + 1].push(absorbsToTrackLang[absorbCastCount].split(" [")[0]);
          }
          if (absorbCast.indexOf("Nature Absorption") < 0 && absorbCast.indexOf("Arcane Absorption") < 0 && absorbCast.indexOf("Fire Absorption") < 0 && absorbCast.indexOf("Frost Absorption") < 0 && absorbCast.indexOf("Shadow Absorption") < 0) {
            healingData.entries.forEach(function (spell, spellCount) {
              if (absorbCast.indexOf("[") > -1) {
                absorbCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                  if (spellId == spell.guid.toString()) {
                    amount += spell.total;
                  }
                })
              }
            })
          } else {
            healingDataTarget.entries.forEach(function (spell, spellCount) {
              if (absorbCast.indexOf("[") > -1) {
                absorbCast.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
                  if (spellId == spell.guid.toString()) {
                    amount += spell.total;
                  }
                })
              }
            })
          }
          if (previousClass != playerByNameAsc.type)
            absorbsArr[absorbCastCount + 1].push("");
          if (amount == 0)
            absorbsArr[absorbCastCount + 1].push("");
          else
            absorbsArr[absorbCastCount + 1].push(amount);
          if (absorbCast.indexOf("Power Word: Shield") < 0)
            totalAmount += amount;
        })
        if (absorbsArr.length == absorbsToTrack.length + 1) {
          addSingleEntryToMultiDimArray(absorbsArr, getStringForLang("absorbedTot", langKeys, langTrans, "", "", "", ""));
          sheet.getRange(absorbs.getRow() + absorbsArr.length - 1, absorbs.getColumn(), 1, maxColumns + 1).setFontStyle("italic");
          sheet.getRange(absorbs.getRow() + absorbsArr.length - 1, absorbs.getColumn(), 1, 1).setHorizontalAlignment("right");
          copyRangeStyle(confTotalAndInformationRowsDefaultTemplate, sheet.getRange(absorbs.getRow() + absorbsArr.length - 1, absorbs.getColumn() + 1, 1, maxColumns).setFontStyle("italic"), null, "center", null);
        }
        if (previousClass != playerByNameAsc.type)
          absorbsArr[absorbsArr.length - 1].push("");
        absorbsArr[absorbsArr.length - 1].push(totalAmount);

        //fill in interrupts)
        if (showInterruptedSpells) {
          var interruptedTotal = 0;
          var interruptedString = "";
          interruptedData.entries[0].entries.forEach(function (spell, spellCount) {
            if (spell != null && spell.guid != null) {
              var targets = "";
              if (spell.details != null) {
                spell.details.forEach(function (spellDetail, spellDetailCount) {
                  if (spellDetail.id == playerByNameAsc.id) {
                    spellDetail.actors.forEach(function (target, targetCount) {
                      var targetNameStripped = target.name.replace(/\s\d+/g, '');
                      if (targets.indexOf(targetNameStripped) < 0) {
                        if (targets == "")
                          targets += targetNameStripped;
                        else
                          targets += targetNameStripped;
                      }
                    })
                    if (interruptedTotal == 0)
                      interruptedString += spell.name + " (" + targets + ")";
                    else
                      interruptedString += ", " + spell.name + " (" + targets + ")";
                    interruptedTotal += spellDetail.total;
                  }
                })
              }
            }
          })
          if (interruptsArr.length == 1) {
            addSingleEntryToMultiDimArray(interruptsArr, getStringForLang("interruptedSpells", langKeys, langTrans, "", "", "", ""));
            sheet.getRange(interrupts.getRow() + 1, interrupts.getColumn(), 1, 1).setHorizontalAlignment("right").setFontSize(confShowInterruptedSpells.getFontSize()).setFontStyle(confShowInterruptedSpells.getFontStyle()).setFontWeight(confShowInterruptedSpells.getFontWeight());
            copyRangeStyle(confShowInterruptedSpells, sheet.getRange(interrupts.getRow() + 1, interrupts.getColumn() + 1, 1, maxColumns), null, "center", null);
            if (showInterruptedSpellsNamesRow) {
              addSingleEntryToMultiDimArray(interruptsArr, getStringForLang("namesAndSourcesInterruptedSpells", langKeys, langTrans, "", "", "", ""));
              sheet.getRange(interrupts.getRow() + 2, interrupts.getColumn(), 1, 1).setHorizontalAlignment("right").setFontSize(confShowInterruptedSpellsNamesRow.getFontSize()).setFontStyle(confShowInterruptedSpellsNamesRow.getFontStyle()).setFontWeight(confShowInterruptedSpellsNamesRow.getFontWeight());
              copyRangeStyle(confShowInterruptedSpellsNamesRow, sheet.getRange(interrupts.getRow() + 2, interrupts.getColumn() + 1, 1, maxColumns), null, "center", null);
              if (darkMode)
                sheet.getRange(interrupts.getRow() + 2, interrupts.getColumn() + 1, 1, maxColumns).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
            }
          }
          if (showInterruptedSpellsNamesRow) {
            if (previousClass != playerByNameAsc.type) {
              interruptsArr[interruptsArr.length - 2].push("");
              interruptsArr[interruptsArr.length - 1].push("");
            }
            if (interruptedTotal > 0)
              interruptsArr[interruptsArr.length - 2].push(interruptedTotal);
            else
              interruptsArr[interruptsArr.length - 2].push("");
            interruptsArr[interruptsArr.length - 1].push(interruptedString);
          }
          else {
            if (previousClass != playerByNameAsc.type) {
              interruptsArr[interruptsArr.length - 1].push("");
            }
            if (interruptedTotal > 0)
              interruptsArr[interruptsArr.length - 1].push(interruptedTotal);
            else
              interruptsArr[interruptsArr.length - 1].push("");
          }
          sheet.getRange(interrupts.getRow() + 2, interrupts.getColumn() + 1 + playerDoneCount + classDoneCount, 14, 1).merge();
        }

        //add player names to headers
        if (previousClass != playerByNameAsc.type) {
          rolesAndNames[0].push("");
          rolesAndNames[1].push("");
          rolesAndNames[2].push(getStringForLang(playerByNameAsc.type + "Plur", langKeys, langTrans, "", "", "", ""));
          sheet.getRange(singleTargetCasts.getRow() - 1, singleTargetCasts.getColumn() + playerDoneCount + classDoneCount, 1, 1).setFontWeight("bold").setHorizontalAlignment("left");
          sheet.getRange(singleTargetCasts.getRow() - 1, singleTargetCasts.getColumn() + playerDoneCount + classDoneCount + 1, 1, playersInThisClass).setFontWeight("bold").setHorizontalAlignment("center").setBackground(getColourForPlayerClass(playerByNameAsc.type)).setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
        }
        rolesAndNames[2].push(playerByNameAsc.name)
        adjustNameRow(sheet.getRange(singleTargetCasts.getRow() - 1, singleTargetCasts.getColumn() + playerDoneCount + classDoneCount + 1), playerByNameAsc.name, 1);
        var role1 = getRoleForPlayerClass(playerByNameAsc.type, langKeys, langTrans, dpsCount, tankCount, healerCount, dpsSpec);
        var role2 = "---";
        rolesAndNames[0].push(role1);
        rolesAndNames[1].push(role2);

        if (previousClass != playerByNameAsc.type) {
          previousClass = playerByNameAsc.type;
        }
        playerDoneCount++;
      }
    })
    if (maxColumnWidth > 335)
      sheet.setColumnWidth(2, maxColumnWidth);
    if (singleTargetCastsArr != null)
      singleTargetCasts.setValues(singleTargetCastsArr);
    if (aoeCastsArr != null)
      aoeCasts.setValues(aoeCastsArr);
    if (secondsActiveArr != null)
      secondsActive.setValues(fillUpMultiDimArrayWithEmptyValues(secondsActiveArr, maxColumns));
    damageTaken.setValues(fillUpMultiDimArrayWithEmptyValues(damageTakenArr, maxColumns));
    debuffs.setValues(fillUpMultiDimArrayWithEmptyValues(debuffsArr, maxColumns));
    if (classCooldownsArr != null)
      classCooldowns.setValues(classCooldownsArr);
    statsAndMisc.setValues(fillUpMultiDimArrayWithEmptyValues(statsAndMiscArr, maxColumns));
    trinketsAndRacials.setValues(fillUpMultiDimArrayWithEmptyValues(trinketsAndRacialsArr, maxColumns));
    engineering.setValues(fillUpMultiDimArrayWithEmptyValues(engineeringArr, maxColumns));
    otherCasts.setValues(fillUpMultiDimArrayWithEmptyValues(otherCastsArr, maxColumns));
    absorbs.setValues(fillUpMultiDimArrayWithEmptyValues(absorbsArr, maxColumns));
    if (showInterruptedSpells)
      interrupts.setValues(fillUpMultiDimArrayWithEmptyValues(interruptsArr, maxColumns));
    sheet.getRange(singleTargetCasts.getRow() - 3, singleTargetCasts.getColumn() + 1, rolesAndNames.length, rolesAndNames[0].length).setValues(rolesAndNames);

    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
      if (protections[i].getDescription() == 'removed after Start') {
        protections[i].remove();
      }
    }

    sheet.hideColumns(singleTargetCasts.getColumn() + playerDoneCount + classDoneCount + 1, maxColumns - singleTargetCasts.getColumn() - playerDoneCount - classDoneCount + 2);
    if (darkMode)
      sheet.getRange(4, 63).setFontColor("#d9d9d9").setValue("done");
    else
      sheet.getRange(4, 63).setFontColor("white").setValue("done");
    sheet.getRange(3, 5).setValue(getStringForLang("step6AftermathPartOne", langKeys, langTrans, "", "", "", ""));
    sheet.getRange(4, 5).setValue(getStringForLang("step6AftermathPartTwo", langKeys, langTrans, "", "", "", ""));
  } else
    SpreadsheetApp.getUi().alert(getStringForLang("noReportEntered", langKeys, langTrans, "", "", "", ""));
  try {
    var conf_del = ss.getSheetByName("configNew" + rnd);
    ss.deleteSheet(conf_del);
  } catch (err) { }
}