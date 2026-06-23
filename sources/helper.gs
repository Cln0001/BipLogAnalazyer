function getStringForLang(key, langkeys, langTrans, param1, param2, param3, param4) {
  if (langkeys.indexOf(key) > -1)
    return langTrans[langkeys.indexOf(key)].replace("<param1>", param1).replace("<param2>", param2).replace("<param3>", param3).replace("<param4>", param4);
  else {
    return "missing/fehlend/失踪/отсутствует";
  }
}

function nameSheetSafely(sheet, baseSheetName, langKeys, langTrans) {
  try {
    sheet.setName(baseSheetName);
  } catch (err) {
    try {
      baseSheetName += "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "");
      sheet.setName(baseSheetName);
    } catch (err2) {
      try {
        baseSheetName += "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "");
        sheet.setName(baseSheetName);
      } catch (err3) {
        baseSheetName += "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "");
        sheet.setName(baseSheetName);
      }
    }
  }
  return baseSheetName;
}

function checkIfArrayContainsEntry(array, entry) {
  var contains = false;
  if (array != null && array.length > 0) {
    for (var i = 0, j = array.length; i < j; i++) {
      if (array[i].toUpperCase() == entry.toUpperCase()) {
        contains = true;
        break;
      }
    }
  }
  return contains;
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

function cleanSheet(sheet, information, darkMode) {
  sheet.showColumns(1, sheet.getMaxColumns());
  shiftRangeByColumns(sheet, information, 1).clearContent().setNumberFormat("0");
  if (darkMode) {
    sheet.getRange(7, 2, sheet.getMaxRows() - 6, sheet.getMaxColumns() - 1).clearNote();
    sheet.getRange(7, 2, sheet.getMaxRows() - 6, sheet.getMaxColumns() - 1).breakApart().clearContent().setNumberFormat("0").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("left").setBackground("#d9d9d9").setFontSize(10).setFontColor("black").setFontWeight("normal").setFontStyle("normal").protect().setWarningOnly(true).setDescription("removed after Start");
    sheet.getRange(1, 1, 7, 1).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(8, 1, sheet.getMaxRows() - 7, 1).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(1, 2, 6, 1).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(1, 8, 1, 1).setFontColor("#d9d9d9");
    sheet.getRange(4, 64, 1, 1).setFontColor("#d9d9d9");
    sheet.getRange(1, 5, 4, sheet.getMaxColumns() - 4).setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(5, 3, 2, sheet.getMaxColumns() - 2).breakApart().clearContent().setNumberFormat("0").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("left").setBackground("#d9d9d9").setFontSize(10).setFontColor("black").setFontWeight("normal").setFontStyle("normal").protect().setWarningOnly(true).setDescription("removed after Start");
  } else {
    sheet.getRange(7, 2, sheet.getMaxRows() - 6, sheet.getMaxColumns() - 1).clearNote();
    sheet.getRange(7, 2, sheet.getMaxRows() - 6, sheet.getMaxColumns() - 1).breakApart().clearContent().setNumberFormat("0").setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("left").setBackground("white").setFontSize(10).setFontColor("black").setFontWeight("normal").setFontStyle("normal").protect().setWarningOnly(true).setDescription("removed after Start");
    sheet.getRange(1, 1, 7, 1).setBackground("white").setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(8, 1, sheet.getMaxRows() - 7, 1).setBackground("white").setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(1, 2, 6, 1).setBackground("white").setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(1, 8, 1, 1).setFontColor("white");
    sheet.getRange(4, 64, 1, 1).setFontColor("white");
    sheet.getRange(1, 5, 4, sheet.getMaxColumns() - 4).setBackground("white").setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(5, 3, 2, sheet.getMaxColumns() - 2).breakApart().clearContent().setNumberFormat("0").setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("left").setBackground("white").setFontSize(10).setFontColor("black").setFontWeight("normal").setFontStyle("normal").protect().setWarningOnly(true).setDescription("removed after Start");
  }
  sheet.getRange(1, 16).setBackground("#9a86d6");
  sheet.getRange(5, 3, 2, sheet.getMaxColumns() - 2).setFontSize(9).setHorizontalAlignment("center");
}

function copyConfigSheet(rnd, sourceSpreadsheet, targetSpreadsheet) {
  var confSheetName = "configNew";
  var confSheet = sourceSpreadsheet.getSheetByName(confSheetName);
  return confSheet.copyTo(targetSpreadsheet).hideSheet().setName(confSheetName + rnd);
}

function getAmountForDebuffSpellId(debuffIdString, debuffsAppliedTotal, totalTimeElapsed) {
  var totalAmount = 0;
  debuffsAppliedTotal.auras.forEach(function (debuffTotal, debuffTotalCount) {
    if (debuffTotal.guid.toString() == debuffIdString) {
      totalAmount = debuffTotal.totalUses;
    }
  })
  return totalAmount;
}

function getUptimeForDebuffSpellId(debuffIdString, debuffsAppliedTotal, totalTimeElapsed) {
  var totalUptime = 0;
  debuffsAppliedTotal.auras.forEach(function (debuffTotal, debuffTotalCount) {
    if (debuffTotal.guid.toString() == debuffIdString) {
      totalUptime = Math.round(debuffTotal.totalUptime * 100 / totalTimeElapsed);
    }
  })
  return totalUptime;
}

function getUsesForDebuffSpellId(debuffIdString, debuffsAppliedTotal) {
  var totalUses = 0;
  debuffsAppliedTotal.auras.forEach(function (debuffTotal, debuffTotalCount) {
    if (debuffTotal.guid.toString() == debuffIdString) {
      totalUses = debuffTotal.totalUses;
    }
  })
  return totalUses;
}

function copyRowStyles(conf, sheet, confRange, castsCount, startRow, startColumn, maxSupportedPlayers, firstColumnIsDefault, firstColumnAlign, darkMode) {
  var confCastsRange = addRowsToRange(conf, shiftRangeByRows(conf, confRange, 1), castsCount - 1);
  var tarCastsRange = sheet.getRange(startRow, startColumn, castsCount, 1);
  copyRangeStyle(confCastsRange, tarCastsRange, null, firstColumnAlign, null);
  if (firstColumnIsDefault) {
    if (darkMode)
      tarCastsRange.setBackground("#d9d9d9").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
    else
      tarCastsRange.setBackground("white").setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
  }
  copyRangeStyle(confCastsRange, addColumnsToRange(sheet, shiftRangeByColumns(sheet, tarCastsRange, 1), maxSupportedPlayers - 1), null, "center", null);
}

function copyRangeStyle(rangeSource, rangeTarget, bold, alignment, fontSize) {
  rangeSource.copyTo(rangeTarget, { formatOnly: true });
  if (bold != null && bold)
    rangeTarget.setFontWeight("bold");
  else
    rangeTarget.setFontWeight("normal");
  if (alignment != null)
    rangeTarget.setHorizontalAlignment(alignment);
  if (fontSize != null)
    rangeTarget.setFontSize(fontSize);
}

function addRowsToRange(sheet, range, rowsToAdd) {
  return sheet.getRange(range.getRow(), range.getColumn(), range.getNumRows() + rowsToAdd, range.getNumColumns());
}

function addColumnsToRange(sheet, range, columnsToAdd) {
  return sheet.getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns() + columnsToAdd);
}

function shiftRangeByRows(sheet, range, rowsToShift) {
  return sheet.getRange(range.getRow() + rowsToShift, range.getColumn(), range.getNumRows(), range.getNumColumns());
}

function shiftRangeByColumns(sheet, range, columnsToShift) {
  return sheet.getRange(range.getRow(), range.getColumn() + columnsToShift, range.getNumRows(), range.getNumColumns());
}

function addSingleEntryToMultiDimArray(multiArray, value) {
  multiArray[multiArray.length] = [];
  multiArray[multiArray.length - 1].push(value);
}

function rangesIntersect(R1, R2) {
  return (R1.getLastRow() >= R2.getRow()) && (R2.getLastRow() >= R1.getRow()) && (R1.getLastColumn() >= R2.getColumn()) && (R2.getLastColumn() >= R1.getColumn());
}

function getOutputRange(sheet, confRange, rowCount, columnCount) {
  var outputRangeBeginCell = sheet.getRange(confRange.getValue().split("[")[1].split("]")[0], 2);
  return sheet.getRange(outputRangeBeginCell.getRow(), outputRangeBeginCell.getColumn(), rowCount, columnCount);
}

function getHeaderFromConfig(confRange) {
  var headerConfText = confRange.getValue();
  return headerConfText.indexOf("{") > -1 ? headerConfText.split("{")[1].split("}")[0] : "";
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function isRangeShowingPlayerNames(confRange) {
  if (confRange != null) {
    var headerConfText = confRange.getValue();
    if (headerConfText.indexOf("--showPlayerNames--") > -1)
      return true;
    else
      return false;
  } else {
    return false;
  }
}

function replaceAll(str, find, replace) {
  return str.replace(new RegExp(find, 'g'), replace);
}

function hideEmptyColumns(sheet, range, columnsToHideCount) {
  if (!sheet.isColumnHiddenByUser(range.getColumn()) && range.getValue().length < 1)
    sheet.hideColumns(range.getColumn(), columnsToHideCount);
}

function adjustNameRow(range, name, adjustment) {
  adjustFontSizeForPlayerNames(range, name, adjustment);
}

function adjustFontSizeForPlayerNames(range, name, adjustment) {
  var stringLength = name.length;
  if (stringLength > 6 && stringLength <= 8)
    range.setFontSize(11 - adjustment);
  else if (stringLength > 8 && stringLength <= 11)
    range.setFontSize(9 - adjustment);
  else if (stringLength > 11)
    range.setFontSize(8 - adjustment);
  else if (stringLength > 0)
    range.setFontSize(12 - adjustment);
}

function fillUpMultiDimArrayWithEmptyValues(array, lengthOfRow) {
  for (var j = 0; j < array.length; j++) {
    for (var i = array[j].length; i <= lengthOfRow; i++) {
      array[j].push("");
    }
  }
  return array;
}

function fillUpMultiDimArrayWithEmptyValuesWithNumberOfRows(array, lengthOfRow, numberOfRows) {
  for (var j = 0; j < numberOfRows; j++) {
    if (array[j] == null)
      array[j] = [];
    for (var i = array[j].length; i <= lengthOfRow; i++) {
      array[j].push("");
    }
  }
  return array;
}

function isAbilityTrackedById(abilityId, abilitiesToTrack) {
  var found = "";
  abilitiesToTrack.forEach(function (ability, abilityCount) {
    if (ability.indexOf("[") > -1) {
      ability.split("[")[1].split("]")[0].split(",").forEach(function (spellId, spellIdCount) {
        if (abilityId != null && abilityId.toString().length > 1 && spellId == abilityId.toString()) {
          found = ability;
        }
      })
    }
  })
  return found;
}

function isAbilityTrackedByName(abilityName, abilitiesToTrack) {
  var found = "";
  abilitiesToTrack.forEach(function (ability, abilityCount) {
    if (abilityName != null && abilityName.toString().length > 1 && ability.toString().indexOf(abilityName.split(" [")[0]) > -1)
      found = ability;
  })
  return found;
}

function shortenSources(sources) {
  var shortenedSources = "";
  sources.split("/").filter(function () { return true }).forEach(function (source, sourceCount) {
    if (source.length > 1) {
      var parts = source.split(" ");
      if (parts.length == 1)
        shortenedSources += parts.map((val, index, arr) => (index == 0) ? val.substring(0, 4) + '.' : val).join(" ") + "/";
      else
        shortenedSources += parts.map((val, index, arr) => (index > 0) ? val.charAt(0) + '.' : val).join(" ") + "/";
    }
  })

  if (shortenedSources.endsWith("."))
    shortenedSources.substring(0, shortenedSources.length - 1);
  if (shortenedSources.length > 40) {
    var shortenedSourcesParts = shortenedSources.split("/");
    var evenMoreShortenedSources = "";
    for (var i = 0, j = shortenedSourcesParts.length; i < j; i++) {
      if (evenMoreShortenedSources.length + shortenedSourcesParts[i].length + 1 > 40) {
        evenMoreShortenedSources += "....";
        break;
      } else
        evenMoreShortenedSources += shortenedSourcesParts[i] + "/";
    }
    shortenedSources = evenMoreShortenedSources;
  }
  if (shortenedSources.endsWith("/"))
    shortenedSources.substring(0, shortenedSources.length - 1);
  return shortenedSources;
}

function sortByProperty(objArray, prop, direction) {
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

function getRoleForPlayerClass(playerClass, langKeys, langTrans, dpsCount, tankCount, healerCount, dpsSpec) {
  if (playerClass == "Druid") {
    if (healerCount >= tankCount && healerCount >= dpsCount)
      return getStringForLang("Healer", langKeys, langTrans, "", "", "", "");
    else if (tankCount >= dpsCount && tankCount >= healerCount)
      return getStringForLang("Tank", langKeys, langTrans, "", "", "", "");
    else if (dpsCount >= tankCount && dpsCount >= healerCount) {
      if (dpsSpec == "Balance")
        return getStringForLang("Caster", langKeys, langTrans, "", "", "", "");
      else
        return getStringForLang("Physical", langKeys, langTrans, "", "", "", "");
    }
  } else if (playerClass == "Hunter") {
    return getStringForLang("Physical", langKeys, langTrans, "", "", "", "");
  } else if (playerClass == "Mage") {
    return getStringForLang("Caster", langKeys, langTrans, "", "", "", "");
  } else if (playerClass == "Paladin") {
    if (healerCount >= tankCount && healerCount >= dpsCount)
      return getStringForLang("Healer", langKeys, langTrans, "", "", "", "");
    else if (tankCount >= dpsCount && tankCount >= healerCount)
      return getStringForLang("Tank", langKeys, langTrans, "", "", "", "");
    else if (dpsCount >= tankCount && dpsCount >= healerCount)
      return getStringForLang("Physical", langKeys, langTrans, "", "", "", "");
  } else if (playerClass == "Priest") {
    if (dpsCount >= tankCount && dpsCount >= healerCount)
      return getStringForLang("Caster", langKeys, langTrans, "", "", "", "");
    else
      return getStringForLang("Healer", langKeys, langTrans, "", "", "", "");
  } else if (playerClass == "Rogue") {
    return getStringForLang("Physical", langKeys, langTrans, "", "", "", "");
  } else if (playerClass == "Shaman") {
    if (healerCount >= tankCount && healerCount >= dpsCount)
      return getStringForLang("Healer", langKeys, langTrans, "", "", "", "");
    else if (tankCount >= dpsCount && tankCount >= healerCount)
      return getStringForLang("Tank", langKeys, langTrans, "", "", "", "");
    else if (dpsCount >= tankCount && dpsCount >= healerCount) {
      if (dpsSpec == "Elemental")
        return getStringForLang("Caster", langKeys, langTrans, "", "", "", "");
      else
        return getStringForLang("Physical", langKeys, langTrans, "", "", "", "");
    }
  } else if (playerClass == "Warlock") {
    return getStringForLang("Caster", langKeys, langTrans, "", "", "", "");
  } else if (playerClass == "Warrior") {
    if (dpsCount >= tankCount && dpsCount >= healerCount)
      return getStringForLang("Physical", langKeys, langTrans, "", "", "", "");
    else
      return getStringForLang("Tank", langKeys, langTrans, "", "", "", "");
  }
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
}

const BatchOperationsManager = {
  operations: [],

  // Add a setValue operation to the queue
  addSetValue(range, value) {
    this.operations.push({
      type: 'setValue',
      range: range,
      value: value
    });
  },

  // Add a setValues operation to the queue
  addSetValues(range, values) {
    this.operations.push({
      type: 'setValues',
      range: range,
      values: values
    });
  },

  // Add formatting operations to the queue
  addFormat(range, formats) {
    this.operations.push({
      type: 'format',
      range: range,
      formats: formats
    });
  },

  // Add a note to the queue
  addNote(range, note) {
    this.operations.push({
      type: 'note',
      range: range,
      note: note
    });
  },
  addProtection(range, description, warningOnly = true) {
    this.operations.push({
      type: 'protection',
      range: range,
      description: description,
      warningOnly: warningOnly
    });
  },

  addHideColumns(sheet, startColumn, numColumns) {
    this.operations.push({
      type: 'hideColumns',
      sheet: sheet,
      startColumn: startColumn,
      numColumns: numColumns
    });
  },

  addColumnWidth(sheet, column, width) {
    this.operations.push({
      type: 'columnWidth',
      sheet: sheet,
      column: column,
      width: width
    });
  },
  addRowHeight(sheet, row, height) {
    this.operations.push({
      type: 'rowHeight',
      sheet: sheet,
      row: row,
      height: height
    });
  },

  addRowHeights(sheet, startRow, numRows, height) {
    this.operations.push({
      type: 'rowHeights',
      sheet: sheet,
      startRow: startRow,
      numRows: numRows,
      height: height
    });
  },
  addColumnWidths(sheet, startColumn, numColumns, width) {
    this.operations.push({
      type: 'columnWidths',
      sheet: sheet,
      startColumn: startColumn,
      numColumns: numColumns,
      width: width
    });
  },
  addDeleteColumns(sheet, column, count) {
    if (!sheet || column == null || count == null || count <= 0) {
      console.log('Invalid deleteColumns operation:', { sheet: !!sheet, column, count });
      return;
    }
    this.operations.push({
      type: 'deleteColumns',
      sheet: sheet,
      column: column,
      count: count
    });
  },

  addDeleteRows(sheet, row, count) {
    if (!sheet || row == null || count == null || count <= 0) {
      console.log('Invalid deleteRows operation:', { sheet: !!sheet, row, count });
      return;
    }
    this.operations.push({
      type: 'deleteRows',
      sheet: sheet,
      row: row,
      count: count
    });
  },

  addCopySheet(sourceSheet, targetSpreadsheet) {
    this.operations.push({
      type: 'copySheet',
      sourceSheet: sourceSheet,
      targetSpreadsheet: targetSpreadsheet
    });
  },

  addAutoResizeColumn(sheet, column) {
    this.operations.push({
      type: 'autoResizeColumn',
      sheet: sheet,
      column: column
    });
  },

  // Add a method to auto-resize and cap width in one operation
  addAutoResizeColumnWithCap(sheet, column, maxWidth = 450) {
    this.operations.push({
      type: 'autoResizeColumnWithCap',
      sheet: sheet,
      column: column,
      maxWidth: maxWidth
    });
  },
  addBorder(range, borderConfig) {
    this.operations.push({
      type: 'border',
      range: range,
      borderConfig: borderConfig
    });
  },



  // Execute all queued operations
  execute() {
    // Group operations by type
    const deleteColumnOps = [];
    const deleteRowOps = [];
    const otherOps = [];

    // Separate deletion operations from others
    this.operations.forEach(op => {
      if (op.type === 'deleteColumns') {
        deleteColumnOps.push(op);
      } else if (op.type === 'deleteRows') {
        deleteRowOps.push(op);
      } else {
        otherOps.push(op);
      }
    });

    // Execute non-deletion operations first
    otherOps.forEach(op => {
      switch (op.type) {
        case 'setValue':
          op.range.setValue(op.value);
          break;

        case 'setValues':
          op.range.setValues(op.values);
          break;

        case 'format':
          if (op.formats.copyFrom) {
            op.formats.copyFrom.copyTo(op.range, { formatOnly: true });
            delete op.formats.copyFrom;
          }
          // Apply all other formatting options
          if (op.formats.background) {
            op.range.setBackground(op.formats.background);
          }
          if (op.formats.fontColor) {
            op.range.setFontColor(op.formats.fontColor);
          }
          if (op.formats.fontWeight) {
            op.range.setFontWeight(op.formats.fontWeight);
          }
          if (op.formats.fontSize) {
            op.range.setFontSize(op.formats.fontSize);
          }
          if (op.formats.fontStyle) {
            op.range.setFontStyle(op.formats.fontStyle);
          }
          if (op.formats.horizontalAlignment) {
            op.range.setHorizontalAlignment(op.formats.horizontalAlignment);
          }
          if (op.formats.verticalAlignment) {
            op.range.setVerticalAlignment(op.formats.verticalAlignment);
          }
          if (op.formats.numberFormat) {
            op.range.setNumberFormat(op.formats.numberFormat);
          }
          if (op.formats.wrapStrategy) {
            op.range.setWrapStrategy(op.formats.wrapStrategy);
          }
          if (op.formats.border) {
            const b = op.formats.border;
            op.range.setBorder(
              b.top || null,
              b.left || null,
              b.bottom || null,
              b.right || null,
              b.vertical || null,
              b.horizontal || null,
              b.color || null,
              b.style || null
            );
          }
          if (op.formats.merge) {
            op.range.merge();
          }
          break;

        case 'note':
          op.range.setNote(op.note);
          break;

        case 'protection':
          const protection = op.range.protect();
          protection.setDescription(op.description);
          if (op.warningOnly) protection.setWarningOnly(true);
          break;

        case 'hideColumns':
          op.sheet.hideColumns(op.startColumn, op.numColumns);
          break;

        case 'columnWidth':
          op.sheet.setColumnWidth(op.column, op.width);
          break;

        case 'columnWidths':
          op.sheet.setColumnWidths(op.startColumn, op.numColumns, op.width);
          break;

        case 'rowHeight':
          op.sheet.setRowHeight(op.row, op.height);
          break;

        case 'rowHeights':
          op.sheet.setRowHeights(op.startRow, op.numRows, op.height);
          break;

        case 'copySheet':
          op.sourceSheet.copyTo(op.targetSpreadsheet);
          break;

        case 'autoResizeColumn':
          op.sheet.autoResizeColumn(op.column);
          break;

        case 'border':
          op.range.setBorder(
            op.borderConfig.left || null,
            op.borderConfig.top || null,
            op.borderConfig.bottom || null,
            op.borderConfig.right || null,
            op.borderConfig.vertical || null,
            op.borderConfig.horizontal || null,
            op.borderConfig.color || "black",
            op.borderConfig.style || SpreadsheetApp.BorderStyle.SOLID
          );
          break;

        case 'autoResizeColumnWithCap':
          op.sheet.autoResizeColumn(op.column);
          const currentWidth = op.sheet.getColumnWidth(op.column);
          if (currentWidth > op.maxWidth) {
            op.sheet.setColumnWidth(op.column, op.maxWidth);
          }
          break;
      }
    });

    // Sort column deletions in descending order (right to left) to avoid index shifting
    deleteColumnOps.sort((a, b) => {
      const aRightmost = a.column + a.count - 1;
      const bRightmost = b.column + b.count - 1;
      return bRightmost - aRightmost;
    });

    // Execute column deletions (ONLY ONCE!)
    deleteColumnOps.forEach(op => {
      if (op.sheet && op.column > 0 && op.count > 0) {
        try {
          op.sheet.deleteColumns(op.column, op.count);
        } catch (e) {
          console.error('Error deleting columns:', e, op);
        }
      }
    });

    // Sort row deletions in descending order (bottom to top) to avoid index shifting
    deleteRowOps.sort((a, b) => b.row - a.row);

    // Execute row deletions
    deleteRowOps.forEach(op => {
      if (op.sheet && op.row > 0 && op.count > 0) {
        try {
          op.sheet.deleteRows(op.row, op.count);
        } catch (e) {
          console.error('Error deleting rows:', e, op);
        }
      }
    });

    // Clear the operations array after execution
    this.operations = [];
  },
  // Clear operations without executing
  clear() {
    this.operations = [];
  },

  // Get count of pending operations (useful for debugging)
  getPendingCount() {
    return this.operations.length;
  }
};

const PLAYER_CLASSES = ["Druid", "Hunter", "Mage", "Priest", "Paladin", "Rogue", "Shaman", "Warlock", "Warrior"];