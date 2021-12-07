const crewSheetName = 'crew';
const thisGs = SpreadsheetApp.getActiveSpreadsheet();
const crewHeaderRow = 4;

function onOpen() {
  thisGs.getSheetByName('Stats').hideSheet;
  thisGs.getSheetByName('CABRatings').hideSheet;
  thisGs.getSheetByName('CollectionBonus').hideSheet;
  thisGs.getSheetByName('Data').hideSheet;
  thisGs.getSheetByName('ImportData').hideSheet;
  thisGs.getSheetByName('IDU').hideSheet;
  thisGs.getSheetByName('Gauntlet Import').hideSheet;

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Upgrade from old Checklist')
    // .addItem('Import/overwrite Crew, Notes, Missions & Settings', 'importAllFromOldSpreadsheet')
    .addItem('Import/overwrite Crew (manual Fused/ Level / Active cols)', 'copyCrewInfoManualFromOldSpreadsheet')
    .addItem('Import/overwrite Crew (Import tab)', 'copyImportTabFromOldSpreadsheet')
    .addItem('Import/overwrite Crew Notes/Keep/Cite cols', 'copyCrewNotesFromOldSpreadsheet')
    .addItem('Import/overwrite Missions', 'copyMissionsFromOldSpreadsheet')
    .addItem('Import/overwrite Settings', 'copySettingsFromOldSpreadsheet')
    .addItem('Copy over user tabs', 'copyUserTabsFromOldSpreadsheet')
    .addToUi();
}

function importAllFromOldSpreadsheet() {
  copyCrewNotesFromOldSpreadsheet();
  copyMissionsFromOldSpreadsheet();
  copySettingsFromOldSpreadsheet();
  copyImportTabFromOldSpreadsheet();
  copyUserTabsFromOldSpreadsheet();
}

function copyCrewInfoManualFromOldSpreadsheet() {
  var sheet = thisGs.getSheetByName(crewSheetName);
  var startingRow = crewHeaderRow + 1;
  var rangeOfNamesToLookup = sheet.getRange("A" + startingRow + ":A" + (sheet.getLastRow()));
  var valuesOfNamesToLookup = rangeOfNamesToLookup.getValues();
  var oldSs = getOldSsUrl();
  var oldSheet = oldSs.getSheetByName(crewSheetName);
  var rangeOfNamesToReadFrom = oldSheet.getRange(5, 1, oldSheet.getLastRow(), 1);
  var valuesOfOldNames = rangeOfNamesToReadFrom.getValues();
  var oldHeaders = oldSheet.getRange(crewHeaderRow, 1, 1, oldSheet.getLastColumn()).getValues();
  var newHeaders = sheet.getRange(crewHeaderRow, 1, 1, sheet.getLastColumn()).getValues();

  oldColLetter = "B";
  newColLetter = "B";
  readMatchValuesAndWrite(sheet, startingRow, oldSheet, oldColLetter, newColLetter, valuesOfOldNames, valuesOfNamesToLookup);

  // Level
  var oldColLetter = findColOfHeader(oldHeaders, "Level");
  var newColLetter = findColOfHeader(newHeaders, "Level");
  if (oldColLetter != 0 && newColLetter != 0) {
    readMatchValuesAndWrite(sheet, startingRow, oldSheet, oldColLetter, newColLetter, valuesOfOldNames, valuesOfNamesToLookup);
  }

  // Active
  var oldColLetter = findColOfHeader(oldHeaders, "Active");
  var newColLetter = findColOfHeader(newHeaders, "Active");
  if (oldColLetter != 0 && newColLetter != 0) {
    readMatchValuesAndWrite(sheet, startingRow, oldSheet, oldColLetter, newColLetter, valuesOfOldNames, valuesOfNamesToLookup);
  }

  copyCrewNotesFromOldSpreadsheet();
}

function copyCrewNotesFromOldSpreadsheet() {
  var sheet = thisGs.getSheetByName(crewSheetName);
  var startingRow = crewHeaderRow + 1;
  var rangeOfNamesToLookup = sheet.getRange("A" + startingRow + ":A" + (sheet.getLastRow()));
  var valuesOfNamesToLookup = rangeOfNamesToLookup.getValues();
  var oldSs = getOldSsUrl();
  var oldSheet = oldSs.getSheetByName(crewSheetName);
  var rangeOfNamesToReadFrom = oldSheet.getRange(5, 1, oldSheet.getLastRow(), 1);
  var valuesOfOldNames = rangeOfNamesToReadFrom.getValues();
  var oldHeaders = oldSheet.getRange(crewHeaderRow, 1, 1, oldSheet.getLastColumn()).getValues();
  var newHeaders = sheet.getRange(crewHeaderRow, 1, 1, sheet.getLastColumn()).getValues();

  var oldColLetter = findColOfHeader(oldHeaders, "Notes");
  var newColLetter = findColOfHeader(newHeaders, "Notes");
  if (oldColLetter != 0 && newColLetter != 0) {
    Logger.log("copying notes");
    readMatchValuesAndWrite(sheet, startingRow, oldSheet, oldColLetter, newColLetter, valuesOfOldNames, valuesOfNamesToLookup);

    // Keep
    var oldColLetter = findColOfHeader(oldHeaders, "Keep");
    var newColLetter = findColOfHeader(newHeaders, "Keep");
    if (oldColLetter != 0 && newColLetter != 0) {
      readMatchValuesAndWrite(sheet, startingRow, oldSheet, oldColLetter, newColLetter, valuesOfOldNames, valuesOfNamesToLookup);
    }
    // cite, if needed
    var oldColLetter = findColOfHeader(oldHeaders, "Cite");
    var newColLetter = findColOfHeader(newHeaders, "Cite");
    if (oldColLetter != 0 && newColLetter != 0) {
      readMatchValuesAndWrite(sheet, startingRow, oldSheet, oldColLetter, newColLetter, valuesOfOldNames, valuesOfNamesToLookup);
    }

  }
}

function copyImportTabFromOldSpreadsheet() {
  var sheet = thisGs.getSheetByName('Import');
  var oldSs = getOldSsUrl();
  var oldSheet = oldSs.getSheetByName('Import');
  var oldImportData = oldSheet.getDataRange().getValues();
  sheet.clearContents;
  sheet.getRange(1, 1, oldImportData.length, oldImportData[0].length).setValues(oldImportData);
}

function copyUserTabsFromOldSpreadsheet() {
  tabsToTransfer = thisGs.getSheetByName("UpgradeFromOldGs").getRange("B2").getValue();
  if (tabsToTransfer !== "") {
    var copySheets = splitString(tabsToTransfer);
    var sourceSpreadsheet = getOldSsUrl();
    targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();

    for (var key in copySheets) {  // OK in V8
      var value = copySheets[key].trim();
      sourceSpreadsheet.getSheetByName(value).copyTo(targetSpreadSheet).setName(value);
    }
  }
}

function copyMissionsFromOldSpreadsheet() {
  var sheetName = "Missions";
  var rangeToCopy = "D10:D129";
  copyRangeFromOldSs(sheetName, rangeToCopy);
  var rangeToCopy = "H10:H15";
  copyRangeFromOldSs(sheetName, rangeToCopy);
}

function copySettingsFromOldSpreadsheet() {
  var sheetName = "Settings";
  var sourceSpreadsheet = getOldSsUrl();
  var sourceSheetName = sheetName;
  var targetSpreadsheet = thisGs;
  var targetSheetName = sheetName;
  var startingRow = 7;
  var evalColLetter = "E";  // i.e., use this column to determine the number of rows to copy. Typically, this is the column that has complete info in every row.
  var sourceRangeStartLetter = "E";
  var sourceRangeEndLetter = "E";
  var targetRangeStartLetter = "E";
  var isReplaceData = true;  // true = *replace* data on target and clear the target range when copying. false = *append* data
  var formulasOnly = "mix";
  // true = copy *only* formulas
  // false = convert formulas to static values
  // "mix" = (don't forget the quotes" combine static *and* formulas
  var manualLastRow = false; // false = let the script figure out the last row. Otherwise, enter a number to manually set the last row to copy.

  copyDataWithoutFormatting({
    sourceSpreadsheet: sourceSpreadsheet,
    sourceSheetName: sourceSheetName,
    targetSpreadsheet: targetSpreadsheet,
    targetSheetName: targetSheetName,
    startingRow: startingRow,
    evalColLetter: evalColLetter,
    sourceRangeStartLetter: sourceRangeStartLetter,
    sourceRangeEndLetter: sourceRangeEndLetter,
    targetRangeStartLetter: targetRangeStartLetter,
    isReplaceData: isReplaceData,
    formulasOnly: formulasOnly,
    manualLastRow: manualLastRow
  });

  var evalColLetter = "F";  // i.e., use this column to determine the number of rows to copy. Typically, this is the column that has complete info in every row.
  var sourceRangeStartLetter = "F";
  var sourceRangeEndLetter = "G";
  var targetRangeStartLetter = "F";
  var manualLastRow = 55;

  copyDataWithoutFormatting({
    sourceSpreadsheet: sourceSpreadsheet,
    sourceSheetName: sourceSheetName,
    targetSpreadsheet: targetSpreadsheet,
    targetSheetName: targetSheetName,
    startingRow: startingRow,
    evalColLetter: evalColLetter,
    sourceRangeStartLetter: sourceRangeStartLetter,
    sourceRangeEndLetter: sourceRangeEndLetter,
    targetRangeStartLetter: targetRangeStartLetter,
    isReplaceData: isReplaceData,
    formulasOnly: formulasOnly,
    manualLastRow: manualLastRow
  });

  var evalColLetter = "H";  // i.e., use this column to determine the number of rows to copy. Typically, this is the column that has complete info in every row.
  var sourceRangeStartLetter = "H";
  var sourceRangeEndLetter = "H";
  var targetRangeStartLetter = "H";
  var manualLastRow = false;

  copyDataWithoutFormatting({
    sourceSpreadsheet: sourceSpreadsheet,
    sourceSheetName: sourceSheetName,
    targetSpreadsheet: targetSpreadsheet,
    targetSheetName: targetSheetName,
    startingRow: startingRow,
    evalColLetter: evalColLetter,
    sourceRangeStartLetter: sourceRangeStartLetter,
    sourceRangeEndLetter: sourceRangeEndLetter,
    targetRangeStartLetter: targetRangeStartLetter,
    isReplaceData: isReplaceData,
    formulasOnly: formulasOnly,
    manualLastRow: manualLastRow
  });
}

// all the functions below are helper functions

function copyRangeFromOldSs(sheetName, range) {
  var targetSheet = thisGs.getSheetByName(sheetName);
  var oldSs = getOldSsUrl();
  var oldSheet = oldSs.getSheetByName(sheetName);
  var oldData = oldSheet.getRange(range).getValues();
  targetSheet.getRange(range).setValues(oldData);
}

function readMatchValuesAndWrite(sheet, startingRow, oldSheet, oldColLetter, newColLetter, valuesOfOldNames, valuesOfNamesToLookup) {
  var oldColNumber = letterToColumn(oldColLetter)
  var rangeToWrite = sheet.getRange(newColLetter + startingRow + ":" + newColLetter + (sheet.getLastRow()));
  var valuesToOverwrite = rangeToWrite.getValues();
  var rangeOfValuesToReadFrom = oldSheet.getRange(startingRow, oldColNumber, oldSheet.getLastRow(), 1);
  var valuesFromOldSheet = rangeOfValuesToReadFrom.getValues();
  var lookupValues = valuesOfOldNames.map((row, i) => row.concat(valuesFromOldSheet[i]))

  for (var i = 0; i < valuesOfNamesToLookup.length; i++) {
    for (var j = 0; j < lookupValues.length; j++) {
      if (lookupValues[j][0] === valuesOfNamesToLookup[i][0]) {
        valuesToOverwrite[i][0] = lookupValues[j][1];
      }
    }
  }
  rangeToWrite.setValues(valuesToOverwrite);
  rangeToWrite.setWrap(true);
}

function getOldSsUrl() {
  oldSsUrl = thisGs.getSheetByName("UpgradeFromOldGs").getRange("B1").getValue();
  var oldSs = SpreadsheetApp.openByUrl(oldSsUrl);
  return oldSs;
}

function splitString(string) {
  var array1 = [{}];
  array1 = string.split(",");
  return array1
}

function letterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function copyDataWithoutFormatting(parameters) {
  var sourceSpreadsheet = parameters.sourceSpreadsheet;
  var sourceSheetName = parameters.sourceSheetName;
  var targetSpreadsheet = parameters.targetSpreadsheet;
  var targetSheetName = parameters.targetSheetName;
  var startingRow = parameters.startingRow;
  var evalColLetter = parameters.evalColLetter;
  var sourceRangeStartLetter = parameters.sourceRangeStartLetter;
  var sourceRangeEndLetter = parameters.sourceRangeEndLetter;
  var targetRangeStartLetter = parameters.targetRangeStartLetter;
  var isReplaceData = parameters.isReplaceData;
  var formulasOnly = parameters.formulasOnly;
  var manualLastRow = parameters.manualLastRow;

  var ss = sourceSpreadsheet.getSheetByName(sourceSheetName);
  var ts = targetSpreadsheet.getSheetByName(targetSheetName);
  var sourceRangeStartColNumber = ss.getRange(sourceRangeStartLetter + "1").getColumn();
  var sourceRangeEndColNumber = ss.getRange(sourceRangeEndLetter + "1").getColumn();
  var sourceNumberOfCols = sourceRangeEndColNumber - sourceRangeStartColNumber + 1;
  var targetRangeStartColNumber = ss.getRange(targetRangeStartLetter + "1").getColumn();

  var sourceLastRow = manualLastRow != false ? manualLastRow : getLastRowOfCol(ss, evalColLetter);

  var sourceToGet = sourceRangeStartLetter + startingRow + ":" + sourceRangeEndLetter + sourceLastRow;
  var sourceNumberOfRows = sourceLastRow - startingRow + 1;

  if (formulasOnly == "mix") {
    var dataToCopy = getValuesAndFormulas(ss, sourceToGet);
  }
  else if (formulasOnly == true) {
    var dataToCopy = ss.getRange(sourceToGet).getFormulas();
  }
  else {
    var dataToCopy = ss.getRange(sourceToGet).getValues();
  }

  var targetFirstEmptyRow = ts.getLastRow() + 1;

  if (isReplaceData) {
    var targetFirstEmptyRow = startingRow;
    ts.getRange(startingRow, targetRangeStartColNumber, targetFirstEmptyRow, sourceNumberOfCols).setValue('');
  }
  ts.getRange(targetFirstEmptyRow, targetRangeStartColNumber, sourceNumberOfRows, sourceNumberOfCols).setValues(dataToCopy);
}

function getValuesAndFormulas(ss, a1Range) {
  var arrOrigData = ss.getRange(a1Range).getValues();
  var arrOrigForm = ss.getRange(a1Range).getFormulas();

  for (lin in arrOrigForm)
    for (col in arrOrigForm[lin])
      if (arrOrigForm[lin][col] != "")
        arrOrigData[lin][col] = arrOrigForm[lin][col];

  return arrOrigData;
}

function getLastRowOfCol(sheet, col) {
  var l = sheet.getLastRow();
  var range = sheet.getRange(col + l);
  if (range.getValue() !== "") {
    var i = l;
  } else {
    var i = range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }
  if (i === 1) { i = 2 }
  return i;
}

function getColLetterFromNumber(colNumber) {
  var sheet = thisGs.getSheetByName(crewSheetName);
  var range = sheet.getRange(1, colNumber, 1, 1);
  return range.getA1Notation().match(/([A-Z]+)/)[0];
}

function findColOfHeader(headerData, string) {
  var colNum = headerData[0].indexOf(string) + 1;
  if (colNum > 0) {
    return getColLetterFromNumber(colNum);
  }
  return 0;
}
