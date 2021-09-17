function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('Stats').hideSheet;
  ss.getSheetByName('CABRatings').hideSheet;
  ss.getSheetByName('CollectionBonus').hideSheet;
  ss.getSheetByName('Data').hideSheet;
  ss.getSheetByName('ImportData').hideSheet;
  ss.getSheetByName('IDU').hideSheet;
  ss.getSheetByName('Gauntlet Import').hideSheet;

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Upgrade from old Checklist')
    .addItem('Import/overwrite Crew, Notes, Missions & Settings', 'importAllFromOldSpreadsheet')
    .addItem('Copy over user tabs', 'copyUserTabsFromOldSpreadsheet')
    .addToUi();
}

function importAllFromOldSpreadsheet() {
  copyCrewNotesFromOldSpreadsheet();
  copyMissionsFromOldSpreadsheet();
  copySettingsFromOldSpreadsheet();
  copyImportTabFromOldSpreadsheet();
}

function copyCrewNotesFromOldSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Crew');
  var startingRow = 5;
  var rangeOfNamesToLookup = sheet.getRange("A" + startingRow + ":A" + (sheet.getLastRow()));
  var valuesOfNamesToLookup = rangeOfNamesToLookup.getValues();
  var rangeToWrite = sheet.getRange("H" + startingRow + ":H" + (sheet.getLastRow()));
  var valuesToOverwrite = rangeToWrite.getValues();
  var oldSs = getOldSs();

  var oldSheet = oldSs.getSheetByName('Crew');
  var rangeOfNamesToReadFrom = oldSheet.getRange(5, 1, oldSheet.getLastRow(), 1);
  var valuesOfOldNames = rangeOfNamesToReadFrom.getValues();
  var rangeOfNotesToReadFrom = oldSheet.getRange(5, 8, oldSheet.getLastRow(), 1);
  var valuesOfOldNotes = rangeOfNotesToReadFrom.getValues();
  const lookupValues = valuesOfOldNames.map((row, i) => row.concat(valuesOfOldNotes[i]))

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

function copyImportTabFromOldSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Import');
  var oldSs = getOldSs();
  var oldSheet = oldSs.getSheetByName('Import');
  var oldImportData = oldSheet.getDataRange().getValues();
  sheet.clearContents;
  sheet.getRange(1, 1, oldImportData.length, oldImportData[0].length).setValues(oldImportData);
}

function copyUserTabsFromOldSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  tabsToTransfer = ss.getSheetByName("UpgradeFromOldGs").getRange("A2").getValue();
  var copySheets = splitString(tabsToTransfer);
  var sourceSpreadsheet = getOldSs();
  targetSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  for (var key in copySheets) {  // OK in V8
    var value = copySheets[key].trim();
    sourceSpreadsheet.getSheetByName(value).copyTo(targetSpreadSheet).setName(value);
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
  var rangeToCopy = "E7:E15";
  copyRangeFromOldSs(sheetName, rangeToCopy);
  var rangeToCopy = "H7";
  copyRangeFromOldSs(sheetName, rangeToCopy);
  var rangeToCopy = "E19:G24";
  copyRangeFromOldSs(sheetName, rangeToCopy);
  var rangeToCopy = "E29:H36";
  copyRangeFromOldSs(sheetName, rangeToCopy);
  var rangeToCopy = "E39:G55";
  copyRangeFromOldSs(sheetName, rangeToCopy);
  var rangeToCopy = "E59:E102";
  copyRangeFromOldSs(sheetName, rangeToCopy);
  var rangeToCopy = "E107:E125";
  copyRangeFromOldSs(sheetName, rangeToCopy);
  var rangeToCopy = "H107:H125";
  copyRangeFromOldSs(sheetName, rangeToCopy);
}

function copyRangeFromOldSs(sheetName, range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName(sheetName);
  var oldSs = getOldSs();
  var oldSheet = oldSs.getSheetByName(sheetName);
  var oldData = oldSheet.getRange(range).getValues();
  targetSheet.getRange(range).setValues(oldData);
}

function getOldSs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  oldSsUrl = ss.getSheetByName("UpgradeFromOldGs").getRange("A1").getValue();
  var oldSs = SpreadsheetApp.openByUrl(oldSsUrl);
  return oldSs;
}

function splitString(string) {
  var array1 = [{}];
  array1 = string.split(",");
  return array1
}
