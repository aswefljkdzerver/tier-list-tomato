// TODO auto load variables
const constsSheetName = "consts"

var labelSymbol = '??';
var rowHeights = [];
var rowHeightIndex = [];
var rowHeightN = [];
var uniqueHeightsIdxs = 0;

var formulaFront = 'HLOOKUP("';
var formulaMid = '", INDIRECT("data!$A$1:$ZR$200"), MATCH(REGEXEXTRACT(TO_TEXT(';
var formulaEnd = '),"[\\w\\s.-]*"), data!$A$1:$A$200, 0),FALSE)';
var nChars = 0;
var currentPatch = 0;
var nameCell = '';
var nameCol = 0;
var nameRow = 0;
var nameCellModule = '';
var height = 0;
var width = 0;
var dataSheetName = '';
var formatSheetName = '';
var baseSheetName = '';
var tierSheetName = '';
var finalTLName = '';
var linkedCellsRows = [];
var linkedCellsCols = [];
var getNames = null;
var variableRowHeights = []


function loadConstsTL() {
  const ss = SpreadsheetApp.getActive();
  const conSheet = ss.getSheetByName("consts");

  const varNameList = [].concat(...conSheet.getRange(1, 1, conSheet.getMaxRows()).getValues()).filter(Boolean);
  const valueList = [].concat(...conSheet.getRange(1, 2, conSheet.getMaxRows()).getValues()).filter(Boolean);

  nChars = valueList[varNameList.indexOf("nChars")];
  nameCell = valueList[varNameList.indexOf("nameCell")];
  nameCol = valueList[varNameList.indexOf("nameCol")];
  nameRow = valueList[varNameList.indexOf("nameRow")];
  currentPatch = valueList[varNameList.indexOf("currentPatch")];
  nameCellModule = valueList[varNameList.indexOf("nameCellModule")];
  height = valueList[varNameList.indexOf("height")];
  width = valueList[varNameList.indexOf("width")];
  dataSheetName = valueList[varNameList.indexOf("dataSheetName")];
  formatSheetName = valueList[varNameList.indexOf("formatSheetName")];
  baseSheetName = valueList[varNameList.indexOf("baseSheetName")];
  tierSheetName = valueList[varNameList.indexOf("tierSheetName")];
  finalTLName = 'Arcanist Tierlist for export';
  linkedCellsRows = [valueList[varNameList.indexOf("linkedCellsRows")]];
  linkedCellsCols = [valueList[varNameList.indexOf("linkedCellsCols")]];
  getNames=sortNames;
  variableRowHeights = [];

  formulaFront = valueList[varNameList.indexOf("lookupFormulaFront")];
  formulaMid = valueList[varNameList.indexOf("lookupFormulaMid")];
  formulaEnd = valueList[varNameList.indexOf("lookupFormulaEnd")];
}


function loadConstsTeam() {
  nChars = 9;

  nameCell = 'B2';
  nameCol = 2;
  nameRow = 2;
  nameCellModule = 'B18';
  height = 16;
  width = 14;
  dataSheetName = 'team data';
  formatSheetName = 'team search_keys';
  baseSheetName = 'team base copy paste source';
  tierSheetName = 'team prototype';
  finalTLName = 'team for export';
  linkedCellsRows = [];
  linkedCellsCols = [];
  variableRowHeights = [];
  getNames = colANames;

  formulaFront = 'HLOOKUP("';
  formulaMid = '", INDIRECT("' + dataSheetName + '!$A$1:$ZR$200"), MATCH(REGEXEXTRACT(TO_TEXT(';
  formulaEnd = '),"[\\w\\s.-]*"), INDIRECT("' + dataSheetName + '!$A$1:$A$200"), 0),FALSE)';
}


function loadConstsExp() {
  nChars = 21;

  nameCell = 'B3';
  nameCol = 2;
  nameRow = 3;
  nameCellModule = 'B7';
  height = 4;
  width = 10;
  dataSheetName = 'Changelog data';
  formatSheetName = 'Changelog search_keys';
  baseSheetName = 'Changelog base copy paste';
  tierSheetName = 'Changelog prototype';
  finalTLName = 'Changelog for export';
  linkedCellsRows = [];
  linkedCellsCols = [];
  variableRowHeights = [4];
  getNames=colANames;

  formulaFront = 'HLOOKUP("';
  formulaMid = '", INDIRECT("' + dataSheetName + '!$A$1:$ZR$200"), MATCH(REGEXEXTRACT(TO_TEXT(';
  formulaEnd = '),"[\\w\\s.-]*"), INDIRECT("' + dataSheetName + '!$A$1:$A$200"), 0),FALSE)';
}


function sortNames(n) {
  // TODO change manually offsetting for correct columns to auto
  var dataSheet = SpreadsheetApp.getActive().getSheetByName(dataSheetName);
  // const headerRow = [].concat(...dataSheet.getRange(1, 1, 1, dataSheet.getMaxColumns()).getValues());
  // var charDisplayNames = dataSheet.getRange(2, [].concat(...dataSheet.getRange(1, 1, 1, dataSheet.getMaxColumns()).getValues()).indexOf("display_name") + 1, nChars).getValues();
  // var hasModule = dataSheet.getRange(2, [].concat(...dataSheet.getRange(1, 1, 1, dataSheet.getMaxColumns()).getValues()).indexOf("module") + 1, nChars).getValues();
  // var normalScores = dataSheet.getRange(2, [].concat(...dataSheet.getRange(1, 1, 1, dataSheet.getMaxColumns()).getValues()).indexOf("final_score") + 1, nChars).getValues();
  // var euphoScores = dataSheet.getRange(2, [].concat(...dataSheet.getRange(1, 1, 1, dataSheet.getMaxColumns()).getValues()).indexOf("final_score_module") + 1, nChars).getValues();
  // var euphoPatch = dataSheet.getRange(2, [].concat(...dataSheet.getRange(1, 1, 1, dataSheet.getMaxColumns()).getValues()).indexOf("euphoria_patch") + 1, nChars).getValues();
  var sheetValues = dataSheet.getRange(1, 1, nChars+1, dataSheet.getMaxColumns()).getValues()
  const headerRow = sheetValues[0];
  sheetValues.splice(0, 1);
  const headerNames = ["module", "display_name", "final_score", "final_score_module", "euphoria_patch"];
  let indexes = {};
  let searchIndex =  0;
  for (let i = 0; i < headerRow.length; i++) {
    if(headerRow[i] === headerNames[searchIndex]){
      indexes[headerRow[i]] = i
      searchIndex++
    };
    if(searchIndex === headerRow.length){
      break;
    };
  };
  for (let i = searchIndex; i < headerNames.length; i++){
    indexes[headerNames[searchIndex]] = -1
  };
  // const indexes = {"display_name": headerRow.indexOf("display_name"), "module": headerRow.indexOf("module"), "final_score": headerRow.indexOf("final_score"), "final_score_module": headerRow.indexOf("final_score_module"), "euphoria_patch": headerRow.indexOf("euphoria_patch")};
  const charDisplayNames = sheetValues.map((a) => [a[indexes["display_name"]]]);
  const hasModule = sheetValues.map((a) => [a[indexes["module"]]]);
  const normalScores = sheetValues.map((a) => [a[indexes["final_score"]]]);
  const euphoScores = sheetValues.map((a) => [a[indexes["final_score_module"]]]);
  const euphoPatch = sheetValues.map((a) => [a[indexes["euphoria_patch"]]]);
  var mergedScores = [...euphoScores];
  for (var i = 0; i < euphoPatch.length; i++) {
    if (mergedScores[i][0] === "none" || euphoPatch[i][0] > currentPatch) {
      mergedScores[i] = normalScores[i];
    }
    mergedScores[i] = [charDisplayNames[i], mergedScores[i], hasModule[i]];
  }
  mergedScores.sort((a, b) => b[1][0] - a[1][0]);
  return [mergedScores.map((a) => a[0]), mergedScores.map((a) => a[2])];
}


function colANames(n) {
  var dataSheet = SpreadsheetApp.getActive().getSheetByName(dataSheetName);
  var changeVer = dataSheet.getRange('A2:A' + (nChars + 1)).getValues();
  var reserved = dataSheet.getRange('C2:C' + (nChars + 1)).getValues();
  return [changeVer, reserved];
}

function makeBase(startRow, endRow, nameAddr) {
  var labeledSheet = SpreadsheetApp.getActive().getSheetByName(formatSheetName);
  var formattedSheet = SpreadsheetApp.getActive().getSheetByName(baseSheetName);
  var sourceRange = labeledSheet.getRange(startRow + ":" + endRow);
  var targetRange = formattedSheet.getRange(startRow, 1);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  sourceRange.copyTo(targetRange);

  var labels = sourceRange.getValues();

  for (var r = 1; r <= height; r++) {
    rowHeights[r-1] = labeledSheet.getRowHeight(r + startRow - 1);
    for (var c = 1; c <= width; c++) {
      var label = labels[r - 1][c - 1];
      if (r == nameRow && c == nameCol) continue;
      if (typeof(label) == "string" && label.startsWith(labelSymbol)) {
        var destCell = formattedSheet.getRange(r + startRow - 1, c);
        destCell.setFormula(formulaFront + label.substring(2) + formulaMid + nameAddr + formulaEnd);
      }
    }
  }

  var lastVal = rowHeights[0];
  var curVal;
  uniqueHeightsIdxs = 0;
  rowHeightIndex[0] = 0;
  rowHeightN[0] = 1;
  for (var r = 1; r < height; r++) {
    curVal = rowHeights[r];
    if (lastVal == curVal) {
      rowHeightN[uniqueHeightsIdxs] += 1;
    } else {
      uniqueHeightsIdxs++;
      rowHeightIndex[uniqueHeightsIdxs] = r;
      rowHeightN[uniqueHeightsIdxs] = 1
    }
    lastVal = curVal;
  }
  for (var i = 0; i <= uniqueHeightsIdxs; i++) {
    formattedSheet.setRowHeightsForced(rowHeightIndex[i] + startRow, rowHeightN[i], rowHeights[rowHeightIndex[i]]);
  }
}


function make_bases() {
  makeBase(1, height, nameCell);
  makeBase(height + 1, height * 2, nameCellModule);
}


function make_tierlist() {
  var formattedSheet = SpreadsheetApp.getActive().getSheetByName(baseSheetName);
  var tierSheet = SpreadsheetApp.getActive().getSheetByName(tierSheetName);
  var sourceRange = formattedSheet.getRange(1 + ":" + height);
  var sourceRangeModule = formattedSheet.getRange((height + 1) + ":" + (height * 2));

  var [charDisplayNames, hasModule] = getNames(nChars);

  for (var i = 0; i < nChars; i++) {
    var source;
    if (hasModule[i][0]) {
      source = sourceRangeModule;
    } else {
      source = sourceRange;
    }
    var targetRange = tierSheet.getRange(i * height + 1, 1);
    source.copyTo(targetRange);
    source.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    tierSheet.getRange(nameRow + (i * height), nameCol).setValue(charDisplayNames[i][0]);
  
    // Row Resize, kinda slow
    for (var j = 0; j <= uniqueHeightsIdxs; j++) {
      tierSheet.setRowHeightsForced(rowHeightIndex[j] + (i * height) + 1, rowHeightN[j], rowHeights[rowHeightIndex[j]]);
    }
    for (var j = 0; j < variableRowHeights.length; j++) {
      tierSheet.setRowHeight((variableRowHeights[j] + (i * height)), 100);
    }
  }
  source.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);

  //Untested BUT this should trim the excess columns
  //tierSheet.deleteColumns(width+1, width+6)
}


function makeExport(){
  var tierSheet = SpreadsheetApp.getActive().getSheetByName(tierSheetName);
  var exportSheet = SpreadsheetApp.getActive().getSheetByName(finalTLName);

  allSourceSheetRange = tierSheet.getRange("1" + ":" + height*nChars);
  allExportSheetRange = exportSheet.getRange("1" + ":" + height*nChars);
  allSourceSheetRange.copyTo(allExportSheetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  allSourceSheetRange.copyTo(allExportSheetRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  allSourceSheetRange.copyTo(allExportSheetRange, SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false);
  allSourceSheetRange.copyTo(allExportSheetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  for (var i = 0; i < linkedCellsCols.length; i++) {
    source = tierSheet.getRange(linkedCellsRows[i] , linkedCellsCols[i]);
    for (var c = 0; c < nChars; c++) {
      target = exportSheet.getRange(linkedCellsRows[i] + c*height, linkedCellsCols[i]);
      source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    }
  }
}

function teambuildingCleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const original = ss.getSheetByName("Draft Teambuilding Tierlist");

  var colMax = sheet.getMaxColumns();
  var rowMax = sheet.getMaxRows();

  //Format
  sheet.clearConditionalFormatRules();
  sheet.getRange(1,1,rowMax,colMax).setBackgrounds(original.getRange(1,1,rowMax,colMax).getBackgrounds());

  for (var col = 1; col <= colMax; col++) {
    //Top Row
    if (sheet.getRange(2,col).getValue() == "CellImage") {
      sheet.getRange(2,col).copyTo(sheet.getRange(2,col), {contentsOnly: true});
    }

    //Everything Else
    if (sheet.getRange(5,col).getValue() == "CellImage") {
      sheet.getRange(5,col,colMax-6).copyTo(sheet.getRange(5,col,colMax-6), {contentsOnly: true});
    }
    //Eternity Exception
    else if (sheet.getRange(7,col).getValue() == "CellImage") {
      sheet.getRange(7,col,colMax-8).copyTo(sheet.getRange(7,col,colMax-8), {contentsOnly: true});
    }
  }
}


function onEdit(e) {
  try {
    if (!e || !e.range || (e.range.getSheet().getName() != "Draft Teambuilding Tierlist")){
      return;
    }
    const editRow = e.range.getRow();
    const editCol = e.range.getColumn();
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getActiveSheet();
    const lookupSheet = ss.getSheetByName("Text Dump");
    const datalookup = lookupSheet.getRange("A2:A").getValues().map(x => x[0]);
    var charName = sheet.getRange(editRow, editCol).getValue();
    var unreleased = charName.indexOf("*");

    if (unreleased != -1) {
      charName = charName.substring(0, unreleased);
    }
    
    if (sheet.getRange(editRow-1, editCol).getValue() != "CellImage" || sheet.getRange(editRow-1, editCol+1).getValue() != "") {
      return; 
    }
    var lookupRow = datalookup.indexOf(charName);
    if (lookupRow === -1 || charName == "") {
      return;
    }

    sheet.getRange(editRow-1, editCol+1).setValue(lookupSheet.getRange(lookupRow+2, 9).getValue());
    sheet.getRange(editRow-1, editCol+1).setTextStyle(lookupSheet.getRange(lookupRow+2, 9).getTextStyle());
  
  } catch (error) {
    throw(error);
  }
}


function boldEssential() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exportSheet2 = ss.getSheetByName(finalTLName);
  const dataSheet2 = ss.getSheetByName(dataSheetName);

  const nameList = [].concat(...dataSheet2.getRange(2, 1, dataSheet2.getMaxRows()-1).getValues());
  const keyList0 = [].concat(...dataSheet2.getRange(2, [].concat(...dataSheet2.getRange(1, 1, 1, dataSheet2.getMaxColumns()).getValues()).indexOf("key_1") + 1, dataSheet2.getMaxRows()-1).getValues());
  const keyList1 = [].concat(...dataSheet2.getRange(2, [].concat(...dataSheet2.getRange(1, 1, 1, dataSheet2.getMaxColumns()).getValues()).indexOf("key_2") + 1, dataSheet2.getMaxRows()-1).getValues());
  const keyList2 = [].concat(...dataSheet2.getRange(2, [].concat(...dataSheet2.getRange(1, 1, 1, dataSheet2.getMaxColumns()).getValues()).indexOf("key_3") + 1, dataSheet2.getMaxRows()-1).getValues());
  const keyList3 = [].concat(...dataSheet2.getRange(2, [].concat(...dataSheet2.getRange(1, 1, 1, dataSheet2.getMaxColumns()).getValues()).indexOf("key_4") + 1, dataSheet2.getMaxRows()-1).getValues());
  const lists = [keyList0, keyList1, keyList2, keyList3];

  var blockHeight = height;

  for (var i = 0; i < height*nChars; i += blockHeight) {
    var name = exportSheet2.getRange(i+nameRow, nameCol).getValue();
    var idx = nameList.indexOf(name);
    
    for(var j =0; j < 4; j++) {

      var key = lists[j][idx];
      if (key == "Flexible") {
        continue;
      }
      // manually offsetting for where name is
      exportSheet2.getRange(i+14, 2*j + 2).setFontWeight("bold");
      exportSheet2.getRange(i+14, 2*j + 2).setFontSize(11);
      exportSheet2.getRange(i+14, 2*j + 2).setVerticalAlignment("middle");
      exportSheet2.getRange(i+14, 2*j + 3).setFontWeight("bold");
      exportSheet2.getRange(i+14, 2*j + 3).setFontSize(11);
      exportSheet2.getRange(i+14, 2*j + 3).setVerticalAlignment("middle");
      exportSheet2.getRange(i+15, 2*j + 2).setFontWeight("bold");
      exportSheet2.getRange(i+15, 2*j + 2).setFontSize(11);
      exportSheet2.getRange(i+15, 2*j + 2).setVerticalAlignment("middle");
    }
  }
}


function boldScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exportSheet2 = ss.getSheetByName(finalTLName);
  const dataSheet2 = ss.getSheetByName(dataSheetName);

  const nameList = [].concat(...dataSheet2.getRange(2, 3, dataSheet2.getMaxRows()-1).getValues());
  const patchList = [].concat(...dataSheet2.getRange(2, [].concat(...dataSheet2.getRange(1, 1, 1, dataSheet2.getMaxColumns()).getValues()).indexOf("euphoria_patch") + 1, dataSheet2.getMaxRows()-1).getValues());

  var blockHeight = height;

  for (var i = 0; i < height*nChars; i += blockHeight) {
    var name = exportSheet2.getRange(i+nameRow, nameCol).getValue();
    var idx = nameList.indexOf(name);
    var patch = patchList[idx];
    if (patch == "none") {
      continue;
    }
    // manually offsetting for eupho vs noneupho final score
    if (patch <= currentPatch) {
      exportSheet2.getRange(i+15, 4).setFontWeight("normal");
      exportSheet2.getRange(i+15, 4).setFontSize(11);
      exportSheet2.getRange(i+15, 4).setVerticalAlignment("middle");
    } else {
      exportSheet2.getRange(i+15, 5).setFontWeight("normal");
      exportSheet2.getRange(i+15, 5).setFontSize(11);
      exportSheet2.getRange(i+15, 5).setVerticalAlignment("middle");
    }
  }
}


function renderTL() {
  loadConstsTL();
  make_bases();
  make_tierlist();
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  SpreadsheetApp.flush();
  makeExport();
  boldScores();
  // Row Resize, kinda slow, refactor this prettier later
  var tierSheet = SpreadsheetApp.getActive().getSheetByName(finalTLName);
  for (var i = 0; i < nChars; i++) {
    for (var j = 0; j <= uniqueHeightsIdxs; j++) {
      tierSheet.setRowHeightsForced(rowHeightIndex[j] + (i * height) + 1, rowHeightN[j], rowHeights[rowHeightIndex[j]]);
    }
    for (var j = 0; j < variableRowHeights.length; j++) {
      tierSheet.setRowHeight((variableRowHeights[j] + (i * height)), 100);
    }
  }
}


function makeExportTL() {
  loadConstsTL();
  makeExport();
  boldScores();
}


function renderExplanation() {
  loadConstsExp();
  make_bases();
  make_tierlist();
  makeExport();
}


function makeExportExp() {
  loadConstsExp();
  makeExport();
}


function renderTeam() {
  loadConstsTeam();
  make_bases();
  make_tierlist();
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  SpreadsheetApp.flush();
  makeExport();
  boldEssential();
  
  // Row Resize, kinda slow, refactor this prettier later
  var tierSheet = SpreadsheetApp.getActive().getSheetByName(finalTLName);
  for (var i = 0; i < nChars; i++) {
    for (var j = 0; j <= uniqueHeightsIdxs; j++) {
      tierSheet.setRowHeightsForced(rowHeightIndex[j] + (i * height) + 1, rowHeightN[j], rowHeights[rowHeightIndex[j]]);
    }
    for (var j = 0; j < variableRowHeights.length; j++) {
      tierSheet.setRowHeight((variableRowHeights[j] + (i * height)), 100);
    }
  }
}


function makeExportTeam() {
  loadConstsTeam();
  makeExport();
  boldEssential();
}

