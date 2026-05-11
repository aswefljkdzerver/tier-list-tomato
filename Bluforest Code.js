/*
BLUFOREST SCRIPT! NOTHING IN HERE INTERACTS WITH ANYTHING ON THE OTHER TWO FILES!
Please don't edit without telling me.

To-do list:
1. Come up with a better way of discerning between template types. Maybe an invisible number in the corner?
*/


//[WORKS!] Adds or subtracts rows and columns so that the dimensions of the sheet match specifications.
function setSheetDimensions(editSheet, targetHeight, targetWidth) {
  const width = editSheet.getMaxColumns()
  const height = editSheet.getMaxRows()
  //Columns
  if (width > targetWidth) {
    editSheet.deleteColumns(targetWidth+1, width-targetWidth)
  } else if (width < targetWidth) {
    editSheet.insertColumns(width, targetWidth-width) 
  }

  //Rows
  if (height > targetHeight) {
    editSheet.deleteRows(targetHeight+1, height-targetHeight)
  } else if (height < targetHeight) {
    editSheet.insertRows(height, targetHeight-height) 
  }
}

//[WORKS!] Using information from data and search_keys, via base copy paste source, make the protoype tier list
function bluRender() {
  //Loading sheets
  const ss = SpreadsheetApp.getActive()
  const ui = SpreadsheetApp.getUi()
  const templateSheet = ss.getSheetById(894156788) //search_keys
  const copySheet = ss.getSheetById(934565830) //base copy paste source
  const dataSheet = ss.getSheetById(2012833542) //data
  const pasteSheet = ss.getSheetById(1854923443) //prototype
  const constants = ss.getSheetById(452915163) //consts
  const sourceSheet = ss.getSheetById(1339075861) //Source Data

  //UI alert
  var warning = ui.alert("Before we begin: \n\n - You're currently looking to run the FULL update. \n - If the tier list template hasn't been changed since the last update, you can use QUICK update, which is less likely to kill your computer. \n - If the characters themselves haven't changed, you can skip directly to the Export stage. Character details are updated automatically. \n - If the overall layout of the tier list template HAS been changed since the last update, then you're in the right place. \n\n Proceed with the FULL update?", ui.ButtonSet.YES_NO)
  if (warning != ui.Button.YES) {
    return
  }

  //Loading constants
  const templateWidth = templateSheet.getMaxColumns()
  const templateHeight = templateSheet.getMaxRows()

  //Loading nChars
  var nChars
  var sourceDataRarityColumn = [].concat(...sourceSheet.getRange(1, 1, 1, sourceSheet.getMaxColumns()).getValues()).indexOf("Rarity")+1
  var sourceDataRarities = [].concat(...sourceSheet.getRange(2, sourceDataRarityColumn, sourceSheet.getMaxRows()-1, 1).getValues())
  for (i in sourceDataRarities) {
    if (Number(sourceDataRarities[i]) < 5) {
      nChars = Number(i)
      break
    }
  }

  //Makes an array of row heights from search_keys (this and the one below assume that search keys is relatively small)
  //Index 0 in this array is Row 1
  var searchKeyHeights = []
  for (var i = 1; i <= templateHeight; i++) {
    searchKeyHeights.push(templateSheet.getRowHeight(i))
  }

  //Makes an array of column widths from search_keys
  //Index 0 in this array is Column 1
  var searchKeyWidths = []
  for (var i = 1; i <= templateWidth; i++) {
    searchKeyWidths.push(templateSheet.getColumnWidth(i))
  }

  //Reformat "base copy paste source"
  copySheet.clear()
  copySheet.clearConditionalFormatRules()
  setSheetDimensions(copySheet, templateHeight, templateWidth)

  //Set the row/column heights/widths
  for (var i = 0; i < searchKeyHeights.length; i++) {
    copySheet.setRowHeightsForced(i+1, 1, searchKeyHeights[i])
  }
  for (var i = 0; i < searchKeyWidths.length; i++) {
    copySheet.setColumnWidth(i+1, searchKeyWidths[i])
  }

  //Copy template to copysheet
  templateSheet.getRange(1, 1, templateHeight, templateWidth).copyTo(copySheet.getRange(1, 1, templateHeight, templateWidth))

  //Scan "base copy paste source" for search key information and store it
  var replaceKeys = []
  var blockHeight = 0
  var readRange
  var readText
  var relNameLoc
  var nameLoc
  var nameLoc2
  for (var i = 1; i <= copySheet.getMaxRows(); i++) {
    for (var j = 1; j <= copySheet.getMaxColumns(); j++) {

      readRange = copySheet.getRange(i, j)
      readText = readRange.getValue()

      if (readText == "??display_name") {
        if (blockHeight == 0) {
          nameLoc = readRange.getA1Notation()
          relNameLoc = [i, j]
        } else {
          nameLoc2 = readRange.getA1Notation()
        }
      } else if (readText.substring(0, 2) == "??") {
        replaceKeys.push([i, j])
      } else if (readText == "END") {
        blockHeight = i
        readRange.setValue("")
      }
    }
  }

  //Sets the character names in base copy paste source
  copySheet.getRange(nameLoc).setValue('=INDIRECT("data!C"&(ROW()-'+relNameLoc[0]+')/'+blockHeight+'+2)')
  copySheet.getRange(nameLoc2).setValue('=INDIRECT("data!C"&(ROW()-'+relNameLoc[0]+')/'+blockHeight+'+2)')

  //Replace search keys with formulae
  for (loc in replaceKeys) {
    readRange = copySheet.getRange(replaceKeys[loc][0], replaceKeys[loc][1])
    readText = readRange.getValue().substring(2)

    if (replaceKeys[loc][0] <= blockHeight) {
      readRange.setValue('=LOOKER(REGEXEXTRACT(TO_TEXT('+nameLoc+'),"[\\w\\s.()-]+"), "data", "'+readText+'")')
    } else {
      readRange.setValue('=LOOKER(REGEXEXTRACT(TO_TEXT('+nameLoc2+'),"[\\w\\s.()-]+"), "data", "'+readText+'")')
    }
  }

  //Just in case
  blockHeight = Number(blockHeight)

  //Reset "prototype"
  pasteSheet.clear()
  pasteSheet.clearConditionalFormatRules()
  setSheetDimensions(pasteSheet, nChars*blockHeight+1, templateWidth)

  //Resize columns
  for (var i = 1; i <= templateWidth; i++) {
    pasteSheet.setColumnWidth(i, searchKeyWidths[i-1])
  }

  //Find which template each character needs to use
  var charList = [].concat(...dataSheet.getRange("A2:A").getValues())
  var charTemplateList = [].concat(...dataSheet.getRange("B2:B").getValues())
  var useTemplate

  //The main event!
  for (var i = 0; i < nChars; i++) {
    //Find which template to copy-paste
    useTemplate = charTemplateList[i]

    //Copy-pasting
    copySheet.getRange(useTemplate*blockHeight+1, 1, blockHeight, templateWidth).copyTo(pasteSheet.getRange(blockHeight*i+1, 1, blockHeight, templateWidth))

    //Resize rows
    for (var j = 1; j <= blockHeight+1; j++) {
      pasteSheet.setRowHeightsForced(blockHeight*i+j, 1, searchKeyHeights[blockHeight*useTemplate+j-1])
    }
  }
  //Last row
  pasteSheet.getRange("1:1").copyTo(pasteSheet.getRange(nChars*blockHeight+1, 1, 1, templateWidth))
}

//Determine which template a given block is using
//readRow is the first line, the one above the block in the template sheet
function whichTemplate(sheet, readRow) {
  if (sheet.getRange(readRow+15, 3).getFormula()) {
    return 1
  } else {
    return 0
  }
}

//[WORKS!] This is a faster render script that skips some of the more time-consuming steps by assuming that the overall block layout hasn't changed, and nobody has tampered with base copy paste source
function bluRenderQuick() {
  //Loading sheets
  const ss = SpreadsheetApp.getActive()
  const ui = SpreadsheetApp.getUi()
  const templateSheet = ss.getSheetById(894156788) //search_keys
  const copySheet = ss.getSheetById(934565830) //base copy paste source
  const dataSheet = ss.getSheetById(2012833542) //data
  const pasteSheet = ss.getSheetById(1854923443) //prototype
  const sourceSheet = ss.getSheetById(1339075861) //Source Data

  //UI alert
  var warning = ui.alert("Before we begin: \n\n - You're currently looking to run the QUICK update. \n - If the tier list TEMPLATE LAYOUT has been changed since the last update, then you MUST use FULL update instead. \n - If the only changes since the previous update have been character details, then you do NOT need to update, as character details are updated automatically. Skip to the Export stage. \n - If the overall layout of the tier list template HAS NOT been changed, but characters have been added, removed, or have had their rankings changed since the last update, then you're in the right place. \n\n Proceed with the QUICK update?", ui.ButtonSet.YES_NO)
  if (warning != ui.Button.YES) {
    return
  }

  //Loading nChars
  var nChars
  var sourceDataRarityColumn = [].concat(...sourceSheet.getRange(1, 1, 1, sourceSheet.getMaxColumns()).getValues()).indexOf("Rarity")+1
  var sourceDataRarities = [].concat(...sourceSheet.getRange(2, sourceDataRarityColumn, sourceSheet.getMaxRows()-1, 1).getValues())
  for (i in sourceDataRarities) {
    if (Number(sourceDataRarities[i]) < 5) {
      nChars = Number(i)
      break
    }
  }

  //Loading constants
  const templateWidth = templateSheet.getMaxColumns()
  const templateHeight = templateSheet.getMaxRows()

  //Loading blockHeight
  var heightChecking = [].concat(...templateSheet.getRange("A:A").getValues())
  const blockHeight = heightChecking.indexOf("END")+1

  //Loading template stats
  var charList = [].concat(...dataSheet.getRange(2, 1, dataSheet.getMaxRows(), 1).getValues())
  var charTemplateList = [].concat(...dataSheet.getRange(2, 2, dataSheet.getMaxRows(), 1).getValues())
  var useTemplate

  //Makes an array of row heights from search_keys (this and the one below assume that search keys is relatively small)
  //Index 0 in this array is Row 1
  var searchKeyHeights = []
  for (var i = 1; i <= templateHeight; i++) {
    searchKeyHeights.push(templateSheet.getRowHeight(i))
  }

  //Checking if we need to add new characters
  var currChars = (pasteSheet.getMaxRows()-1)/blockHeight
  setSheetDimensions(pasteSheet, nChars*blockHeight+1, pasteSheet.getMaxColumns())

  //Adds new characters (if necessary)
  if (currChars < nChars) {
    for (var i = currChars; i < nChars; i++) {
      //Find which template to copy-paste
      useTemplate = charTemplateList[i]

      //Copy-pasting
      copySheet.getRange(useTemplate*blockHeight+1, 1, blockHeight, templateWidth).copyTo(pasteSheet.getRange(blockHeight*i+1, 1, blockHeight, templateWidth))

      //Resize rows
      for (var j = 1; j <= blockHeight+1; j++) {
        pasteSheet.setRowHeightsForced(blockHeight*i+j, 1, searchKeyHeights[blockHeight*useTemplate+j-1])
      }
    }
    //Last row
    pasteSheet.getRange("1:1").copyTo(pasteSheet.getRange(nChars*blockHeight+1, 1, 1, templateWidth))
  }

  //Checks to see if all of the current templates are correct, and corrects if necessary
  for (var i = 0; i < nChars; i++) {
    //Finding the use template
    useTemplate = charTemplateList[i]

    //If the templates are different
    if (whichTemplate(pasteSheet, blockHeight*i+1) != charTemplateList[i]) {
      //Copy-pasting
      copySheet.getRange(useTemplate*blockHeight+1, 1, blockHeight, templateWidth).copyTo(pasteSheet.getRange(blockHeight*i+1, 1, blockHeight, templateWidth))

      //Resize rows
      for (var j = 1; j <= blockHeight+1; j++) {
        pasteSheet.setRowHeightsForced(blockHeight*i+j, 1, searchKeyHeights[blockHeight*useTemplate+j-1])
      } 
    }
  }
}

function bluExport() {
  //Loading sheets
  const ss = SpreadsheetApp.getActive()
  const ui = SpreadsheetApp.getUi()
  const templateSheet = ss.getSheetById(894156788) //search_keys
  const pasteSheet = ss.getSheetById(1854923443) //prototype
  const exportSheet = ss.getSheetById(391521893) //Arcanist Tier List for export

  //UI alert
  var warning = ui.alert("Hey! Listen! \n\n You've clicked on the \"make export\" button, meaning you have: \n 1. Made changes to the tier list information, \n 2. Pressed quick update or full update if necessary, \n 3. And are ready to push the tier list to the main document. \n\n Proceed?", ui.ButtonSet.YES_NO)
  if (warning != ui.Button.YES) {
    return
  }

  //Loading blockHeight
  var heightChecking = [].concat(...templateSheet.getRange("A:A").getValues())
  const blockHeight = heightChecking.indexOf("END")+1

  //Loading constants
  const maxRows = pasteSheet.getMaxRows()
  const maxCols = pasteSheet.getMaxColumns()
  const nChars = (maxRows-1)/blockHeight

  //Makes an array of row heights from search_keys (this and the one below assume that search keys is relatively small)
  //Index 0 in this array is Row 1
  var templateHeight = templateSheet.getMaxRows()
  var searchKeyHeights = []
  for (var i = 1; i <= templateHeight; i++) {
    searchKeyHeights.push(templateSheet.getRowHeight(i))
  }

  //Resizing and resetting the Export sheet
  exportSheet.clear()
  exportSheet.clearConditionalFormatRules()
  setSheetDimensions(exportSheet, maxRows, maxCols)

  //Copies and pastes the content over
  var readRange = pasteSheet.getRange(1, 1, maxRows, maxCols)
  var pasteRange = exportSheet.getRange(1, 1, maxRows, maxCols).getGridId()
  readRange.copyValuesToRange(pasteRange, 1, maxCols, 1, maxRows)
  readRange.copyFormatToRange(pasteRange, 1, maxCols, 1, maxRows)

  var useTemplate

  //Resize rows and do temporary hacks
  for (var i = 0; i < nChars; i++) {
    //Find which template is being used
    useTemplate = whichTemplate(exportSheet, blockHeight*i+1)

    //Temporary link hack, make this better later
    pasteSheet.getRange(blockHeight*i+18, 10, 2, 1).copyTo(exportSheet.getRange(blockHeight*i+18, 10, 2, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA)

    //Resize rows
    for (var j = 1; j <= blockHeight+1; j++) {
      exportSheet.setRowHeightsForced(blockHeight*i+j, 1, searchKeyHeights[blockHeight*useTemplate+j-1])
    }
  }
}

//[HOLY SHIT IT'S SO FUCKING FAST] Does a batch copy-paste to save time on API calls.
function fasterTeambuildingCleanup() {
  //UI Alert
  const ui = SpreadsheetApp.getUi()
  var alert = ui.alert("Make sure you're running this on a COPY of the draft teambuilding sheet. NOWHERE ELSE!\nIf you're trying to run this anywhere else, press \"Cancel\" immediately. Otherwise, press \"OK\".", ui.ButtonSet.OK_CANCEL)
  if (alert != ui.Button.OK) {
    return
  }
  const ss = SpreadsheetApp.getActive()
  const copySheet = ss.getSheetByName("Draft Teambuilding Tierlist")

  //const sheet = ss.getSheetByName("Copy of Draft Teambuilding Tierlist")
  const sheet = ss.getActiveSheet()

  const pasteSheetId = sheet.getSheetId()

  var colMax = sheet.getMaxColumns()
  var rowMax = sheet.getMaxRows()

  var bg = sheet.getRange(1,1,rowMax,colMax).getBackgrounds()
  sheet.clearConditionalFormatRules()
  sheet.getRange(1,1,rowMax,colMax).setBackgrounds(bg)

  var topRowImages = []
  var columnImages = []

  //Record the columns that need to be value-pasted
  for (var col = 1; col <= colMax; col++) {
    //Top Row
    if (sheet.getRange(2,col).getValue() == "CellImage") {
      topRowImages.push(col) 
    }

    //Everything Else
    if (sheet.getRange(5,col).getValue() == "CellImage") {
      columnImages.push(col)
      col++
    }
  }

  var cpReqList = []

  for (var i in topRowImages) {
    cpReqList.push({
      copyPaste: {
        source: {
          sheetId: pasteSheetId,
          startRowIndex: 1,
          endRowIndex: 2,
          startColumnIndex: topRowImages[i]-1,
          endColumnIndex: topRowImages[i]
        },
        destination: {
          sheetId: pasteSheetId,
          startRowIndex: 1,
          endRowIndex: 2,
          startColumnIndex: topRowImages[i]-1,
          endColumnIndex: topRowImages[i]
        },
        pasteType: "PASTE_VALUES"
      }
    })
  }

  for (var i in columnImages) {
    cpReqList.push({
      copyPaste: {
        source: {
          sheetId: pasteSheetId,
          startRowIndex: 4,
          endRowIndex: 47,
          startColumnIndex: columnImages[i]-1,
          endColumnIndex: columnImages[i]
        },
        destination: {
          sheetId: pasteSheetId,
          startRowIndex: 4,
          endRowIndex: 47,
          startColumnIndex: columnImages[i]-1,
          endColumnIndex: columnImages[i]
        },
        pasteType: "PASTE_VALUES"
      }
    })
  }

  Sheets.Spreadsheets.batchUpdate({requests: cpReqList}, ss.getId())

  ui.alert("Done! Isn't it insane how much faster that is than the old version?")
}