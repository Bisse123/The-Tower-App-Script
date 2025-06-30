const shared = {
  // Prompts if the user wants to continue import if new sheet has a different export version than the old one 
  "askImportDifferentExportVersion": function(oldVersion, newVersion) {
    var ui = SpreadsheetApp.getUi()
    var msg = "The old sheet uses a different (older) export version (" + oldVersion + "). Do you want to continue the import to version " + newVersion + "?"
    var response = ui.alert(
      msg,
      ui.ButtonSet.YES_NO
    )
    if (response == ui.Button.YES) {
      return true
    }
    return false
  },

  // Alerts the user if something went wrong
  uiAlert: function(inputText) {
    var ui = SpreadsheetApp.getUi()
    ui.alert(
      inputText,
      ui.ButtonSet.OK
    )
  },

  compareVersions: function(oldVersion, newVersion) {
    var oldVersionNumber = parseInt(oldVersion.replace(/\D/g, ""), 10)
    var newVersionNumber = parseInt(newVersion.replace(/\D/g, ""), 10)

    if (oldVersionNumber > newVersionNumber) return -1
    if (oldVersionNumber < newVersionNumber) return oldVersionNumber
    return 0
  },

  openSpreadsheet: function(sheet, idType) {
    var sheetID = this.findSheetTypeID(sheet, idType).id
    if (!sheetID) {
      return null
    }
    sheetID = this.extractSheetId(sheetID)
    if (!sheetID) {
      return null
    }
    return SpreadsheetApp.openById(sheetID)
  },

  findSheetTypeID: function(sheet, idType) {
    var idType = idType || "IDS Master's ID"
    var idValues = sheet.getDataRange().getValues()
    var regex = new RegExp(idType, "i")
    for (var i = 0; i < idValues.length; i++) {
      for (var j = 0; j < idValues[i].length; j++) {
        if (regex.test(idValues[i][j]) && idValues[i][j].indexOf("script") === -1) {
          return {"id": idValues[i][j + 2], "cell": sheet.getRange(i + 1, j + 2), "isImported": sheet.getRange(i + 2, j + 4)}
        }
      }
    }
    return {"id": null, "cell": null, "isImported": null}
  },
  extractSheetId: function(input) {
  input = input.trim()
  var idPattern = /^[a-zA-Z0-9-_]{20,}$/
  var urlPattern = /\/spreadsheets\/d\/([a-zA-Z0-9-_]{20,})/

  if (idPattern.test(input)) {
    return input
  }
  var match = input.match(urlPattern)
  if (match && match[1]) {
    return match[1]
  }
  return null
  },
  splitNameAndVersion: function(sheetName) {
    var lastSpace = sheetName.lastIndexOf(" ")
    if (lastSpace === -1) {
      return { base: sheetName, version: "" }
    }
    var base = sheetName.substring(0, lastSpace)
    var version = sheetName.substring(lastSpace + 1)
    return { base: base, version: version }
  },
}

  // Prompts if the user wants to update their lab ID in the ID Master sheet
function askUpdateAndDeleteSheet() {
  var newSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  if (!newSpreadsheet) {
    shared.uiAlert("Spreadsheet not found")
    return
  }

  var newSheetID = newSpreadsheet.getId()
  if (!newSheetID) {
    shared.uiAlert("No Sheet ID")
    return
  }

  var newHPSheet = newSpreadsheet.getSheetByName("Home Page")
  if (!newHPSheet) {
    shared.uiAlert("Home Page Sheet not found")
    return
  }

  var sheetType = newHPSheet.getRange("B2").getValue()
  var sheetTypeFunction = sheetVars(sheetType)
  if (!sheetTypeFunction) {
    shared.uiAlert("Incorrect sheetType: " + sheetType)
    return
  }

  var newIdSheet = newSpreadsheet.getSheetByName("IDS")
  if (!newIdSheet) {
    shared.uiAlert("IDS sheet not found")
    return
  }

  var idMasterSpreadsheet = shared.openSpreadsheet(newIdSheet)
  if (!idMasterSpreadsheet) {
    shared.uiAlert("IDS Master Spreadsheet failed to open, please ensure it's linked correctly in the IDS Sheet")
    return
  }
  
  var idType = sheetType + " ID"
  var idMasterID = idMasterSpreadsheet.getId()
  var res = updateSheet(idType, newSheetID, idMasterID)
  shared.uiAlert(res)
}

function updateSheet(idType, newSheetID, idMasterID) {
  try {
    var newSpreadsheet = SpreadsheetApp.openById(newSheetID);
    var newIdSheet = newSpreadsheet.getSheetByName("IDS");
    var newSheetTypeID = shared.findSheetTypeID(newIdSheet);
    var isImported = newSheetTypeID.isImported;
    if (isImported.getValue() !== "✅") {
      return {
        success: false,
        message: "Can not update until old sheet has been Imported.",
        updated: false
      };
    }

    var idMasterSpreadsheet = SpreadsheetApp.openById(idMasterID);
    var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS");

    var sheetTypeID = shared.findSheetTypeID(idMasterSheet, idType);
    var idCell = sheetTypeID.cell;
    if (!idCell) {
      return {
        success: false,
        message: "Old sheet ID not found: " + idType,
        updated: false
      };
    }
    var oldSpreadSheet = shared.openSpreadsheet(idMasterSheet, idType);
    var oldSheetID = oldSpreadSheet.getId();

    var oldFile = Drive.Files.get(oldSheetID, {fields: "id, name, parents"});
    var newFile = Drive.Files.get(newSheetID, {fields: "id, name, parents"});

    var newNameParts = shared.splitNameAndVersion(newSpreadsheet.getName());
    var oldNameParts = shared.splitNameAndVersion(oldSpreadSheet.getName());
    var baseName = oldNameParts.base;
    var newVersion = newNameParts.version;
    var finalName = baseName + (newVersion ? " " + newVersion : "");
    newSpreadsheet.rename(finalName);
    Drive.Files.update(
      {
        name: finalName
      },
      newSheetID,
      null,
      {
        addParents: oldFile.parents.join(","),
        removeParents: newFile.parents.join(",")
      }
    );

    Drive.Files.update({trashed: true}, oldSheetID);
    idCell.setValue(newSheetID);

    return {
      success: true,
      message: "New ID Set, new sheet moved and renamed, old sheet deleted.",
      updated: true
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString(),
      updated: false
    };
  }
}