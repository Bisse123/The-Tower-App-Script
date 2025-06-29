const sheetVars = (sheetType) => {
    var sheetTypeFuntions = {
    "Laboratory": lab,
    "Workshop": workshop,
    "Ultimate Weapon": ultimate,
    "Themes & Songs": themes,
    "Bots": bots,
    "Relics": relics,
    "Vault": vault,
    "Cards": cards,
    "Modules": modules,
    "Guardians": guardians,
  }
  return sheetTypeFuntions[sheetType]
}

function startImportData(sheetType, sheetID) {
  Logger.log("Starting import for sheet type: " + sheetType + " with ID: " + sheetID)
  var sheetTypeFunction = sheetVars(sheetType)
  var newSpreadsheet =  SpreadsheetApp.openById(sheetID)
  
  if (newSpreadsheet && sheetType && sheetTypeFunction) {
    sheetTypeFunction.importData(sheetType, newSpreadsheet)
    var newIdSheet = newSpreadsheet.getSheetByName("IDS")
    var isImportedRange = shared.findSheetTypeID(newIdSheet).isImported
    isImportedRange.setValue("✅")
  }
}

function importData() {
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

  var newExportSheet = newSpreadsheet.getSheetByName("EXPORT") || newSpreadsheet.getSheetByName("STATS")
  if (!newExportSheet) {
    shared.uiAlert("Can not find sheet version")
    return
  }
  var newVersion = newExportSheet.getRange("A1").getValue()
  
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
  var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
  var oldSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
  var oldsheetID = oldSpreadsheet.getId()
  var oldExportSheet = oldSpreadsheet.getSheetByName("EXPORT") || oldSpreadsheet.getSheetByName("STATS")
  var oldVersion = oldExportSheet.getRange("A1").getValue()
  
  if (oldsheetID === newSheetID) {
    shared.uiAlert("This is the linked " + sheetType + " spreadsheet")
    return
  }

  var compareResult = shared.compareVersions(oldVersion, newVersion)

  if (compareResult === -1) {
    shared.uiAlert("The version of the old sheet (" + oldVersion + ") is newer than the new sheet (" + newVersion + "). Import aborted.")
    return
  }
  else if (compareResult > 0) {
    if (!sheetTypeFunction.isCompatibleVersion(compareResult)) {
      shared.uiAlert("The old sheet uses a different (older) export version (" + oldVersion + ") which does not support importing to the new sheet " + newVersion + ". Import aborted.")
      return
    }
  }
  startImportData(sheetType, newSheetID)
  shared.uiAlert("Import completed!")
}

function doGet(e) {
  function returnError(msg) {
    var template = HtmlService.createTemplateFromFile('onError')
    template.errorMessage = msg
    return template.evaluate().setWidth(400).setHeight(200)
  }

  var newSheetID = e.parameter.spreadsheetID
  if (!newSheetID || newSheetID === "<Script loading...>" ) {
    return returnError("No Sheet ID provided.") 
  }

  var newSpreadsheet = SpreadsheetApp.openById(newSheetID)
  if (!newSpreadsheet) return returnError("Spreadsheet not found.")

  var newHPSheet = newSpreadsheet.getSheetByName("Home Page")
  if (!newHPSheet) return returnError("Home Page Sheet not found.")

  var sheetType = newHPSheet.getRange("B2").getValue()
  var sheetTypeFunction = sheetVars(sheetType)
  if (!sheetTypeFunction) return returnError("Incorrect sheetType: " + sheetType)

  var newExportSheet = newSpreadsheet.getSheetByName("EXPORT") || newSpreadsheet.getSheetByName("STATS")
  if (!newExportSheet) return returnError("Cannot find sheet version (EXPORT or STATS).")
  var newVersion = newExportSheet.getRange("A1").getValue()

  var newIdSheet = newSpreadsheet.getSheetByName("IDS")
  if (!newIdSheet) return returnError("IDS sheet not found.")
  
  var isImported = shared.findSheetTypeID(newIdSheet).isImported.getValue()

  Logger.log(isImported)
  var idMasterSpreadsheet = shared.openSpreadsheet(newIdSheet)
  if (!idMasterSpreadsheet) return returnError("IDS Master Spreadsheet failed to open, please ensure it's linked correctly in the IDS Sheet")

  var idType = sheetType + " ID"
  var idMasterID = idMasterSpreadsheet.getId()
  var idMasterSheet = idMasterSpreadsheet.getSheetByName("IDS")
  var oldSpreadsheet = shared.openSpreadsheet(idMasterSheet, idType)
  if (!oldSpreadsheet) return returnError("Linked old spreadsheet not found.")
  var oldsheetID = oldSpreadsheet.getId()
  var oldExportSheet = oldSpreadsheet.getSheetByName("EXPORT") || oldSpreadsheet.getSheetByName("STATS")
  if (!oldExportSheet) return returnError("Old EXPORT sheet not found.")
  var oldVersion = oldExportSheet.getRange("A1").getValue()

  if (oldsheetID === newSheetID) {
    return returnError("This is the linked " + sheetType + " spreadsheet.")
  }

  var compareResult = shared.compareVersions(oldVersion, newVersion)

  if (compareResult === -1) {
    return returnError("The version of the old sheet (" + oldVersion + ") is newer than the new sheet (" + newVersion + "). Import aborted.")
  }
  // else if (compareResult > 0) {
  //   if (!sheetTypeFunction.isCompatibleVersion(compareResult)) {
  //     return returnError("The old sheet uses a different (older) export version (" + oldVersion + ") which does not support importing to the new sheet " + newVersion + ". Import aborted.")
  //   }
  // }

  // All checks passed, show main import page
  var template = HtmlService.createTemplateFromFile('WebApp')
  template.newSheetID = newSheetID
  template.oldsheetID = oldsheetID
  template.idMasterID = idMasterID
  template.sheetType = sheetType
  template.newVersion = newVersion
  template.oldVersion = oldVersion
  template.idType = idType
  template.isImported = isImported
  template.compareResult = compareResult

  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Import Data');
}

function getPickerHtml(idType, newSheetID, oldsheetID, idMasterID) {
  var template = HtmlService.createTemplateFromFile('picker');
  template.idType = idType;
  template.newSheetID = newSheetID;
  template.oldsheetID = oldsheetID;
  template.idMasterID = idMasterID;
  template.origin = "https://script.google.com";

  return template.evaluate().getContent();
}

function getOAuthToken() {
  try {
    const token = ScriptApp.getOAuthToken();
    return {
      success: true,
      token: token,
      message: 'Token retrieved successfully'
    };
  } catch (error) {
    console.error('Error getting OAuth token:', error);
    return {
      success: false,
      token: null,
      message: error.toString()
    };
  }
}

function checkSheetAccess(newSheetID, oldsheetID, idMasterID) {
  try {
    const fileIds = [newSheetID, oldsheetID, idMasterID];
    const accessibleFiles = [];
    const inaccessibleFiles = [];
    const accessibleDetails = [];
    
    console.log('Checking access to', fileIds.length, 'predefined sheets...');
    
    fileIds.forEach(fileId => {
      try {
        const file = Drive.Files.get(fileId, {
          fields: 'id,name,mimeType,webViewLink'
        });
        
        console.log('Access confirmed for:', file.name, '(' + fileId + ')');
        accessibleFiles.push(fileId);
        accessibleDetails.push({
          id: file.id,
          name: file.name,
          mimeType: file.mimeType,
          webViewLink: file.webViewLink
        });
        
      } catch (error) {
        console.log('No access to file:', fileId, 'Error:', error.toString());
        let name = null;
        try {
          const ss = SpreadsheetApp.openById(fileId);
          name = ss.getName();
        } catch (e) {
        }
        inaccessibleFiles.push({ id: fileId, name: name });
      }
    });
    
    console.log('Access check complete. Accessible:', accessibleFiles.length, 'Inaccessible:', inaccessibleFiles.length);
    
    return {
      success: true,
      accessibleFiles: accessibleFiles,
      inaccessibleFiles: inaccessibleFiles,
      accessibleDetails: accessibleDetails,
      totalFiles: fileIds.length,
      message: `Access check complete. ${accessibleFiles.length} of ${fileIds.length} sheets are accessible.`
    };
    
  } catch (error) {
    console.error('Error checking sheet access:', error);
    return {
      success: false,
      accessibleFiles: [],
      inaccessibleFiles: [],
      accessibleDetails: [],
      message: error.toString()
    };
  }
}