const sheetVars = (sheetType) => {
  var sheetTypeFunctions = {
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
  };
  return sheetTypeFunctions[sheetType];
};

const spreadsheets = (() => {
  const storedSpreadsheets = {
    newSpreadsheet: "",
    oldSpreadsheet: "",
    idMasterSpreadsheet: "",
  };
  return function (spreadsheet, sheetID) {
    if (!spreadsheet) {
      console.log("No spreadsheet name provided.");
      return null;
    }
    if (storedSpreadsheets[spreadsheet]) {
      return storedSpreadsheets[spreadsheet];
    }
    if (!sheetID) {
      console.log("Spreadsheet not defined and no sheet ID provided.");
      return null;
    }
    var spreadsheetInfo = SheetsAPI.getSpreadsheet(sheetID);
    if (!spreadsheetInfo) {
      console.log(`Spreadsheet not found with ID: ${sheetID}`);
      return null;
    }
    storedSpreadsheets[spreadsheet] = spreadsheetInfo;
    return spreadsheetInfo;
  }
})();


function doGet(e) {
  console.log("doGet called with parameters: " + JSON.stringify(e.parameter));
  var template = HtmlService.createTemplateFromFile("WebApp");
  template.newSheetID = e.parameter.newSheetID;
  template.oldSheetID = e.parameter.oldSheetID;
  template.idMasterID = e.parameter.idMasterID;
  template.sheetType = e.parameter.sheetType;
  template.API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');
  template.APP_ID = PropertiesService.getScriptProperties().getProperty('APP_ID');

  return template
    .evaluate()
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setTitle("Import Data");
}

function onOpen(e) {
  console.log("onOpen called with event: " + JSON.stringify(e));
  try {
    var sheetType = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Home Page").getRange("B2").getValue();
    if (sheetVars(sheetType)) {
      console.log("Sheet type found in Home Page B2: " + sheetType);
      var ui = SpreadsheetApp.getUi();
      ui.createMenu("Import Data")
        .addItem("Import Data", "showImportDialog")
        .addToUi();
    }
  } catch (error) {
  }
}

function showImportDialog() {
  console.log("showImportDialog called");
  try {
    var sheetType = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Home Page").getRange("B2").getValue();
    var newSheetID = SpreadsheetApp.getActiveSpreadsheet().getId();
    var idMasterInfo = shared.findSheetTypeID(newSheetID, "IDS");
    var idMasterID = idMasterInfo ? shared.extractSheetId(idMasterInfo.id) : "";
    var oldSheetInfo = idMasterID ? shared.findSheetTypeID(idMasterID, "IDS", sheetType + " ID") : "";
    var oldSheetID = oldSheetInfo ? shared.extractSheetId(oldSheetInfo.id) : "";

    var template = HtmlService.createTemplateFromFile("WebApp");
    template.newSheetID = newSheetID;
    template.oldSheetID = oldSheetID;
    template.idMasterID = idMasterID;
    template.sheetType = sheetType;
    template.API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');
    template.APP_ID = PropertiesService.getScriptProperties().getProperty('APP_ID');
  
    template.API_KEY = PropertiesService.getScriptProperties().getProperty('API_KEY');
    template.APP_ID = PropertiesService.getScriptProperties().getProperty('APP_ID');
  

    var html = template
      .evaluate()
      // .setWidth(600)
      // .setHeight(300)
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setTitle("Import Data");
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    console.log("Error in showImportDialog: " + error.message);
    SpreadsheetApp.getUi().alert("Error: " + error.message);
  }
}

function importData(newSheetID, oldSheetID, idMasterID, sheetType, versionDifference) {
  console.log(
    `Starting import for sheet type: ${sheetType} with ID: ${newSheetID}`
  );
  try {
    if (!sheetType) {
      console.log("Sheet type is not defined.");
      return {
        success: false,
        message: "Sheet type is not defined.",
      };
    }
    if (!newSheetID || !oldSheetID || !idMasterID) {
      console.log("One or more required IDs are missing.");
      return {
        success: false,
        message: "One or more required IDs are missing.",
      };
    }
    // Check if new spreadsheet exists
    var newSpreadsheet = spreadsheets("newSpreadsheet", newSheetID);
    if (!newSpreadsheet) {
      return {
        success: false,
        message: `New spreadsheet not found with ID: ${newSheetID}`,
      };
    }

    if (!SheetsAPI.getSheetByName(newSpreadsheet, "IDS")) {
      return {
        success: false,
        message:`IDS sheet not found in the new ${sheetType} spreadsheet.`,
      };
    }

    if (!SheetsAPI.getSheetByName(newSpreadsheet, "EXPORT") && !SheetsAPI.getSheetByName(newSpreadsheet, "STATS")) {
      console.log(`Export sheet not found in new ${sheetType} spreadsheet`);
      return {
        success: false,
        message: `Export sheet not found in new ${sheetType} spreadsheet`,
      };
    }

    var idMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
    if (!idMasterSpreadsheet) {
      return {
        success: false,
        message: `IDS Master Spreadsheet not found with ID: ${idMasterID}`,
      };
    }

    if (!SheetsAPI.getSheetByName(idMasterSpreadsheet, "IDS")) {
      return {
        success: false,
        message: `IDS sheet not found in the IDS Master Spreadsheet.`,
      };
    }

    var oldSpreadsheet = spreadsheets("oldSpreadsheet", oldSheetID);
    if (!oldSpreadsheet) {
      return {
        success: false,
        message: `Old spreadsheet not found with ID: ${oldSheetID}`,
      };
    }

    var sheetTypeFunction = sheetVars(sheetType);

    if (!sheetTypeFunction) {
      console.log(`Sheet type function not found for: ${sheetType}`);
      return {
        success: false,
        message: `Sheet type function not found for: ${sheetType}`,
      };
    }

    console.log(
      `Importing data for sheet type: ${sheetType} with ID: ${newSheetID}`
    );

    var result = sheetTypeFunction.importData(versionDifference);
    if (!result || !result.success) {
      return {
        success: false,
        message: `Error importing data for ${sheetType}: ${result.message}`,
      };
    }
    
    var isImportedInfo = shared.findSheetTypeID(newSheetID, "IDS");
    if (!isImportedInfo || !isImportedInfo.isImported) {
      return {
        success: false,
        message: "Can not find import cell in the new IDS sheet.",
      };
    }
    
    try {
      SheetsAPI.setValue(newSheetID, isImportedInfo.isImported.range, "✅");
    } catch (error) {
      console.log(`Error updating imported status:  ${error.toString()}`);
      return {
        success: false,
        message: `Error updating imported status:  ${error.toString()}`,
      };
    }

    return {
      success: true,
      message: `Import of ${sheetType} data completed successfully.`,
      updated: true,
    };
  } catch (error) {
    console.log("Error during import: " + error.message);
    return {
      success: false,
      message: "Error during import: " + error.message,
    };
  }
}