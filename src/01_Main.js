const sheetVars = (sheetType) => {
  var sheetTypeFunctions = {
    Laboratory: lab,
    Workshop: workshop,
    "Ultimate Weapon": ultimate,
    "Themes & Songs": themes,
    Bots: bots,
    Relics: relics,
    Vault: vault,
    Cards: cards,
    Modules: modules,
    Guardians: guardians,
  };
  return sheetTypeFunctions[sheetType];
};

function doGet(e) {
  console.log("doGet called with parameters: " + JSON.stringify(e.parameter));
  var template = HtmlService.createTemplateFromFile("WebApp");
  template.newSheetID = e.parameter.newSheetID;
  template.oldSheetID = e.parameter.oldSheetID;
  template.idMasterID = e.parameter.idMasterID;
  template.sheetType = e.parameter.sheetType;

  return template
    .evaluate()
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setTitle("Import Data");
}

function importData(newSheetID, oldSheetID, idMasterID, sheetType) {
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
    var newSpreadsheet = SheetsAPI.getSpreadsheet(newSheetID);
    if (!newSpreadsheet) {
      return {
        success: false,
        message: `New spreadsheet not found with ID: ${newSheetID}`,
      };
    }

    // Check if IDS sheet exists
    if (!SheetsAPI.getSheetByName(newSheetID, "IDS")) {
      return {
        success: false,
        message:`IDS sheet not found in the new ${sheetType} spreadsheet.`,
      };
    }

    if (!SheetsAPI.hasSheet(newSheetID, "EXPORT") && !SheetsAPI.hasSheet(newSheetID, "STATS")) {
      console.log(`Export sheet not found in new ${sheetType} spreadsheet`);
      return {
        success: false,
        message: `Export sheet not found in new ${sheetType} spreadsheet`,
      };
    }

    // Check if ID Master spreadsheet exists
    if (!SheetsAPI.getSpreadsheet(idMasterID)) {
      return {
        success: false,
        message: `IDS Master Spreadsheet not found with ID: ${idMasterID}`,
      };
    }

    if (!SheetsAPI.getSheetByName(idMasterID, "IDS")) {
      return {
        success: false,
        message: `IDS sheet not found in the IDS Master Spreadsheet.`,
      };
    }

    // Check if old spreadsheet exists
    if (!SheetsAPI.getSpreadsheet(oldSheetID)) {
      return {
        success: false,
        message: `New spreadsheet not found with ID: ${oldSheetID}`,
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

    var result = sheetTypeFunction.importData(newSheetID, oldSheetID);
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