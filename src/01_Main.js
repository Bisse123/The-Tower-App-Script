const SheetsAPI = {
  // Get spreadsheet metadata and sheets
  getSpreadsheet: function (spreadsheetId) {
    try {
      return Sheets.Spreadsheets.get(spreadsheetId);
    } catch (error) {
      console.error("Error getting spreadsheet:", error);
      return null;
    }
  },

  // Get sheet by name from spreadsheet
  getSheetByName: function (spreadsheetId, sheetName) {
    try {
      const spreadsheet = this.getSpreadsheet(spreadsheetId);
      const sheet = spreadsheet.sheets.find(
        (s) => s.properties.title === sheetName
      );
      return sheet ? sheet.properties : null;
    } catch (error) {
      console.error("Error getting sheet by name:", error);
      return null;
    }
  },

  // Check if a sheet exists in a spreadsheet
  hasSheet: function (spreadsheetId, sheetName) {
    try {
      return this.getSheetByName(spreadsheetId, sheetName) !== null;
    } catch (error) {
      console.error("Error checking sheet existence:", error);
      return false;
    }
  },

  // Get values from a range
  getValues: function (spreadsheetId, range) {
    try {
      const response = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
      return response.values;
    } catch (error) {
      console.error("Error getting values:", error);
      return null;
    }
  },

  // Get a single value from a cell
  getValue: function (spreadsheetId, range) {
    try {
      const values = this.getValues(spreadsheetId, range);
      return values && values.length > 0 && values[0].length > 0 ? values[0][0] : null;
    } catch (error) {
      console.error("Error getting single value:", error);
      return null;
    }
  },

  // Set multiple values in a range
  setValues: function (spreadsheetId, range, values) {
    try {
      const requestBody = {
        values: values,
      };
      return Sheets.Spreadsheets.Values.update(
        requestBody,
        spreadsheetId,
        range,
        {
          valueInputOption: "USER_ENTERED",
        }
      );
    } catch (error) {
      console.error("Error setting values:", error);
      return null;
    }
  },

  // Set a single value in a cell
  setValue: function (spreadsheetId, range, value) {
    try {
      const requestBody = {
        values: [[value]],
      };
      return Sheets.Spreadsheets.Values.update(
        requestBody,
        spreadsheetId,
        range,
        {
          valueInputOption: "USER_ENTERED",
        }
      );
    } catch (error) {
      console.error("Error setting value:", error);
      return null;
    }
  },

  // Get all data from a sheet
  getDataRange: function (spreadsheetId, sheetName) {
    try {
      return this.getValues(spreadsheetId, sheetName);
    } catch (error) {
      console.error("Error getting data range:", error);
      return null;
    }
  },
  batchUpdateValues: function (spreadsheetId, updates) {
    try {
      const data = updates.map(update => ({
        range: update.range,
        values: update.values
      }));
      const requestBody = {
        data: data,
        valueInputOption: "USER_ENTERED"
      };
      return Sheets.Spreadsheets.Values.batchUpdate(requestBody, spreadsheetId);
    } catch (error) {
      console.error("Error in batchUpdateValues:", error);
      return null;
    }
  },
};

const sheetVars = (sheetType) => {
  var sheetTypeFuntions = {
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
  return sheetTypeFuntions[sheetType];
};

function startImportData(sheetType, sheetID) {
  console.log(
    `Starting import for sheet type: ${sheetType} with ID: ${sheetID}`
  );
  try {
    var sheetTypeFunction = sheetVars(sheetType);

    var spreadsheet = SheetsAPI.getSpreadsheet(sheetID);
    if (!spreadsheet) {
      return {
        success: false,
        message: `Spreadsheet not found with ID: ${sheetID}`,
      };
    }
    if (!sheetTypeFunction) {
      console.log(`Sheet type function not found for: ${sheetType}`);
      return {
        success: false,
        message: `Sheet type function not found for: ${sheetType}`,
      };
    }
    if (!sheetType) {
      console.log("Sheet type is not defined.");
      return {
        success: false,
        message: "Sheet type is not defined.",
      };
    }

    console.log(
      `Importing data for sheet type: ${sheetType} with ID: ${sheetID}`
    );
    var result = sheetTypeFunction.importData(sheetType, sheetID);
    if (!result || !result.success) {
      return {
        success: false,
        message: `Error importing data for ${sheetType}: ${result.message}`,
      };
    }
    
    var isImportedInfo = shared.findSheetTypeID(sheetID, "IDS");
    if (!isImportedInfo || !isImportedInfo.isImported) {
      return {
        success: false,
        message: "Is Imported range not found in the new IDS sheet.",
      };
    }
    
    try {
      SheetsAPI.setValue(sheetID, isImportedInfo.isImported.range, "✅");
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

function importData(newSheetID, oldSheetID, idMasterID, sheetType) {
  // Check if new spreadsheet exists
  var newSpreadsheet = SheetsAPI.getSpreadsheet(newSheetID);
  if (!newSpreadsheet) {
    return {
      success: false,
      message: "New spreadsheet not found with ID: " + newSheetID,
    };
  }

  // Check if IDS sheet exists
  var newIdSheet = SheetsAPI.getSheetByName(newSheetID, "IDS");
  if (!newIdSheet) {
    return {
      success: false,
      message: "IDS sheet not found in the new spreadsheet.",
    };
  }

  // Check if ID Master spreadsheet exists
  var idMasterSpreadsheet = SheetsAPI.getSpreadsheet(idMasterID);
  if (!idMasterSpreadsheet) {
    return {
      success: false,
      message: "IDS Master Spreadsheet not found with ID: " + idMasterID,
    };
  }

  var idMasterSheet = SheetsAPI.getSheetByName(idMasterID, "IDS");
  if (!idMasterSheet) {
    return {
      success: false,
      message: "IDS Master Sheet not found in the IDS Master Spreadsheet.",
    };
  }

  // Check if old spreadsheet exists
  var oldSpreadsheet = SheetsAPI.getSpreadsheet(oldSheetID);
  if (!oldSpreadsheet) {
    return {
      success: false,
      message: "Linked old spreadsheet not found with ID: " + oldSheetID,
    };
  }

  return startImportData(sheetType, newSheetID);
}

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
