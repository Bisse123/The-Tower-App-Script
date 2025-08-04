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
    "IDS Collection - all IDS-Sheets on one file": collection,
    "IDS Master": master,
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
      console.log(`No spreadsheet name provided.`);
      return null;
    }
    if (storedSpreadsheets[spreadsheet]) {
      return storedSpreadsheets[spreadsheet];
    }
    if (!sheetID) {
      console.log(`Spreadsheet not defined and no sheet ID provided.`);
      return null;
    }
    var spreadsheetInfo = SheetsAPI.getSpreadsheet(sheetID);
    if (!spreadsheetInfo) {
      console.log(`Spreadsheet not found with ID: ${sheetID}`);
      return null;
    }
    storedSpreadsheets[spreadsheet] = spreadsheetInfo;
    return spreadsheetInfo;
  };
})();

function doGet(e) {
  // console.log(`doGet called with parameters: ${JSON.stringify(e.parameter)}`);
  var template = HtmlService.createTemplateFromFile("14_WebApp");
  if (e.parameter.newSheetID === "<Script loading...>") {
    e.parameter.newSheetID = "";
  }
  
  var newSheetID = e.parameter.newSheetID || "";
  var oldSheetID = e.parameter.oldSheetID || "";
  var idMasterID = e.parameter.idMasterID || "";
  var sheetType = e.parameter.sheetType || "";

  // Special case for IDS Master - same logic as showImportDialog
  if (sheetType === "IDS Master") {
    // For IDS Master accessed via doGet, we need to determine which spreadsheet is the IDS Master
    // This could come from a parameter or we might need to infer it
    if (!idMasterID && newSheetID) {
      // If no idMasterID provided but newSheetID is provided, assume newSheetID is the IDS Master
      idMasterID = newSheetID;
    }
    
    console.log("IDS Master detected in doGet, setting parameters accordingly");
    template.newSheetID = "";  // IDS Master doesn't have predefined new/old sheets
    template.oldSheetID = "";
    template.idMasterID = idMasterID;  // The IDS Master spreadsheet ID
    template.sheetType = sheetType;
  } else {
    // Regular processing for individual sheet types
    template.newSheetID = newSheetID;
    template.oldSheetID = oldSheetID;
    template.idMasterID = idMasterID;
    template.sheetType = sheetType;
  }

  template.API_KEY =
    PropertiesService.getScriptProperties().getProperty("API_KEY");
  template.APP_ID =
    PropertiesService.getScriptProperties().getProperty("APP_ID");

  template.viewType = "webapp";
  template.accessRequired = false;

  return template
    .evaluate()
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setTitle("Import Data");
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen(e) {
  // console.log(`onOpen called with event: ${JSON.stringify(e)}`);
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Import Data")
      .addItem("Get Started", "showGetStartedDialog")
      .addItem("Import Data", "showImportDialog")
      .addToUi();
  } catch (error) {}
}

function onInstall(e) {
  // console.log(`onOpen called with event: ${JSON.stringify(e)}`);
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Import Data")
      .addItem("Get Started", "showGetStartedDialog")
      .addItem("Import Data", "showImportDialog")
      .addToUi();
  } catch (error) {}
}

function showGetStartedDialog() {
  try {
    var template = HtmlService.createTemplateFromFile("14_getStartedApp");
    var html = template
      .evaluate()
      .setWidth(1200)
      .setHeight(700)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
    SpreadsheetApp.getUi().showModalDialog(html, "Get Started");
  } catch (error) {
    console.log(`Error in showGetStartedDialog: ${error.message}`);
    SpreadsheetApp.getUi().alert("Error: " + error.message);
  }
}

function showImportDialog() {
  try {
    var newSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var newSheetID = newSpreadsheet.getId();
    var sheetType = newSpreadsheet
      .getSheetByName("Home Page")
      .getRange("B2")
      .getValue();
    if (!sheetVars(sheetType)) {
      console.log(`Sheet type not found in the new spreadsheet.`);
      throw new Error("Sheet type not found in the new spreadsheet.");
    }

    // Special case for IDS Master
    if (sheetType === "IDS Master") {
      console.log("IDS Master detected, showing import dialog with limited parameters");
      var template = HtmlService.createTemplateFromFile("14_WebApp");
      template.newSheetID = "";  // IDS Master doesn't have predefined new/old sheets
      template.oldSheetID = "";
      template.idMasterID = newSheetID;  // The current spreadsheet IS the IDS Master
      template.sheetType = sheetType;
      template.API_KEY =
        PropertiesService.getScriptProperties().getProperty("API_KEY");
      template.APP_ID =
        PropertiesService.getScriptProperties().getProperty("APP_ID");

      template.viewType = "sidebar";
      template.accessRequired = false;

      var html = template
        .evaluate()
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setTitle("Import Data - IDS Master");
      SpreadsheetApp.getUi().showSidebar(html);
      return;
    }

    // Regular processing for individual sheet types
    var sheet = newSpreadsheet.getSheetByName("IDS");
    if (!sheet) {
      console.log(`IDS sheet not found in the new spreadsheet.`);
      throw new Error("IDS sheet not found in the new spreadsheet.");
    }
    try {
      var values = sheet.getDataRange().getValues();
      if (!values || values.length === 0) {
        console.log(`No data found in the IDS sheet.`);
        throw new Error("No data found in the IDS sheet.");
      }
      var idMasterURL = "";
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          if (
            typeof values[row][col] === "string" &&
            values[row][col].indexOf("IDS Master's ID") !== -1 &&
            values[row][col].indexOf("script") === -1
          ) {
            var idMasterURL = values[row][col + 2];
            break;
          }
        }
        if (idMasterURL) {
          break;
        }
      }
      if (!idMasterURL) {
        console.log(`IDS Master's ID not found in the new spreadsheet.`);
        throw new Error("IDS Master's ID not found in the new spreadsheet.");
      }
      var idMasterID = idMasterURL ? shared.extractSheetId(idMasterURL) : "";
      if (!idMasterID) {
        console.log(`IDS Master's ID not found in the new spreadsheet.`);
        throw new Error("IDS Master's ID not found in the new spreadsheet.");
      }
      var oldSheetUrl = "";

      var idsSubsheet = newSpreadsheet.getSheetByName("_IDS");
      if (idsSubsheet) {
        var lastColumn = idsSubsheet.getLastColumn();
        if (lastColumn > 0) {
          var firstRowValues = idsSubsheet
            .getRange(1, 1, 1, lastColumn)
            .getValues()[0];

          var sheetTypeMapping = {
            Laboratory: "Labs",
            Workshop: "WS",
            "Ultimate Weapon": "UWs",
            "Themes & Songs": "Themes & Songs",
            Bots: "Bots",
            Relics: "Relics",
            Vault: "Vault",
            Cards: "Cards",
            Modules: "Modules",
            Guardians: "Guardians",
          };

          var searchValue = sheetTypeMapping[sheetType] || sheetType;

          var foundIndex = firstRowValues.indexOf(searchValue);
          if (foundIndex !== -1 && foundIndex < firstRowValues.length - 1) {
            var foundValue = firstRowValues[foundIndex + 1];
            if (foundValue) {
              oldSheetUrl = foundValue;
            }
          }
        }
      }

      if (!oldSheetUrl) {
        console.log(`Old sheet URL not found in the _IDS subsheet.`);
        throw new Error("Old sheet URL not found in the _IDS subsheet.");
      }
      var oldSheetID = oldSheetUrl ? shared.extractSheetId(oldSheetUrl) : "";
      if (!oldSheetID) {
        console.log(`Old sheet ID not found in the _IDS subsheet.`);
        throw new Error("Old sheet ID not found in the _IDS subsheet.");
      }

      var template = HtmlService.createTemplateFromFile("14_WebApp");
      template.newSheetID = newSheetID;
      template.oldSheetID = oldSheetID;
      template.idMasterID = idMasterID;
      template.sheetType = sheetType;
      template.API_KEY =
        PropertiesService.getScriptProperties().getProperty("API_KEY");
      template.APP_ID =
        PropertiesService.getScriptProperties().getProperty("APP_ID");

      template.viewType = "sidebar";
      template.accessRequired = false;

      var html = template
        .evaluate()
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setTitle("Import Data");
      SpreadsheetApp.getUi().showSidebar(html);
    } catch (error) {
      console.log(`Error in showImportDialog: ${error.message}`);
      var template = HtmlService.createTemplateFromFile("14_WebApp");
      template.newSheetID = newSheetID;
      template.oldSheetID = "";
      template.idMasterID = "";
      template.sheetType = "";
      template.API_KEY =
        PropertiesService.getScriptProperties().getProperty("API_KEY");
      template.APP_ID =
        PropertiesService.getScriptProperties().getProperty("APP_ID");

      template.viewType = "sidebar";
      template.accessRequired = true;

      var html = template
        .evaluate()
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setTitle("Import Data");
      SpreadsheetApp.getUi().showSidebar(html);
    }
  } catch (error) {
    console.log(`Error in showImportDialog: ${error.message}`);
    var template = HtmlService.createTemplateFromFile("14_WebApp");
    template.newSheetID = "";
    template.oldSheetID = "";
    template.idMasterID = "";
    template.sheetType = "";
    template.API_KEY =
      PropertiesService.getScriptProperties().getProperty("API_KEY");
    template.APP_ID =
      PropertiesService.getScriptProperties().getProperty("APP_ID");

    template.viewType = "sidebar";
    template.accessRequired = false;

    var html = template
      .evaluate()
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setTitle("Import Data");
    SpreadsheetApp.getUi().showSidebar(html);
  }
}

function importData(
  newSheetID,
  oldSheetID,
  idMasterID,
  sheetType,
  versionDifference
) {
  // console.log(`Starting import for sheet type: ${sheetType} with ID: ${newSheetID}`);
  try {
    if (!sheetType) {
      console.log(`Sheet type is not defined.`);
      return {
        success: false,
        message: "Sheet type is not defined.",
      };
    }
    
    if (!newSheetID || !oldSheetID || !idMasterID) {
      console.log(`One or more required IDs are missing.`);
      return {
        success: false,
        message: "One or more required IDs are missing.",
      };
    }
    
    var newSpreadsheet = spreadsheets(`${sheetType} newSpreadsheet`, newSheetID);
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`,
      };
    }

    if (!SheetsAPI.getSheetByName(newSpreadsheet, "IDS")) {
      console.log(`IDS sheet not found in the new ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `IDS sheet™ not found in the new ${sheetType} spreadsheet™.`,
      };
    }

    var idMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
    if (!idMasterSpreadsheet) {
      console.log(`IDS Master Spreadsheet not found with ID: ${idMasterID}`);
      return {
        success: false,
        message: `IDS Master Spreadsheet™ not found with ID: ${idMasterID}`,
      };
    }

    if (!SheetsAPI.getSheetByName(idMasterSpreadsheet, "IDS")) {
      console.log(`IDS sheet not found in the IDS Master Spreadsheet`);
      return {
        success: false,
        message: `IDS sheet™ not found in the IDS Master Spreadsheet™`,
      };
    }

    var oldSpreadsheet = spreadsheets(`${sheetType} oldSpreadsheet`, oldSheetID);
    if (!oldSpreadsheet) {
      console.log(`Old spreadsheet not found with ID: ${oldSheetID}`);
      return {
        success: false,
        message: `Old spreadsheet™ not found with ID: ${oldSheetID}`,
      };
    }

    var sheetTypeFunction = sheetVars(sheetType);

    if (!sheetTypeFunction) {
      console.log(`Sheet type function not found for: ${sheetType}`);
      return {
        success: false,
        message: `Sheet™ type function not found for: ${sheetType}`,
      };
    }

    var result = sheetTypeFunction.importData(versionDifference);
    if (!result || !result.success) {
      console.log(
        `Error importing data for ${sheetType}: ${
          result ? result.message : "Unknown error"
        }`
      );
      return {
        success: false,
        message: `Error importing data for ${sheetType}: ${
          result && result.message ? result.message : "Unknown error"
        }`,
      };
    }

    return {
      success: true,
      message: `Import of ${sheetType} data completed successfully.`,
      updated: true,
    };
  } catch (error) {
    console.log(`Error during import: ${error.message}`);
    return {
      success: false,
      message: `Error during import: ${error.message}`,
    };
  }
}
