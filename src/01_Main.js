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
  var template = HtmlService.createTemplateFromFile("13_WebApp");
  if (e.parameter.newSheetID === "<Script loading...>") {
    e.parameter.newSheetID = "";
  }
  template.newSheetID = e.parameter.newSheetID;
  template.oldSheetID = e.parameter.oldSheetID;
  template.idMasterID = e.parameter.idMasterID;
  template.sheetType = e.parameter.sheetType;
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
    var template = HtmlService.createTemplateFromFile("13_getStartedApp")
    var html = template
      .evaluate()
      .setWidth(1200)
      .setHeight(700)
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
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
    var sheet = newSpreadsheet.getSheetByName("IDS");
    if (!sheet) {
      console.log(`IDS sheet not found in the new spreadsheet.`);
      throw new Error("IDS sheet not found in the new spreadsheet.");
    }
    try {
      var values = sheet.getDataRange().getValues();
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
      var idMasterID = idMasterURL ? shared.extractSheetId(idMasterURL) : "";
      var oldSheetInfo = idMasterID ? shared.findSheetTypeID(idMasterID, "IDS", sheetType + " ID") : "";
      var oldSheetID = oldSheetInfo ? shared.extractSheetId(oldSheetInfo.id) : "";

      var template = HtmlService.createTemplateFromFile("13_WebApp");
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
      var template = HtmlService.createTemplateFromFile("13_WebApp");
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
  }
  catch (error) {
    console.log(`Error in showImportDialog: ${error.message}`);
    var template = HtmlService.createTemplateFromFile("13_WebApp");
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
    // Check if new spreadsheet exists
    var newSpreadsheet = spreadsheets("newSpreadsheet", newSheetID);
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

    if (
      !SheetsAPI.getSheetByName(newSpreadsheet, "EXPORT") &&
      !SheetsAPI.getSheetByName(newSpreadsheet, "STATS")
    ) {
      console.log(`Export sheet not found in new ${sheetType} spreadsheet`);
      return {
        success: false,
        message: `Export sheet™ not found in new ${sheetType} spreadsheet™`,
      };
    }

    var newSheetInfo = shared.findSheetTypeID(newSheetID, "IDS");
    if (
      !newSheetInfo ||
      !newSheetInfo.accessStatus ||
      newSheetInfo.accessStatus.value !== "✅"
    ) {
      console.log(`New sheet has not been granted access to IDS Master.`);
      return {
        success: false,
        message: `New sheet™ has not been granted access to IDS Master.`,
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
      console.log(`IDS sheet not found in the IDS Master Spreadsheet.`);
      return {
        success: false,
        message: `IDS sheet™ not found in the IDS Master Spreadsheet™.`,
      };
    }
    var idMasterInfo = shared.findSheetTypeID(
      idMasterID,
      "IDS",
      sheetType + " ID"
    );
    if (
      !idMasterInfo ||
      !idMasterInfo.accessStatus ||
      !["✅", "Wrong ID or Version"].includes(idMasterInfo.accessStatus.value)
    ) {
      console.log(
        `IDS Master has not granted access to the old ${sheetType} sheet.`
      );
      return {
        success: false,
        message: `IDS Master has not granted access to the old ${sheetType} sheet™.`,
      };
    }

    var oldSpreadsheet = spreadsheets("oldSpreadsheet", oldSheetID);
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

    // console.log(`Importing data for sheet type: ${sheetType} with ID: ${newSheetID}`);

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

    try {
      SheetsAPI.setValue(newSheetID, newSheetInfo.importStatus.range, "✅");
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
    console.log(`Error during import: ${error.message}`);
    return {
      success: false,
      message: `Error during import: ${error.message}`,
    };
  }
}
