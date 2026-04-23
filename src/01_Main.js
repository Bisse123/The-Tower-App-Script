//Testing auto redeploy
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
    "Player & Stuff": playerStuff,
    "IDS Collection": collection,
    "IDS Master": master,
    "Effective Paths": ePaths,
  };
  return sheetTypeFunctions[sheetType];
};

const spreadsheets = (() => {
  // Wrapper function to maintain backward compatibility with CacheManager
  return function (spreadsheetTypeName, sheetID) {
    if (!spreadsheetTypeName) {
      console.log(`No spreadsheet type name provided.`);
      return null;
    }

    // Use CacheManager to get or fetch spreadsheet metadata
    // CacheManager handles:
    // 1. Checking cache if sheetID matches
    // 2. Overwriting cache if sheetID doesn't match
    // 3. Using previously cached sheetID if none provided
    const result = CacheManager.getSpreadsheet(spreadsheetTypeName, sheetID);

    if (!result) {
      if (!sheetID) {
        console.log(
          `Spreadsheet not found in cache and no sheet ID provided for: ${spreadsheetTypeName}`
        );
      } else {
        console.log(`Spreadsheet not found with ID: ${sheetID}`);
      }
      return null;
    }

    return result;
  };
})();

const ADDON_CONSENT_READY_SIGNAL_KEY = "ADDON_CONSENT_READY_SIGNAL";

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  // console.log(`doGet called with parameters: ${JSON.stringify(params)}`);

  var targetPage = (params.page || "").toString().toLowerCase();
  var openGetStarted =
    targetPage === "getstarted" ||
    targetPage === "get-started" ||
    params.getStarted === "true";

  if (openGetStarted) {
    var getStartedTemplate = HtmlService.createTemplateFromFile("20_getStartedApp");
    getStartedTemplate.API_KEY =
      PropertiesService.getScriptProperties().getProperty("API_KEY");
    getStartedTemplate.APP_ID =
      PropertiesService.getScriptProperties().getProperty("APP_ID");
    getStartedTemplate.viewType = "webapp";
    return getStartedTemplate
      .evaluate()
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setTitle("Get Started");
  }

  var template = HtmlService.createTemplateFromFile("20_WebApp");
  if (params.newSheetID === "<Script loading...>") {
    params.newSheetID = "";
  }
  var newSheetID = params.newSheetID || "";
  var oldSheetID = params.oldSheetID || "";
  var idMasterID = params.idMasterID || "";
  var sheetType = params.sheetType || "";

  // Special case for IDS Master - same logic as showImportDialog
  if (sheetType === "IDS Master") {
    // For IDS Master accessed via doGet, we need to determine which spreadsheet is the IDS Master
    // This could come from a parameter or we might need to infer it
    if (!oldSheetID && idMasterID) {
      // If no idMasterID provided but newSheetID is provided, assume newSheetID is the IDS Master
      oldSheetID = idMasterID;
    }

    console.log("IDS Master detected in doGet, setting parameters accordingly");
    template.newSheetID = "";
    template.oldSheetID = oldSheetID;
    template.idMasterID = "";
    template.sheetType = sheetType;
  } else {
    if (typeof sheetType === "string" && sheetType.includes("IDS Collection")) {
      sheetType = "IDS Collection";
    }
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
  createMenu();
}

function onInstall(e) {
  createMenu();
}

function createMenu() {
    try {
        var ui = SpreadsheetApp.getUi();
        ui.createMenu("Import Data")
            .addItem("Get Started", "showGetStartedDialog")
            // .addItem("Import Data", "showImportDialog")
            .addItem("Update Sheet", "showUpdateDialog")
            .addToUi();
    } catch (error) {}
}

function showGetStartedDialog() {
  try {
    var template = HtmlService.createTemplateFromFile("20_getStartedApp");
    template.API_KEY =
      PropertiesService.getScriptProperties().getProperty("API_KEY");
    template.APP_ID =
      PropertiesService.getScriptProperties().getProperty("APP_ID");
    template.viewType = "sidebar";
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

    // Special case for IDS Master
    if (sheetType === "IDS Master") {
      console.log(
        "IDS Master detected, showing import dialog with limited parameters"
      );
      var template = HtmlService.createTemplateFromFile("20_WebApp");
      template.newSheetID = "";
      template.oldSheetID = newSheetID;
      template.idMasterID = "";
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
    if (sheetType === "IDS Collection - all IDS-Sheets on one file") {
      sheetType = "IDS Collection";
    }

    if (!sheetVars(sheetType)) {
      console.log(`Sheet type not found in the new spreadsheet.`);
      throw new Error("Sheet type not found in the new spreadsheet.");
    }

    var sheetName = "IDS";
    var searchValue = "IDS Master's ID";
    if (sheetType === "IDS Collection") {
      sheetName = "Home Page";
      searchValue = "Load your file here";
    }
    // Regular processing for individual sheet types
    var sheet = newSpreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      console.log(`IDS sheet not found in the new spreadsheet.`);
      throw new Error("IDS sheet not found in the new spreadsheet.");
    }
    try {
      var sheetIDs = findSheetIDs(sheet, sheetType, searchValue);

      var oldSheetID = sheetIDs.oldSheetID;
      var idMasterID = sheetIDs.idMasterID;

      var template = HtmlService.createTemplateFromFile("20_WebApp");
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
      var template = HtmlService.createTemplateFromFile("20_WebApp");
      template.newSheetID = newSheetID;
      template.oldSheetID = "";
      template.idMasterID = "";
      template.sheetType = sheetType;
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
    var template = HtmlService.createTemplateFromFile("20_WebApp");
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

function showUpdateDialog() {
  try {
    var template = HtmlService.createTemplateFromFile("20_WebApp");
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
  } catch (error) {
    console.log(`Error in showUpdateDialog: ${error.message}`);
    var template = HtmlService.createTemplateFromFile("20_WebApp");
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

function showAddonConsentDialog(authorizationUrl) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(ADDON_CONSENT_READY_SIGNAL_KEY);

  var template = HtmlService.createTemplateFromFile("28_addon_consent_dialog");
  template.authorizationUrl = authorizationUrl || "";

  var html = template
    .evaluate()
    .setWidth(560)
    .setHeight(280)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");

  SpreadsheetApp.getUi().showModalDialog(html, "Additional Permissions Required");
}

function markAddonConsentReadySignal() {
  PropertiesService.getUserProperties().setProperty(
    ADDON_CONSENT_READY_SIGNAL_KEY,
    String(Date.now())
  );
  return true;
}

function consumeAddonConsentReadySignal() {
  var userProperties = PropertiesService.getUserProperties();
  var signal = userProperties.getProperty(ADDON_CONSENT_READY_SIGNAL_KEY);

  if (!signal) {
    return false;
  }

  userProperties.deleteProperty(ADDON_CONSENT_READY_SIGNAL_KEY);
  return true;
}

function getUpdateDialogParameters() {
  try {
    var oldSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var oldSheetID = oldSpreadsheet.getId();

    var homePageSheet = oldSpreadsheet.getSheetByName("Home Page");
    if (!homePageSheet) {
      throw new Error("Home Page sheet not found in the active spreadsheet.");
    }

    var sheetType = "";

    if (oldSpreadsheet.getSheetByName("eHP") || oldSpreadsheet.getSheetByName("eDamage") || oldSpreadsheet.getSheetByName("eEcon")) {
      sheetType = "Effective Paths";
    } else {
      sheetType = homePageSheet.getRange("B2").getValue();
    }
    
    if (!sheetType) {
      throw new Error("Sheet type not found in the active spreadsheet.");
    }

    if (sheetType === "IDS Collection - all IDS-Sheets on one file") {
      sheetType = "IDS Collection";
    }

    if (sheetType === "IDS Master") {
      return {
        success: true,
        newSheetID: "",
        oldSheetID: oldSheetID,
        idMasterID: "",
        sheetType: sheetType,
        accessRequired: false,
      };
    }

    if (!sheetVars(sheetType)) {
      throw new Error("Sheet type not found in the active spreadsheet.");
    }

    var sheetName = "IDS";
    var searchValue = "IDS Master's ID";
    if (sheetType === "IDS Collection") {
      sheetName = "Home Page";
      searchValue = "Load your file here";
    }

    var sheet = oldSpreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("IDS sheet not found in the active spreadsheet.");
    }

    try {
      var sheetIDs = findSheetIDs(sheet, sheetType, searchValue);

      return {
        success: true,
        newSheetID: "",
        oldSheetID: oldSheetID,
        idMasterID: sheetIDs.idMasterID,
        sheetType: sheetType,
        accessRequired: false,
      };
    } catch (error) {
      return {
        success: true,
        newSheetID: "",
        oldSheetID: oldSheetID,
        idMasterID: "",
        sheetType: sheetType,
        accessRequired: true,
        message: error.message || error.toString(),
      };
    }
  } catch (error) {
    console.log(`Error in getUpdateDialogParameters: ${error.message}`);
    return {
      success: false,
      newSheetID: "",
      oldSheetID: "",
      idMasterID: "",
      sheetType: "",
      accessRequired: true,
      message: error.message || error.toString(),
    };
  }
}

function exportData(oldSheetID, sheetType, versionDifference) {
  try {
    if (!sheetType) {
      console.log(`Sheet type is not defined.`);
      return {
        success: false,
        message: "Sheet type is not defined.",
      };
    }

    if (!oldSheetID) {
      console.log(`Old sheet ID is missing.`);
      return {
        success: false,
        message: "Old sheet ID is missing.",
      };
    }

    var oldSpreadsheet = spreadsheets(
      `${sheetType} oldSpreadsheet`,
      oldSheetID
    );
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

    // Extract sheet visibility information from old spreadsheet
    var sheetVisibility = {};
    if (oldSpreadsheet && oldSpreadsheet.sheets) {
      for (var i = 0; i < oldSpreadsheet.sheets.length; i++) {
        var sheet = oldSpreadsheet.sheets[i];
        sheetVisibility[sheet.properties.title] =
          sheet.properties.hidden || false;
      }
    }

    // Export the data from the old spreadsheet
    var exportResult = sheetTypeFunction.exportData(versionDifference);
    if (!exportResult || !exportResult.success) {
      console.log(
        `Error exporting data for ${sheetType}: ${
          exportResult ? exportResult.message : "Unknown error"
        }`
      );
      return {
        success: false,
        message: `Error exporting data for ${sheetType}: ${
          exportResult && exportResult.message
            ? exportResult.message
            : "Unknown error"
        }`,
      };
    }

    return {
      success: true,
      message: `Export of ${sheetType} data completed successfully.`,
      data: exportResult.data,
      sheetVisibility: sheetVisibility,
    };
  } catch (error) {
    console.log(`Error during export: ${error.message}`);
    return {
      success: false,
      message: `Error during export: ${error.message}`,
    };
  }
}

function importData(newSheetID, sheetType, data, sheetVisibility, idMasterID) {
  try {
    if (!sheetType) {
      console.log(`Sheet type is not defined.`);
      return {
        success: false,
        message: "Sheet type is not defined.",
      };
    }

    if (!newSheetID) {
      console.log(`New sheet ID is missing.`);
      return {
        success: false,
        message: "New sheet ID is missing.",
      };
    }

    if (!data) {
      console.log(`Data to import is missing.`);
      return {
        success: false,
        message: "Data to import is missing.",
      };
    }

    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID
    );
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`,
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

    // Apply sheet visibility using the provided visibility data
    if (sheetVisibility && Object.keys(sheetVisibility).length > 0) {
      var hideShowSheetsResult = SheetsAPI.applySheetVisibility(
        newSpreadsheet,
        sheetVisibility
      );
      if (!hideShowSheetsResult || !hideShowSheetsResult.success) {
        console.log(
          `Error updating sheet visibility: ${
            hideShowSheetsResult
              ? hideShowSheetsResult.message
              : "Unknown error"
          }`
        );
        return {
          success: false,
          message: `Error updating sheet visibility: ${
            hideShowSheetsResult && hideShowSheetsResult.message
              ? hideShowSheetsResult.message
              : "Unknown error"
          }`,
        };
      }
    }

    // Add idMasterID to data if provided
    if (idMasterID) {
      data.idMasterID = idMasterID;
    }

    // Import the data to the new spreadsheet
    var importResult = sheetTypeFunction.importData(data);
    if (!importResult || !importResult.success) {
      console.log(
        `Error importing data for ${sheetType}: ${
          importResult ? importResult.message : "Unknown error"
        }`
      );
      return {
        success: false,
        message: `${sheetType}: ${
          importResult && importResult.message
            ? importResult.message
            : "Unknown error"
        }`,
        failedUpdates: importResult.failedUpdates || [],
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

function findSheetIDs(sheet, sheetType, searchValue) {
  try {
    var values = sheet.getDataRange().getValues();
    if (!values || values.length === 0) {
      console.log(`No data found in the sheet.`);
      throw new Error("No data found in the sheet.");
    }

    // Find IDS Master ID (common for both cases)
    var idMasterURL = "";
    for (var row = 0; row < values.length; row++) {
      for (var col = 0; col < values[row].length; col++) {
        if (
          typeof values[row][col] === "string" &&
          values[row][col].indexOf(searchValue) !== -1 &&
          values[row][col].indexOf("script") === -1
        ) {
          idMasterURL = values[row][col + 2];
          break;
        }
      }
      if (idMasterURL) {
        break;
      }
    }

    if (!idMasterURL) {
      console.log(`IDS Master's ID not found in the sheet.`);
      throw new Error("IDS Master's ID not found in the sheet.");
    }

    var idMasterID = idMasterURL ? shared.extractSheetId(idMasterURL) : "";
    if (!idMasterID) {
      console.log(`IDS Master's ID could not be extracted from the sheet.`);
      throw new Error("IDS Master's ID could not be extracted from the sheet.");
    }

    // Special case for "Load your file here" - return idMasterID as oldSheetID
    if (sheetType === "IDS Collection") {
      return { oldSheetID: idMasterID, idMasterID: "" };
    }

    if (sheetType === "Effective Paths") {
      return { oldSheetID: "", idMasterID: idMasterID };
    }

    // Regular case - find oldSheetID from _IDS subsheet
    var oldSheetUrl = "";
    var newSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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
          "Player & Stuff": "Player & Stuff",
        };
        var foundIndex = firstRowValues.indexOf(sheetTypeMapping[sheetType]);
        if (foundIndex !== -1 && foundIndex < firstRowValues.length - 1) {
          var foundValue = firstRowValues[foundIndex + 1];
          if (foundValue) {
            oldSheetUrl = foundValue;
          }
        }
      }
    }

    if (!oldSheetUrl) {
      console.log(
        `Old sheet URL not found in the _IDS subsheet for search value: ${searchValue}.`
      );
      throw new Error(
        `Old sheet URL not found in the _IDS subsheet for search value: ${searchValue}.`
      );
    }

    var oldSheetID = oldSheetUrl ? shared.extractSheetId(oldSheetUrl) : "";
    if (!oldSheetID) {
      console.log(
        `Old sheet ID could not be extracted from the _IDS subsheet.`
      );
      throw new Error(
        "Old sheet ID could not be extracted from the _IDS subsheet."
      );
    }

    return { oldSheetID: oldSheetID, idMasterID: idMasterID };
  } catch (error) {
    throw error;
  }
}
