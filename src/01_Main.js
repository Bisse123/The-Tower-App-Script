const sheetVars = (sheetType) => {
  var sheetTypeFunctions = {
    Laboratory: lab,
    Workshop: workshop,
    "Ultimate Weapon": ultimate,
    "Themes, Songs & Relics": themesAndRelics,
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

const spreadsheets = (spreadsheetTypeName, sheetID) => {
  if (!spreadsheetTypeName) {
    console.log(`No spreadsheet type name provided.`);
    return null;
  }

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

const ADDON_CONSENT_READY_SIGNAL_KEY = "ADDON_CONSENT_READY_SIGNAL";

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  var targetPage = (params.page || "").toString().toLowerCase();

  if (targetPage === "savefile" || targetPage === "save-file" || params.saveFile === "true") {
    var sheetType = params.sheetType || "";
    if (sheetType === "IDS Collection - all IDS-Sheets on one file") {
      sheetType = "IDS Collection";
    }
    var saveFileTemplate = HtmlService.createTemplateFromFile("20_SavedFileApp");
    saveFileTemplate.API_KEY =
      PropertiesService.getScriptProperties().getProperty("API_KEY");
    saveFileTemplate.APP_ID =
      PropertiesService.getScriptProperties().getProperty("APP_ID");
    saveFileTemplate.viewType = "webapp";
    saveFileTemplate.idMasterID = params.idMasterID
      ? shared.extractSheetId(params.idMasterID) || ""
      : "";
    saveFileTemplate.sheetType = sheetType;
    return saveFileTemplate
      .evaluate()
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setTitle("Import Data From Game");
  }

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
    var effectivePathsID = RegExp("^[a-zA-Z0-9-_]{44}$").test(params.effectivePathsID || "")
       ? params.effectivePathsID
       : "";
    getStartedTemplate.effectivePathsID = effectivePathsID;
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

  if (sheetType === "IDS Master") {
    if (!oldSheetID && idMasterID) {
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
            .addItem("Update Sheet", "showUpdateDialog")
            .addItem("Import Data From Game (playerInfo.dat)", "openSaveFileDialog")
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
    template.effectivePathsID = "";
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

function openSaveFileDialog() {
  var template = HtmlService.createTemplateFromFile("20_SavedFileApp");
  template.API_KEY =
    PropertiesService.getScriptProperties().getProperty("API_KEY");
  template.APP_ID =
    PropertiesService.getScriptProperties().getProperty("APP_ID");
  template.viewType = "sidebar";

  template.idMasterID = "";
  template.sheetType = "";

  var html = template
    .evaluate()
    .setWidth(1280)
    .setHeight(720)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");

  SpreadsheetApp.getUi().showModalDialog(html, "Load Data From Save File");
}

function findIdMasterIdInIdsTab(idsSheet) {
  var values = idsSheet.getDataRange().getValues();
  if (!values || values.length === 0) {
    return "";
  }

  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < values[row].length; col++) {
      var cell = values[row][col];
      if (
        typeof cell === "string" &&
        cell.indexOf("IDS Master") !== -1 &&
        cell.indexOf("script") === -1
      ) {
        var idMasterURL = values[row][col + 2];
        return idMasterURL ? shared.extractSheetId(idMasterURL) || "" : "";
      }
    }
  }

  return "";
}

function showAddonConsentDialog(authorizationUrl) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(ADDON_CONSENT_READY_SIGNAL_KEY);

  var template = HtmlService.createTemplateFromFile("29_addon_consent_dialog");
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

function getGetStartedParameters() {
  try {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSpreadsheet) {
      throw new Error("Active spreadsheet not found.");
    }

    var hasEffectivePathsSheet = Boolean(
      activeSpreadsheet.getSheetByName("eHP") ||
        activeSpreadsheet.getSheetByName("eDamage") ||
        activeSpreadsheet.getSheetByName("eEcon")
    );

    if (!hasEffectivePathsSheet) {
      return {
        success: false,
        sheetId: "",
        message: "Effective Paths sheets not found in active spreadsheet.",
      };
    }

    return {
      success: true,
      sheetId: activeSpreadsheet.getId(),
    };
  } catch (error) {
    console.log(`Error in getGetStartedParameters: ${error.message || error}`);
    return {
      success: false,
      sheetId: "",
      message: error.message || error.toString(),
    };
  }
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
      return {
        success: true,
        newSheetID: "",
        oldSheetID: oldSheetID,
        idMasterID: "",
        sheetType: sheetType,
        accessRequired: false,
      };
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

    var idsSheet = oldSpreadsheet.getSheetByName("IDS");
    if (!idsSheet) {
      throw new Error("IDS sheet not found in the active spreadsheet.");
    }

    var idMasterID = findIdMasterIdInIdsTab(idsSheet);

    return {
      success: true,
      newSheetID: "",
      oldSheetID: oldSheetID,
      idMasterID: idMasterID,
      sheetType: sheetType,
      accessRequired: true,
    };
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

function getSaveFileParameters() {
  try {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSpreadsheet) {
      return {
        success: false,
        idMasterID: "",
        sheetType: "",
        message: "Active spreadsheet not found.",
      };
    }

    var sheetType = "";
    if (
      activeSpreadsheet.getSheetByName("eHP") ||
      activeSpreadsheet.getSheetByName("eDamage") ||
      activeSpreadsheet.getSheetByName("eEcon")
    ) {
      sheetType = "Effective Paths";
    } else {
      var homePageSheet = activeSpreadsheet.getSheetByName("Home Page");
      if (homePageSheet) {
        sheetType = homePageSheet.getRange("B2").getValue();
      }
    }

    if (sheetType === "IDS Master") {
      return {
        success: true,
        idMasterID: activeSpreadsheet.getId(),
        sheetType: sheetType,
      };
    }

    if (typeof sheetType === "string" && sheetType.indexOf("IDS Collection") !== -1) {
      return {
        success: true,
        idMasterID: activeSpreadsheet.getId(),
        sheetType: "IDS Collection",
      };
    }

    var idsSheet = activeSpreadsheet.getSheetByName("IDS");
    if (!idsSheet) {
      return {
        success: false,
        idMasterID: "",
        sheetType: sheetType,
        message: "IDS tab not found in the active spreadsheet.",
      };
    }

    var idMasterID = findIdMasterIdInIdsTab(idsSheet);
    if (!idMasterID) {
      return {
        success: false,
        idMasterID: "",
        sheetType: sheetType,
        message: "Could not find the IDS Master's ID in the IDS tab.",
      };
    }

    return {
      success: true,
      idMasterID: idMasterID,
      sheetType: sheetType,
    };
  } catch (error) {
    console.log(`Error in getSaveFileParameters: ${error.message || error}`);
    return {
      success: false,
      idMasterID: "",
      sheetType: "",
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

    var sheetVisibility = {};
    if (oldSpreadsheet && oldSpreadsheet.sheets) {
      for (var i = 0; i < oldSpreadsheet.sheets.length; i++) {
        var sheet = oldSpreadsheet.sheets[i];
        sheetVisibility[sheet.properties.title] =
          sheet.properties.hidden || false;
      }
    }

    var exportResult = sheetTypeFunction.exportData(versionDifference, oldSheetID);
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

    if (idMasterID) {
      data.idMasterID = idMasterID;
    }

    var importResult = sheetTypeFunction.importData(data, newSheetID);
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
