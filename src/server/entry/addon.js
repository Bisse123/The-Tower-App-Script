const ADDON_CONSENT_READY_SIGNAL_KEY = "ADDON_CONSENT_READY_SIGNAL";

/**
 * Simple trigger: builds the menu when a spreadsheet opens.
 * @param {Object} e
 * @returns {void}
 */
function onOpen(e) {
  createMenu();
}

/**
 * Simple trigger: builds the menu when the add-on is installed.
 * @param {Object} e
 * @returns {void}
 */
function onInstall(e) {
  createMenu();
}

/**
 * Adds the Import Data menu. Silently skipped where there is no UI.
 * @returns {void}
 */
function createMenu() {
    try {
        var ui = SpreadsheetApp.getUi();
        ui.createMenu("Import Data")
            .addItem("Get Started", "showGetStartedDialog")
            .addItem("Update Sheet", "showUpdateDialog")
            .addItem("Import Data From Game (playerInfo.dat)", "openSaveFileDialog")
            .addToUi();
    } catch (error) {
        console.log(`createMenu skipped: ${errors.text(error)}`);
    }
}

/**
 * Opens the Get Started modal dialog.
 * @returns {void}
 */
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
    errors.report("showGetStartedDialog", error, null, errors.CODES.RECOVERED);
    SpreadsheetApp.getUi().alert("Error: " + error.message);
  }
}

/**
 * Opens the update sidebar.
 * @returns {void}
 */
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
    errors.report("showUpdateDialog", error, null, errors.CODES.RECOVERED);
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

/**
 * Opens the save-file import modal dialog.
 * @returns {void}
 */
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

/**
 * Finds the IDS Master sheet ID in an IDS tab.
 * @param {Sheet} idsSheet
 * @returns {string|null} The extracted ID, or null when not found.
 */
function findIdMasterIdInIdsTab(idsSheet) {
  var values = idsSheet.getDataRange().getValues();
  if (!values || values.length === 0) {
    return null;
  }

  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < values[row].length; col++) {
      if (labelUtils.isSheetTypeCell(values[row][col], "IDS Master")) {
        var idMasterURL = values[row][col + 2];
        return idMasterURL ? sheetRefs.extractSheetId(idMasterURL) || null : null;
      }
    }
  }

  return null;
}

/**
 * Opens the additional-permissions dialog.
 * @param {string} [authorizationUrl]
 * @returns {void}
 */
function showAddonConsentDialog(authorizationUrl) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(ADDON_CONSENT_READY_SIGNAL_KEY);

  var template = HtmlService.createTemplateFromFile("client/pages/consent_dialog");
  template.authorizationUrl = authorizationUrl || "";

  var html = template
    .evaluate()
    .setWidth(560)
    .setHeight(280)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");

  SpreadsheetApp.getUi().showModalDialog(html, "Additional Permissions Required");
}

/**
 * Client-callable. Records that the user finished the consent dialog.
 * @returns {boolean} Always true.
 */
function markAddonConsentReadySignal() {
  PropertiesService.getUserProperties().setProperty(
    ADDON_CONSENT_READY_SIGNAL_KEY,
    String(Date.now())
  );
  return true;
}

/**
 * Reads and clears the consent-ready signal.
 * @returns {boolean} Whether a signal was waiting.
 */
function consumeAddonConsentReadySignal() {
  var userProperties = PropertiesService.getUserProperties();
  var signal = userProperties.getProperty(ADDON_CONSENT_READY_SIGNAL_KEY);

  if (!signal) {
    return false;
  }

  userProperties.deleteProperty(ADDON_CONSENT_READY_SIGNAL_KEY);
  return true;
}

/**
 * Client-callable. The active spreadsheet's ID, if it holds Effective Paths.
 * @returns {{success: boolean, sheetId: string}} A failure envelope on error.
 */
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
      return errors.reject(
        "getGetStartedParameters",
        errors.CODES.SHEET_STRUCTURE,
        "Effective Paths sheets not found in active spreadsheet.",
        {
          sheetId: "",
        },
      );
    }

    return {
      success: true,
      sheetId: activeSpreadsheet.getId(),
    };
  } catch (error) {
    var errorReport = errors.report("getGetStartedParameters", error);
    return errors.fail(errorReport, null, {
      sheetId: "",
    });
  }
}

/**
 * Client-callable. The old sheet, IDS Master and sheet type for the update
 * workflow, read from the active spreadsheet.
 * @returns {Object} { success, oldSheetID, idMasterID, sheetType, ... }
 */
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
    var errorReport = errors.report("getUpdateDialogParameters", error);
    return errors.fail(errorReport, null, {
      newSheetID: "",
      oldSheetID: "",
      idMasterID: "",
      sheetType: "",
      accessRequired: true,
    });
  }
}

/**
 * Client-callable. The IDS Master ID and sheet type for the save-file
 * workflow, read from the active spreadsheet.
 * @returns {{success: boolean, idMasterID: string, sheetType: string}}
 */
function getSaveFileParameters() {
  try {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSpreadsheet) {
      return errors.reject(
        "getSaveFileParameters",
        errors.CODES.INTERNAL,
        "Active spreadsheet not found.",
        {
          idMasterID: "",
          sheetType: "",
        },
      );
    }

    var activeSheetType = "";
    if (
      activeSpreadsheet.getSheetByName("eHP") ||
      activeSpreadsheet.getSheetByName("eDamage") ||
      activeSpreadsheet.getSheetByName("eEcon")
    ) {
      activeSheetType = "Effective Paths";
    } else {
      var homePageSheet = activeSpreadsheet.getSheetByName("Home Page");
      if (homePageSheet) {
        activeSheetType = String(
          homePageSheet.getRange("B2").getValue() || ""
        ).trim();
      }
    }

    if (activeSheetType === "IDS Master") {
      return {
        success: true,
        idMasterID: activeSpreadsheet.getId(),
        sheetType: "IDS Master",
      };
    }

    if (activeSheetType.indexOf("IDS Collection") !== -1) {
      return {
        success: true,
        idMasterID: activeSpreadsheet.getId(),
        sheetType: "IDS Collection",
      };
    }

    var idsSheet = activeSpreadsheet.getSheetByName("IDS");
    if (!idsSheet) {
      return errors.reject(
        "getSaveFileParameters",
        errors.CODES.SHEET_STRUCTURE,
        "IDS tab not found in the active spreadsheet.",
        {
          idMasterID: "",
          sheetType: "",
        },
      );
    }

    var idMasterID = findIdMasterIdInIdsTab(idsSheet);
    if (!idMasterID) {
      return errors.reject(
        "getSaveFileParameters",
        errors.CODES.SHEET_STRUCTURE,
        "Could not find the IDS Master's ID in the IDS tab.",
        {
          idMasterID: "",
          sheetType: "",
        },
      );
    }

    return {
      success: true,
      idMasterID: idMasterID,
      sheetType: "",
    };
  } catch (error) {
    var errorReport = errors.report("getSaveFileParameters", error);
    return errors.fail(errorReport, null, {
      idMasterID: "",
      sheetType: "",
    });
  }
}
