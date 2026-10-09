/**
 * Client-callable. An OAuth token for the Picker.
 * @returns {string}
 */
function getOAuthToken() {
  try {
    const token = ScriptApp.getOAuthToken();
    return {
      success: true,
      token: token,
      authorizationUrl: "",
      message: "Token retrieved successfully",
    };
  } catch (error) {
    var errorReport = errors.report("getOAuthToken", error, { note: `Error getting OAuth token` });

    return errors.fail(errorReport, null, {
      token: null,
      authorizationUrl: getScopeAuthorizationUrl(),
    });
  }
}

/**
 * Client-callable. The URL where the user grants missing scopes.
 * @returns {string} "" when it cannot be built.
 */
function getScopeAuthorizationUrl() {
  try {
    return (
      ScriptApp.getAuthorizationInfo(
        ScriptApp.AuthMode.FULL,
      ).getAuthorizationUrl() || ""
    );
  } catch (authInfoError) {
    errors.report("getScopeAuthorizationUrl", authInfoError, { note: `Error getting authorization URL` }, errors.CODES.RECOVERED);
    return "";
  }
}

/**
 * Client-callable. Whether every required scope is granted.
 * @returns {boolean}
 */
function checkScopePermissions() {
  try {
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
    return true;
  } catch (error) {
    errors.report(
      "checkScopePermissions",
      error,
      { note: `Scope permission check failed` }, errors.CODES.ACCESS_DENIED);
    return false;
  }
}

/**
 * Client-callable. Whether the script can reach a sheet, and how.
 * @param {string} sheetID
 * @returns {{success: boolean, accessible: boolean, owned: boolean, canEdit: boolean, sheetID: string}}
 */
function checkSheetAccess(sheetID) {
  try {
    if (!sheetID) {
      return errors.reject(
        "checkSheetAccess",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        {
          accessible: false,
          owned: false,
          canEdit: false,
        },
        { note: "Missing sheetID parameter" },
      );
    }

    try {
      const file = CacheManager.getFile(sheetID);
      const parentFolderID =
        file.parents && file.parents.length > 0 ? file.parents[0] : null;
      const owners = file.owners || [];
      const isOwner = owners.some((owner) => owner.me === true);
      const canEdit = !!(file.capabilities && file.capabilities.canEdit);

      return {
        success: true,
        message: `Sheet access verified`,
        accessible: true,
        owned: isOwner,
        canEdit: canEdit,
        sheetID: sheetID,
        name: file.name,
        parentFolderID: parentFolderID,
      };
    } catch (error) {
      errors.report(
        "checkSheetAccess",
        error,
        { note: `Sheet access denied for ${sheetID}`, sheetID: sheetID }, errors.CODES.ACCESS_DENIED);

      return {
        success: true,
        message: `Sheet access denied`,
        accessible: false,
        owned: false,
        canEdit: false,
        sheetID: sheetID,
      };
    }
  } catch (error) {
    var errorReport = errors.report("checkSheetAccess", error, {
      note: `Error checking sheet access`,
      sheetID: sheetID,
    });
    return errors.fail(errorReport, null, {
      accessible: false,
      owned: false,
      canEdit: false,
    });
  }
}

/**
 * Client-callable. Whether the script can reach a template.
 * @param {string} templateID
 * @returns {{success: boolean, accessible: boolean, templateID: string}}
 */
function checkTemplateAccess(templateID) {
  try {
    if (!templateID) {
      return errors.reject(
        "checkTemplateAccess",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        {
          accessible: false,
        },
        { note: "Missing templateID parameter" },
      );
    }

    try {
      var file = CacheManager.getFile(templateID);

      return {
        success: true,
        message: `Template access verified`,
        accessible: true,
        templateID: templateID,
        name: file.name,
      };
    } catch (error) {
      errors.report(
        "checkTemplateAccess",
        error,
        { note: `Template access denied for ${templateID}`, templateID: templateID }, errors.CODES.ACCESS_DENIED);

      return {
        success: true,
        message: `Template access denied`,
        accessible: false,
        templateID: templateID,
      };
    }
  } catch (error) {
    var errorReport = errors.report("checkTemplateAccess", error, {
      note: `Error checking template access`,
      templateID: templateID,
    });
    return errors.fail(errorReport, null, {
      accessible: false,
    });
  }
}

/**
 * Client-callable. Template access for one sheet type in an IDS Master.
 * @param {string} idMasterID
 * @param {string} sheetType
 * @returns {Object} What processTemplateAccess returned.
 */
function checkFileTemplateAccess(idMasterID, sheetType) {

  var result = null;
  for (var attempt = 1; attempt <= 2; attempt++) {
    if (attempt > 1) Utilities.sleep(700);

    var idsMasterData = fetchIdsMasterData(idMasterID, attempt > 1);
    if (!idsMasterData.success) {
      return idsMasterData;
    }

    result = processTemplateAccess(idsMasterData, sheetType, "all");
    if (!result || !result.versionLoading) break;
  }

  return result;
}

/**
 * Client-callable. Whether the new sheet points back at its IDS Master.
 * @param {string} newSheetID
 * @param {string} sheetType
 * @returns {{success: boolean}} A failure envelope on error.
 */
function checkNewSheetReference(newSheetID, sheetType) {
  try {
    if (!newSheetID) {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing newSheetID parameter." },
      );
    }

    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID,
    );
    if (!newSpreadsheet) {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `New spreadsheet™ not found with ID: ${newSheetID}` },
      );
    }

    var newIdSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
    if (!newIdSheet) {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.SHEET_STRUCTURE,
        `IDS sheet™ not found in the new spreadsheet™.`,
      );
    }

    var newSpreadsheetInfo = labelUtils.findSheetTypeID(newSheetID, "IDS");
    if (!newSpreadsheetInfo || !newSpreadsheetInfo.id) {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.SHEET_STRUCTURE,
        `Could not find sheet type ID for ${newSheetID}`,
      );
    }
    var accessStatus = newSpreadsheetInfo.accessStatus;
    if (!accessStatus || accessStatus.value !== "✅") {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.ACCESS_DENIED,
        `New sheet have not been granted access to IDS Master.`,
      );
    }

    return {
      success: true,
      message: `New sheet reference is valid.`,
    };
  } catch (error) {
    var errorReport = errors.report("checkNewSheetReference", error, {
      note: `Error checking new sheet reference`,
      newSheetID: newSheetID,
      sheetType: sheetType,
    });
    return errors.fail(errorReport);
  }
}
