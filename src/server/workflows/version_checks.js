/**
 * Client-callable. Compares a sheet's version against its template's.
 * @param {string} sheetID
 * @param {string} sheetType
 * @param {boolean} [forceRefresh]
 * @returns {{success: boolean, comparisonResult: string}} A failure envelope on error.
 */
function compareSheetVersions(sheetID, sheetType, forceRefresh = false) {
  var spreadsheet = spreadsheets(`${sheetType} spreadsheet`, sheetID);
  if (!spreadsheet) {
    return errors.reject(
      "compareSheetVersions",
      errors.CODES.NOT_FOUND,
      "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
      null,
      { note: `Spreadsheet™ not found with ID: ${sheetID}` },
    );
  }
  var homePageSheet = SheetsAPI.getSheetByName(spreadsheet, "Home Page");
  if (!homePageSheet) {
    return errors.reject(
      "compareSheetVersions",
      errors.CODES.SHEET_STRUCTURE,
      `Home Page sheet™ not found in ${sheetType} spreadsheet™`,
    );
  }

  var homePageData = SheetsAPI.batchGetValues(
    sheetID,
    ["Home Page"],
    true,
    forceRefresh,
  );
  if (!homePageData || homePageData.length === 0) {
    return errors.reject(
      "compareSheetVersions",
      errors.CODES.SHEET_STRUCTURE,
      `Could not fetch Home Page data from sheet`,
    );
  }
  var homePageValues = homePageData[0].values;
  var versionInfo = versionUtils.findSheetVersion(
    sheetID,
    "Home Page",
    sheetType,
    homePageValues,
  );
  if (
    !versionInfo ||
    !versionInfo.currentVersion ||
    !versionInfo.latestVersion
  ) {
    return errors.reject(
      "compareSheetVersions",
      errors.CODES.SHEET_STRUCTURE,
      `Could not find complete version information in Home Page sheet™`,
    );
  }
  var comparisonResult = versionUtils.compareVersions(
    versionInfo.currentVersion,
    versionInfo.latestVersion,
  );
  if (comparisonResult !== "older") {
    console.log(`Sheet is up to date`);
    return {
      success: true,
      currentVersion: versionInfo.currentVersion,
      latestVersion: versionInfo.latestVersion,
      comparisonResult: comparisonResult,
    };
  }

  return {
    success: true,
    currentVersion: versionInfo.currentVersion,
    latestVersion: versionInfo.latestVersion,
    comparisonResult: comparisonResult,
  };
}

/**
 * Client-callable. Whether a sheet's version has a converter.
 * @param {string} oldSheetID
 * @param {string} sheetType
 * @returns {{success: boolean, versionDifference: string}} A failure envelope on error.
 */
function checkExportCompatibility(oldSheetID, sheetType) {
  try {
    if (!oldSheetID) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing oldSheetID parameter." },
      );
    }
    if (!sheetType) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing sheetType parameter." },
      );
    }

    var oldSpreadsheet = spreadsheets(
      `${sheetType} oldSpreadsheet`,
      oldSheetID,
    );
    if (!oldSpreadsheet) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `Could not access old ${sheetType} spreadsheet with ID: ${oldSheetID}` },
      );
    }

    var oldHomePageSheet = SheetsAPI.getSheetByName(
      oldSpreadsheet,
      "Home Page",
    );
    if (!oldHomePageSheet) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.SHEET_STRUCTURE,
        `Home Page sheet not found in old ${sheetType} spreadsheet`,
      );
    }

    var oldHomePageData = SheetsAPI.batchGetValues(oldSheetID, ["Home Page"]);
    if (!oldHomePageData || oldHomePageData.length === 0) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.SHEET_STRUCTURE,
        `Could not read Home Page data from old ${sheetType} spreadsheet`,
      );
    }

    var oldHomePageValues = oldHomePageData[0].values;

    var oldVersionInfo = versionUtils.findSheetVersion(
      oldSheetID,
      "Home Page",
      sheetType,
      oldHomePageValues,
    );
    if (!oldVersionInfo || !oldVersionInfo.currentVersion) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.INTERNAL,
        `Current Version not found in old ${sheetType} spreadsheet.`,
      );
    }
    var oldVersion = oldVersionInfo.currentVersion;

    var sheetTypeFunction = sheetVars(sheetType);
    if (sheetTypeFunction) {
      var versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
      if (!versionDifference) {
        return errors.reject(
          "checkExportCompatibility",
          errors.CODES.VERSION_OUTDATED,
          `Your ${sheetType} is version ${oldVersion}, which this script cannot read.`,
        );
      }
      return {
        success: true,
        message: `Old ${sheetType} version (${oldVersion}) is compatible for export`,
        oldVersion: oldVersion,
        versionDifference: versionDifference,
      };
    } else {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.SHEET_STRUCTURE,
        `No compatibility function found for ${sheetType}. Cannot verify export compatibility.`,
      );
    }
  } catch (error) {
    var errorReport = errors.report("checkExportCompatibility", error, {
      note: `Error checking export compatibility`,
      oldSheetID: oldSheetID,
      sheetType: sheetType,
    });
    return errors.fail(errorReport);
  }
}
