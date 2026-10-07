/**
 * Client-callable. Reads a sheet type's data out of the old spreadsheet.
 * @param {string} oldSheetID
 * @param {string} sheetType
 * @param {string} versionDifference Which converter to use, e.g. "v6.4.3".
 * @returns {Object} { success, data, sheetVisibility } or a failure envelope.
 */
function exportData(oldSheetID, sheetType, versionDifference) {
  try {
    if (!sheetType) {
      return errors.reject(
        "exportData",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Sheet type is not defined." },
      );
    }

    if (!oldSheetID) {
      return errors.reject(
        "exportData",
        errors.CODES.INVALID_INPUT,
        "Old sheet ID is missing.",
      );
    }

    var oldSpreadsheet = spreadsheets(
      `${sheetType} oldSpreadsheet`,
      oldSheetID
    );
    if (!oldSpreadsheet) {
      return errors.reject(
        "exportData",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `Old spreadsheet™ not found with ID: ${oldSheetID}` },
      );
    }

    var sheetTypeFunction = sheetVars(sheetType);
    if (!sheetTypeFunction) {
      return errors.reject(
        "exportData",
        errors.CODES.SHEET_STRUCTURE,
        `Sheet™ type function not found for: ${sheetType}`,
      );
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
      return errors.propagate(
        "exportData",
        exportResult,
        `Error exporting data for ${sheetType}: ${
          exportResult && exportResult.message
            ? exportResult.message
            : "Unknown error"
        }`,
      );
    }

    return {
      success: true,
      message: `Export of ${sheetType} data completed successfully.`,
      data: exportResult.data,
      sheetVisibility: sheetVisibility,
    };
  } catch (error) {
    var errorReport = errors.report("exportData", error, {
      note: `Error during export`,
      oldSheetID: oldSheetID,
      sheetType: sheetType,
      versionDifference: versionDifference,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Client-callable. Writes exported data into the new spreadsheet.
 * @param {string} newSheetID
 * @param {string} sheetType
 * @param {Object} data What exportData returned.
 * @param {Object} sheetVisibility Tab visibility to restore.
 * @param {string} idMasterID
 * @returns {Object} { success, message } or a failure envelope.
 */
function importData(newSheetID, sheetType, data, sheetVisibility, idMasterID) {
  try {
    if (!sheetType) {
      return errors.reject(
        "importData",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Sheet type is not defined." },
      );
    }

    if (!newSheetID) {
      return errors.reject(
        "importData",
        errors.CODES.INVALID_INPUT,
        "New sheet ID is missing.",
      );
    }

    if (!data) {
      return errors.reject(
        "importData",
        errors.CODES.INVALID_INPUT,
        "Data to import is missing.",
      );
    }

    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID
    );
    if (!newSpreadsheet) {
      return errors.reject(
        "importData",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `New spreadsheet™ not found with ID: ${newSheetID}` },
      );
    }

    var sheetTypeFunction = sheetVars(sheetType);
    if (!sheetTypeFunction) {
      return errors.reject(
        "importData",
        errors.CODES.SHEET_STRUCTURE,
        `Sheet™ type function not found for: ${sheetType}`,
      );
    }

    if (sheetVisibility && Object.keys(sheetVisibility).length > 0) {
      var hideShowSheetsResult = SheetsAPI.applySheetVisibility(
        newSpreadsheet,
        sheetVisibility
      );
      if (!hideShowSheetsResult || !hideShowSheetsResult.success) {
        return errors.propagate(
          "importData",
          hideShowSheetsResult,
          "The script could not set which tabs are visible in the new sheet.",
        );
      }
    }

    if (idMasterID) {
      data.idMasterID = idMasterID;
    }

    var importResult = sheetTypeFunction.importData(data, newSheetID);
    if (!importResult || !importResult.success) {

      return errors.propagate(
        "importData",
        importResult,
        `${sheetType}: ${
          importResult && importResult.message
            ? importResult.message
            : "Unknown error"
        }`,
        {
          failedUpdates: (importResult && importResult.failedUpdates) || [],
        },
      );
    }

    return {
      success: true,
      message: `Import of ${sheetType} data completed successfully.`,
      updated: true,
    };
  } catch (error) {
    var errorReport = errors.report("importData", error, {
      note: `Error during import`,
      newSheetID: newSheetID,
      sheetType: sheetType,
      data: data,
      sheetVisibility: sheetVisibility,
      idMasterID: idMasterID,
    });
    return errors.fail(errorReport);
  }
}
