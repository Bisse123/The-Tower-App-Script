/**
 * Writes sheet IDs into the IDS Master's IDS tab.
 * @param {string} idMasterID
 * @param {Array<{sheetType: string, sheetID: string}>} idDataEntries
 * @returns {{success: boolean, message: string}} A failure envelope on error.
 */
function updateIdsMaster(idMasterID, idDataEntries) {
  var idsMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
  if (!idsMasterSpreadsheet) {
    return errors.reject(
      "updateIdsMaster",
      errors.CODES.NOT_FOUND,
      "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
      null,
      { note: `IDS Master Spreadsheet™ not found with ID: ${idMasterID}` },
    );
  }

  var idMasterIDSheet = SheetsAPI.getSheetByName(idsMasterSpreadsheet, "IDS");
  if (!idMasterIDSheet) {
    return errors.reject(
      "updateIdsMaster",
      errors.CODES.SHEET_STRUCTURE,
      `IDS sheet™ not found in ID master spreadsheet™`,
    );
  }
  try {
    var batchUpdate = [];
    var idsMasterData = SheetsAPI.batchGetValues(idMasterID, ["IDS"]);
    var idsMasterValues = idsMasterData[0].values;
    var thisSheetInfo = labelUtils.findSheetTypeID(
      idMasterID,
      "IDS",
      "This Sheet ID",
      idsMasterValues,
    );
    if (thisSheetInfo && thisSheetInfo.cell && thisSheetInfo.cell.range) {
      batchUpdate.push({
        range: thisSheetInfo.cell.range,
        values: [[idMasterID]],
      });
    }

    idDataEntries.forEach((entry) => {
      var sheetType = entry.sheetType;
      if (sheetType === "IDS Master") {
        return;
      }
      var idMasterSpreadsheetInfo = labelUtils.findSheetTypeID(
        idMasterID,
        "IDS",
        sheetType,
        idsMasterValues,
      );
      if (
        idMasterSpreadsheetInfo &&
        idMasterSpreadsheetInfo.cell &&
        idMasterSpreadsheetInfo.cell.range
      ) {
        batchUpdate.push({
          range: idMasterSpreadsheetInfo.cell.range,
          values: [[entry.newSheetID]],
        });
      }
    });

    if (batchUpdate.length > 0) {
      try {
        SheetsAPI.batchUpdateValues(idMasterID, [batchUpdate]);
      } catch (error) {
        var errorReport = errors.report("updateIdsMaster", error, {
          note: `Error updating ID Master sheet`,
          idMasterID: idMasterID,
          idDataEntries: idDataEntries,
        });
        return errors.fail(errorReport);
      }
    }

    CacheManager.RemoveSpreadsheet("idMasterSpreadsheet");
    return {
      success: true,
      message: "New IDS Master set successfully",
      gid: idMasterIDSheet.sheetId,
    };
  } catch (error) {
    var errorReport = errors.report("setNewIdsMaster", error, {
      idMasterID: idMasterID,
      idDataEntries: idDataEntries,
    });
    return errors.fail(errorReport, null, {
      gid: idMasterIDSheet.sheetId,
    });
  }
}

/**
 * The grid id of the IDS Master's IDS tab.
 * @param {string} idMasterID
 * @returns {{success: boolean, gid: number}} A failure envelope on error.
 */
function getIdsMasterGid(idMasterID) {
  try {
    var idsMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
    if (!idsMasterSpreadsheet) {
      return errors.reject(
        "getIdsMasterGid",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `IDS Master Spreadsheet™ not found with ID: ${idMasterID}` },
      );
    }

    var idMasterIDSheet = SheetsAPI.getSheetByName(idsMasterSpreadsheet, "IDS");
    if (!idMasterIDSheet) {
      return errors.reject(
        "getIdsMasterGid",
        errors.CODES.SHEET_STRUCTURE,
        `IDS sheet™ not found in ID master spreadsheet™`,
      );
    }

    return {
      success: true,
      message: "IDS Master IDs already set during import",
      gid: idMasterIDSheet.sheetId,
    };
  } catch (error) {
    var errorReport = errors.report("getIdsMasterGid", error, {
      idMasterID: idMasterID,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Writes one sheet's ID into another sheet's IDS tab.
 * @param {string} spreadsheetID Where to write.
 * @param {string} sheetID What to write.
 * @param {string} sheetType
 * @returns {{success: boolean, message: string}} A failure envelope on error.
 */
function updateSheetID(spreadsheetID, sheetID, sheetType) {
  try {
    if (!spreadsheetID) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing spreadsheetID parameter." },
      );
    }
    if (!sheetID) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing sheetID parameter." },
      );
    }
    if (!sheetType) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing sheetType parameter." },
      );
    }
    var spreadsheet = spreadsheets(`${sheetType} spreadsheet`, spreadsheetID);
    if (!spreadsheet) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `Could not access spreadsheet with ID: ${spreadsheetID}` },
      );
    }
    var idSheet = SheetsAPI.getSheetByName(spreadsheet, "IDS");
    if (!idSheet) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.SHEET_STRUCTURE,
        "The script could not find a tab it needs in your sheet.",
        null,
        { note: `IDS sheet not found in spreadsheet with ID: ${spreadsheetID}` },
      );
    }
    var idValues = SheetsAPI.batchGetValues(spreadsheetID, ["IDS"]);
    if (!idValues || idValues.length === 0) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.SHEET_STRUCTURE,
        "The script could not find a tab it needs in your sheet.",
        null,
        { note: `Could not read IDS sheet data from spreadsheet with ID: ${spreadsheetID}` },
      );
    }
    var values = idValues[0].values;
    var sheetTypeInfo = labelUtils.findSheetTypeID(
      spreadsheetID,
      "IDS",
      "IDS Master",
      values,
    );
    if (!sheetTypeInfo || !sheetTypeInfo.cell) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.SHEET_STRUCTURE,
        `Could not find IDS Master entry in IDS sheet.`,
      );
    }
    var currentSheetID = sheetRefs.extractSheetId(sheetTypeInfo.id);
    if (currentSheetID === sheetID) {
      return {
        success: true,
        message: `Sheet ID is already up to date in IDS sheet.`,
      };
    }
    SheetsAPI.batchUpdateValues(spreadsheetID, [
      {
        range: sheetTypeInfo.cell.range,
        values: [[sheetID]],
      },
    ]);
    return {
      success: true,
      message: `Successfully updated IDS Master ID in IDS sheet.`,
    };
  } catch (error) {
    var errorReport = errors.report("updateSheetID", error, {
      note: `Error updating sheet ID`,
      spreadsheetID: spreadsheetID,
      sheetID: sheetID,
      sheetType: sheetType,
    });
    return errors.fail(errorReport);
  }
}
