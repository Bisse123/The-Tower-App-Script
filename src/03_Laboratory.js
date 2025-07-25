const lab = {
  importData: function (versionDifference) {
    try {
      var newSpreadsheet = spreadsheets("newSpreadsheet");
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }
      var newSheetID = newSpreadsheet.spreadsheetId;

      var oldSpreadsheet = spreadsheets("oldSpreadsheet");
      if (!oldSpreadsheet) {
        console.log(`Old spreadsheet not found`);
        return {
          success: false,
          message: "Old spreadsheet not found",
        };
      }
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var getVersionFunction = this.convertVersionFunctions[versionDifference];
      if (!getVersionFunction) {
        console.log(`Unsupported version difference: ${versionDifference}`);
        return {
          success: false,
          message: `Unsupported version difference: ${versionDifference}`,
        };
      }
      var oldDataResult = getVersionFunction();
      if (!oldDataResult || !oldDataResult.success) {
        console.log(`Error processing lab data: ${oldDataResult.message}`);
        return oldDataResult;
      }

      var oldLabLevels = oldDataResult.oldLabLevels || [];

      // Batch get required data for update function only
      var requiredRanges = ["Master Sheet"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;

      var labResult = this.updateLabLevels(
        "Master Sheet",
        oldLabLevels,
        masterSheetData
      );
      if (!labResult || !labResult.success) {
        console.log(`Error updating lab levels: ${labResult.message}`);
        return labResult;
      }

      var batchUpdate = labResult.batchUpdate || [];

      if (batchUpdate.length > 0) {
        var updateResult = SheetsAPI.batchUpdateValues(
          newSheetID,
          batchUpdate
        );
        if (!updateResult) {
          console.log(`Error applying batch updates to new spreadsheet`);
          return {
            success: false,
            message: "Error applying batch updates to new spreadsheet™",
          };
        }
        return {
          success: true,
          message: labResult.message,
        };
      }

      return {
        success: true,
        message: `No updates needed for Laboratory`,
      };
    } catch (error) {
      console.log(`Error in importLabData: ${error.toString()}`);
      return {
        success: false,
        message: "Error importing lab data: " + error.message,
      };
    }
  },

  updateLabLevels: function (sheetName, labUpdates, masterSheetData) {
    try {
      var headerValues = ["Labs"];

      if (!masterSheetData || masterSheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: "Not enough data in Master Sheet",
        };
      }

      var headerRow = masterSheetData[0];
      var lastRow = masterSheetData.length;

      // Find columns where header is in headerValues
      var columnsToCheck = [];
      for (var i = 0; i < headerRow.length; i++) {
        if (headerValues.includes(headerRow[i])) {
          columnsToCheck.push(i + 1);
        }
      }

      if (columnsToCheck.length === 0) {
        console.log(`No Labs columns found in Master Sheet`);
        return {
          success: false,
          message: "No Labs columns found in Master Sheet",
        };
      }

      // Prepare update map from labName to update values
      var updateMap = {};
      labUpdates.forEach(function (update) {
        updateMap[update[0]] = [update[1], update[2]];
      });

      var batchUpdate = [];
      // Iterate each "Labs" column
      columnsToCheck.forEach(function (col) {
        var updates = [];
        // Find labNames in each column (skip header row, ignore last row with sums)
        var numRows = lastRow - 2;

        for (var row = 1; row < numRows + 1; row++) {
          if (row >= masterSheetData.length) break;

          var cellValue = masterSheetData[row][col - 1];
          if (!cellValue || cellValue.trim() === "") break;

          var update = updateMap[cellValue];
          if (update) {
            updates.push([update[0] || 0, update[1] || ""]);
          } else {
            // Keep existing values
            var currentLevel = masterSheetData[row][col] || 0;
            var currentTarget = masterSheetData[row][col + 1] || "";
            updates.push([currentLevel, currentTarget]);
          }
        }
        // Add batch update for this column's Level and Target columns
        if (updates.length > 0) {
          var startCol = shared.columnToLetter(col + 1); // Level column
          var endCol = shared.columnToLetter(col + 2); // Target column
          var range = `${sheetName}!${startCol}2:${endCol}${
            updates.length + 1
          }`;

          batchUpdate.push({
            range: range,
            values: updates,
          });
        }
      });

      // Execute batch updates
      if (batchUpdate.length > 0) {
        // Return batch update data instead of calling API directly
        return {
          success: true,
          message: "Lab levels updated successfully",
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: "No updates needed for lab levels",
      };
    } catch (error) {
      console.log(`Error in updateLabLevels: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating lab levels: ${error.message}`,
      };
    }
  },

  version10: function () {
    try {
      var newSpreadsheet = spreadsheets("newSpreadsheet");
      if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
        console.log(`_IDS sheet not found in new lab spreadsheet`);
        return {
          success: false,
          message: "_IDS sheet™ not found in new lab spreadsheet™",
        };
      }

      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
        console.log(`Master Sheet not found in new lab spreadsheet`);
        return {
          success: false,
          message: "Master Sheet™ not found in new lab spreadsheet™",
        };
      }

      // Get header row to find Labs column
      var headerBatchResult = SheetsAPI.batchGetValues(newSheetID, [
        "_IDS!1:1",
      ]);
      if (
        !headerBatchResult ||
        headerBatchResult.length === 0 ||
        !headerBatchResult[0].values
      ) {
        console.log(`Could not read header row from _IDS sheet`);
        return {
          success: false,
          message: "Could not read header row from _IDS sheet",
        };
      }
      var headerValues = headerBatchResult[0].values;

      var headerRow = headerValues[0];
      var importLabColStart = headerRow.indexOf("Labs") + 1;

      if (importLabColStart === 0) {
        console.log(`Labs column not found in _IDS sheet`);
        return {
          success: false,
          message: "Labs column not found in _IDS sheet",
        };
      }

      // Get lab levels data starting from row 2
      var labLevelsRange =
        "_IDS!" +
        shared.columnToLetter(importLabColStart) +
        "2:" +
        shared.columnToLetter(importLabColStart + 2);

      var labBatchResult = SheetsAPI.batchGetValues(newSheetID, [
        labLevelsRange,
      ]);
      if (
        !labBatchResult ||
        labBatchResult.length === 0 ||
        !labBatchResult[0].values
      ) {
        console.log(`Could not read lab levels data`);
        return {
          success: false,
          message: "Could not read lab levels data",
        };
      }
      var oldLabLevelsValues = labBatchResult[0].values;

      // Filter out empty rows
      var oldLabLevels = oldLabLevelsValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      return {
        success: true,
        message: "Lab levels processed successfully",
        oldLabLevels: oldLabLevels,
      };
    } catch (error) {
      console.log(`Error processing lab data: ${error.toString()}`);
      return {
        success: false,
        message: `Error processing lab data: ${error.message}`,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
    };
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = Object.keys(this.convertVersionFunctions);

    var sortedThresholds = versionCompatibility.slice().sort(function (a, b) {
      return shared.compareVersions(b, a) === "newer" ? 1 : -1;
    });

    for (var i = 0; i < sortedThresholds.length; i++) {
      var threshold = sortedThresholds[i];
      var compareResult = shared.compareVersions(oldVersion, threshold);

      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }

    return null;
  },
};
