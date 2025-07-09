const lab = {
  importData: function importData(versionDifference) {
    function importLabData(versionDifference) {
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

        var getVersionFunction = convertVersionFunctions[versionDifference];
        if (!getVersionFunction) {
          console.log(`Unsupported version difference: ${versionDifference}`);
          return {
            success: false,
            message: `Unsupported version difference: ${versionDifference}`,
          };
        }
        var result = getVersionFunction(newSheetID, oldSheetID);
        if (!result || !result.success) {
          console.log(`Error processing lab data: ${result.message}`);
          return result;
        }

        var oldLabLevels = result.oldLabLevels || [];
        return updateLabLevels(newSheetID, oldLabLevels);
      } catch (error) {
        console.log(`Error in importLabData: ${error.toString()}`);
        return {
          success: false,
          message: "Error importing lab data: " + error.message,
        };
      }
    }

    function version10(newSheetID, oldSheetID) {
      try {
        var newSpreadsheet = SpreadsheetApp.openById(newSheetID);
        if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
          console.log(`_IDS sheet not found in new lab spreadsheet`);
          return {
            success: false,
            message: "_IDS sheet not found in new lab spreadsheet",
          };
        }

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
          console.log(`Master Sheet not found in new lab spreadsheet`);
          return {
            success: false,
            message: "Master Sheet not found in new lab spreadsheet",
          };
        }

        // Get header row to find Labs column
        var headerValues = SheetsAPI.getValues(newSheetID, "_IDS!1:1");

        if (!headerValues || headerValues.length === 0) {
          console.log(`Could not read header row from _IDS sheet`);
          return {
            success: false,
            message: "Could not read header row from _IDS sheet",
          };
        }

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

        var oldLabLevelsValues = SheetsAPI.getValues(
          newSheetID,
          labLevelsRange
        );
        if (!oldLabLevelsValues) {
          console.log(`Could not read lab levels data`);
          return {
            success: false,
            message: "Could not read lab levels data",
          };
        }

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
    }

    function updateLabLevels(newSheetID, labUpdates) {
      try {
        var headerValues = ["Labs"];

        // Get all data from Master Sheet to determine range
        var allData = SheetsAPI.getValues(newSheetID, "Master Sheet");
        if (!allData || allData.length < 2) {
          console.log(`Not enough data in Master Sheet`);
          return {
            success: false,
            message: "Not enough data in Master Sheet",
          };
        }

        var headerRow = allData[0];
        var lastRow = allData.length;

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
            if (row >= allData.length) break;

            var cellValue = allData[row][col - 1];
            if (cellValue === "") break;

            var update = updateMap[cellValue];
            // If labName is found then push update regardless of the imported levels
            if (update) {
              updates.push([update[0], update[1]]);
            } else {
              // Keep existing values
              var currentLevel = allData[row][col] || "";
              var currentTarget = allData[row][col + 1] || "";
              updates.push([currentLevel, currentTarget]);
            }
          }
          // Add batch update for this column's Level and Target columns
          if (updates.length > 0) {
            var startCol = shared.columnToLetter(col + 1); // Level column
            var endCol = shared.columnToLetter(col + 2); // Target column
            var range =
              "Master Sheet!" + startCol + "2:" + endCol + (updates.length + 1);

            batchUpdate.push({
              range: range,
              values: updates,
            });
          }
        });

        // Execute batch updates
        if (batchUpdate.length > 0) {
          SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
          // console.log("Lab levels updated successfully");
          return {
            success: true,
            message: "Lab levels updated successfully",
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
    }

    var convertVersionFunctions = {
      "v1.0": version10,
    };

    return importLabData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = {
      "v1.0": true,
    };
    
    var highestThreshold = null;
    var compatibleThreshold = null;
    
    for (var threshold in versionCompatibility) {
      var compareResult = shared.compareVersions(oldVersion, threshold);
      
      if (compareResult === "older" || compareResult === "same") {
        if (versionCompatibility[threshold]) {
          compatibleThreshold = threshold;
        }
        break;
      }
      
      if (!highestThreshold || shared.compareVersions(highestThreshold, threshold) === "older") {
        highestThreshold = threshold;
      }
    }
    
    if (!compatibleThreshold && highestThreshold) {
      var compareWithHighest = shared.compareVersions(highestThreshold, oldVersion);
      if (compareWithHighest === "older") {
        compatibleThreshold = highestThreshold;
      }
    }
    
    return compatibleThreshold;
  },
};
