const lab = {
  convertVersionFunctions: {},

  importData: function importData(sheetType, newLabSpreadsheetId) {
    function importLabData(sheetType, newLabSpreadsheetId) {
      try {
        var idType = sheetType + " ID";

        // Check if required sheets exist
        if (!SheetsAPI.hasSheet(newLabSpreadsheetId, "IDS")) {
          console.log("IDS sheet not found in new lab spreadsheet");
          return {
            success: false,
            message: "IDS sheet not found in new lab spreadsheet",
          }
        }
        if (!SheetsAPI.hasSheet(newLabSpreadsheetId, "EXPORT")) {
          console.log("EXPORT sheet not found in new lab spreadsheet");
          return {
            success: false,
            message: "EXPORT sheet not found in new lab spreadsheet",
          }
        }

        // Get version from EXPORT sheet
        var newLabVersion = SheetsAPI.getValue(
          newLabSpreadsheetId,
          "EXPORT!A1"
        );

        // Get ID Master spreadsheet info
        var idMasterSpreadsheetInfo = shared.findSheetTypeID(
          newLabSpreadsheetId,
          "IDS"
        );
        if (!idMasterSpreadsheetInfo || !idMasterSpreadsheetInfo.id) {
          console.log("Could not find ID Master spreadsheet info");
          return {
            success: false,
            message: "Could not find ID Master spreadsheet info",
          };
        }
        
        var idMasterSpreadsheetId = shared.extractSheetId(
          idMasterSpreadsheetInfo.id
        );
        if (!idMasterSpreadsheetId) {
          console.log("Could not find ID Master spreadsheet");
          return {
            success: false,
            message: "Could not find ID Master spreadsheet",
          };
        }

        var oldLabSpreadsheetInfo = shared.findSheetTypeID(
          idMasterSpreadsheetId,
          "IDS",
          idType
        );
        if (!oldLabSpreadsheetInfo || !oldLabSpreadsheetInfo.id) {
          console.log("Could not find old lab spreadsheet info");
          return {
            success: false,
            message: "Could not find old lab spreadsheet info",
          };
        }

        var oldLabSpreadsheetId = shared.extractSheetId(
          oldLabSpreadsheetInfo.id
        );
        if (!oldLabSpreadsheetId) {
          console.log("Could not find old lab spreadsheet");
          return {
            success: false,
            message: "Could not find old lab spreadsheet",
          };
        }

        var oldLabVersion = SheetsAPI.getValue(
          oldLabSpreadsheetId,
          "EXPORT!A1"
        );
        var versionCheck = shared.compareVersions(oldLabVersion, newLabVersion);

        if (versionCheck === 0) {
          console.log("Same Version - proceeding with lab data import");

          // Check if _IDS sheet exists
          if (!SheetsAPI.hasSheet(newLabSpreadsheetId, "_IDS")) {
            console.log("_IDS sheet not found in new lab spreadsheet");
            return {
              success: false,
              message: "_IDS sheet not found in new lab spreadsheet",
            };
          }

          // Get header row to find Labs column
          var headerValues = SheetsAPI.getValues(
            newLabSpreadsheetId,
            "_IDS!1:1"
          );
          if (!headerValues || headerValues.length === 0) {
            console.log("Could not read header row from _IDS sheet");
            return {
              success: false,
              message: "Could not read header row from _IDS sheet",
            };
          }

          var headerRow = headerValues[0];
          var importLabColStart = headerRow.indexOf("Labs") + 1;

          if (importLabColStart === 0) {
            console.log("Labs column not found in _IDS sheet");
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
            newLabSpreadsheetId,
            labLevelsRange
          );
          if (!oldLabLevelsValues) {
            console.log("Could not read lab levels data");
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
                String(cell || '').trim() !== ""
            )
          );

          // Check if Master Sheet exists
          if (!SheetsAPI.hasSheet(newLabSpreadsheetId, "Master Sheet")) {
            console.log("Master Sheet not found in new lab spreadsheet");
            return {
              success: false,
              message: "Master Sheet not found in new lab spreadsheet",
            };
          }

          return updateLabLevels(newLabSpreadsheetId, oldLabLevels);
        } else {
          console.log("Version mismatch - skipping lab data import");
          return {
            success: false,
            message: "Version mismatch - skipping lab data import",
          };
        }
      } catch (error) {
        console.log("Error in importLabData: " + error.toString());
        return {
          success: false,
          message: "Error importing lab data: " + error.message,
        };
      }
    }

    function updateLabLevels(spreadsheetId, labUpdates) {
      try {
        var headerValues = ["Labs"];

        // Get all data from Master Sheet to determine range
        var allData = SheetsAPI.getValues(spreadsheetId, "Master Sheet");
        if (!allData || allData.length < 2) {
          console.log("Not enough data in Master Sheet");
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
          console.log("No Labs columns found in Master Sheet");
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

        // Prepare batch updates
        var batchUpdates = [];

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

            batchUpdates.push({
              range: range,
              values: updates,
            });
          }
        });

        // Execute batch updates
        if (batchUpdates.length > 0) {
          batchUpdates.forEach(function (update) {
            SheetsAPI.setValues(spreadsheetId, update.range, update.values);
          });
          console.log("Lab levels updated successfully");
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
        console.log("Error in updateLabLevels: " + error.toString());
        return {
          success: false,
          message: "Error updating lab levels: " + error.message,
        };
      }
    }
    
    return importLabData(sheetType, newLabSpreadsheetId);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
