const relics = {
  importData: function (versionDifference) {
    try {
      var newSpreadsheet = spreadsheets("newSpreadsheet");
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }
      var newSheetID = newSpreadsheet.spreadsheetId;

      var oldSpreadsheet = spreadsheets("oldSpreadsheet");
      if (!oldSpreadsheet) {
        console.log(`Old spreadsheet not found`);
        return {
          success: false,
          message: "Old spreadsheet™ not found",
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
        console.log(`Error processing relics data: ${oldDataResult.message}`);
        return oldDataResult;
      }

      var oldRelics = oldDataResult.oldRelics || [];

      var requiredRanges = ["Relics"];
      var newRelicsBatchResult = SheetsAPI.batchGetValues(
        newSheetID,
        requiredRanges
      );
      if (!newRelicsBatchResult || newRelicsBatchResult.length === 0) {
        console.log("Error getting relics sheet data");
        return {
          success: false,
          message: "Error getting relics sheet data",
        };
      }

      var newRelicsData =
        newRelicsBatchResult[0] && newRelicsBatchResult[0].values
          ? newRelicsBatchResult[0].values
          : null;

      var relicsResult = this.updateRelics("Relics", oldRelics, newRelicsData);
      if (!relicsResult || !relicsResult.success) {
        console.log(`Error updating relics: ${relicsResult.message}`);
        return relicsResult;
      }

      var batchUpdate = relicsResult.batchUpdate || [];

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
          message: relicsResult.message,
        };
      }

      return {
        success: true,
        message: `No updates needed for Relics`,
      };
    } catch (error) {
      console.log(`Error in importRelicsData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing relics data: ${error.message}`,
      };
    }
  },

  updateRelics: function (sheetName, oldRelics, newRelicsData) {
    try {
      if (!newRelicsData || newRelicsData.length < 3) {
        console.log(`Not enough data in new Relics sheet`);
        return {
          success: false,
          message: `Not enough data in new Relics sheet`,
        };
      }

      var newRelicHeaderRow = null;
      var newRelicNameCol = null;
      var newRelicUnlockedCol = null;

      // Scan each row to find the header
      for (var row = 0; row < newRelicsData.length; row++) {
        var rowValues = newRelicsData[row];
        var relicNameIndex = rowValues.indexOf("Relic Name");
        var relicUnlockedIndex = rowValues.indexOf("Unlocked");
        if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
          newRelicHeaderRow = row + 1; // Convert to 1-based
          newRelicNameCol = relicNameIndex + 1; // Convert to 1-based
          newRelicUnlockedCol = relicUnlockedIndex + 1; // Convert to 1-based
          break;
        }
      }

      if (!newRelicHeaderRow) {
        console.log(`Could not find header row in new Relics sheet`);
        return {
          success: false,
          message: `Could not find header row in new Relics sheet`,
        };
      }

      // Extract unlocked relic names from old data (skip header row)
      var oldRelicsNames = [];
      oldRelics.forEach(function (relic) {
        if (
          relic[relic.length - 1] === true ||
          relic[relic.length - 1] === "TRUE" ||
          relic[relic.length - 1] === "true"
        ) {
          oldRelicsNames.push(relic[0]);
        }
      });

      // Extract new relic names from already fetched data (skip header row)
      var startRow = newRelicHeaderRow + 1;
      var newRelicsNames = newRelicsData
        .slice(startRow - 1) // Skip to data rows (convert to 0-based)
        .map(function (row) {
          return row[newRelicNameCol - 1] || "";
        }) // Extract name column
        .filter(function (name) {
          return String(name).trim() !== "";
        }); // Filter out empty names

      // Create unlocked status array
      var newRelicsUnlocked = [];
      newRelicsNames.forEach(function (relicName) {
        if (oldRelicsNames.includes(relicName)) {
          newRelicsUnlocked.push([true]);
        } else {
          newRelicsUnlocked.push([false]);
        }
      });
      // Update the unlocked column
      if (newRelicsUnlocked.length > 0) {
        var endRow = startRow + newRelicsUnlocked.length - 1;
        var unlockedRange = `${sheetName}!${shared.columnToLetter(
          newRelicUnlockedCol
        )}${startRow}:${shared.columnToLetter(newRelicUnlockedCol)}${endRow}`;

        // Return batch update data instead of calling API directly
        return {
          success: true,
          message: `Relics updated successfully: ${newRelicsUnlocked.length} relics processed`,
          batchUpdate: [
            {
              range: unlockedRange,
              values: newRelicsUnlocked,
            },
          ],
        };
      }
      // console.log(`No updates needed for relics`);
      return {
        success: true,
        message: `No updates needed for relics`,
      };
    } catch (error) {
      console.log("Error in updateRelics: " + error.toString());
      return {
        success: false,
        message: "Error updating relics: " + error.message,
      };
    }
  },

  version10: function () {
    try {
      var newSpreadsheet = spreadsheets("newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;

      var oldSpreadsheet = spreadsheets("oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Check if Relics sheet exists in old spreadsheet
      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "Relics")) {
        console.log("Relics sheet not found in old relic spreadsheet");
        return {
          success: false,
          message: `Relics sheet not found in old relic spreadsheet™`,
        };
      }

      var oldRelicsBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        "Relics",
      ]);
      if (
        !oldRelicsBatchResult ||
        oldRelicsBatchResult.length === 0 ||
        !oldRelicsBatchResult[0].values
      ) {
        console.log(`Could not read data from old Relics sheet`);
        return {
          success: false,
          message: `Could not read data from old Relics sheet`,
        };
      }
      var oldRelicsData = oldRelicsBatchResult[0].values;

      var oldRelicHeaderRow = null;
      var oldRelicNameCol = null;
      var oldRelicUnlockedCol = null;

      // Scan each row to find the header
      for (var row = 0; row < oldRelicsData.length; row++) {
        var rowValues = oldRelicsData[row];
        var relicNameIndex = rowValues.indexOf("Relic Name");
        var relicUnlockedIndex = rowValues.indexOf("Unlocked");
        if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
          oldRelicHeaderRow = row + 1; // Convert to 1-based
          oldRelicNameCol = relicNameIndex + 1; // Convert to 1-based
          oldRelicUnlockedCol = relicUnlockedIndex + 1; // Convert to 1-based
          break;
        }
      }

      if (!oldRelicHeaderRow) {
        console.log(`Could not find header row in old Relics sheet`);
        return {
          success: false,
          message: `Could not find header row in old Relics sheet`,
        };
      }

      var startRow = oldRelicHeaderRow + 1;
      var oldRelicsRange =
        "Relics!" +
        shared.columnToLetter(oldRelicNameCol) +
        startRow +
        ":" +
        shared.columnToLetter(oldRelicUnlockedCol);

      var oldRelicsBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        oldRelicsRange,
      ]);
      if (
        !oldRelicsBatchResult ||
        oldRelicsBatchResult.length === 0 ||
        !oldRelicsBatchResult[0].values
      ) {
        console.log(`Could not read old relics data`);
        return {
          success: false,
          message: `Could not read old relics data`,
        };
      }
      var oldRelicsValues = oldRelicsBatchResult[0].values;

      // Filter out empty rows
      var oldRelics = oldRelicsValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null && cell !== undefined && String(cell).trim() !== ""
        )
      );

      // Check if Relics sheet exists in new spreadsheet
      if (!SheetsAPI.getSheetByName(newSpreadsheet, "Relics")) {
        console.log("Relics sheet not found in new relic spreadsheet");
        return {
          success: false,
          message: `Relics sheet not found in new relic spreadsheet™`,
        };
      }

      return {
        success: true,
        oldRelics: oldRelics,
      };
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
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
