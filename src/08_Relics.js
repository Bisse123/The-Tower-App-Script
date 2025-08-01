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
          newRelicHeaderRow = row + 1;
          newRelicNameCol = relicNameIndex + 1;
          newRelicUnlockedCol = relicUnlockedIndex + 1;
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

      var startRow = newRelicHeaderRow + 1;
      
      // Build unlocked status array directly by iterating through new relics data
      var newRelicsUnlocked = [];
      newRelicsData.slice(startRow - 1).forEach(function (row) {
        var relicName = row[newRelicNameCol - 1] || "";
        if (String(relicName).trim() !== "") {
          if (oldRelics.includes(relicName)) {
            newRelicsUnlocked.push([true]);
          } else {
            newRelicsUnlocked.push([false]);
          }
        }
      });
      if (newRelicsUnlocked.length > 0) {
        var endRow = startRow + newRelicsUnlocked.length - 1;
        var unlockedRange = `${sheetName}!${shared.columnToLetter(
          newRelicUnlockedCol
        )}${startRow}:${shared.columnToLetter(newRelicUnlockedCol)}${endRow}`;

        var batchUpdate = [
          {
            range: unlockedRange,
            values: newRelicsUnlocked,
          },
        ];
        return {
          success: true,
          message: `Relics updated successfully: ${newRelicsUnlocked.length} relics processed`,
          batchUpdate: batchUpdate,
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

      var oldRelicHeaderRow = -1;

      // Scan each row to find the header
      for (var row = 0; row < oldRelicsData.length; row++) {
        var rowValues = oldRelicsData[row];
        var relicNameIndex = rowValues.indexOf("Relic Name");
        var relicUnlockedIndex = rowValues.indexOf("Unlocked");
        if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
          oldRelicHeaderRow = row + 1;
          break;
        }
      }

      if (oldRelicHeaderRow === -1) {
        console.log(`Could not find header row in old Relics sheet`);
        return {
          success: false,
          message: `Could not find header row in old Relics sheet`,
        };
      }

      var startRow = oldRelicHeaderRow + 1;

      var oldRelics = [];
      oldRelicsData.slice(startRow - 1).forEach(function(row) {
        var relicName = row[relicNameIndex];
        var isUnlocked = row[relicUnlockedIndex];
        
        if (relicName && (isUnlocked === true || isUnlocked === "TRUE" || isUnlocked === "true")) {
          oldRelics.push(relicName);
        }
      });

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
