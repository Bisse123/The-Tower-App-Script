const guardians = {
  exportData: function (versionDifference) {
    try {
      var getVersionFunction = this.convertVersionFunctions[versionDifference];
      if (!getVersionFunction) {
        console.log(`Unsupported version: ${versionDifference}`);
        return {
          success: false,
          message: `Unsupported version: ${versionDifference}`,
        };
      }
      
      var oldDataResult = getVersionFunction();
      if (!oldDataResult || !oldDataResult.success) {
        console.log(`${oldDataResult.message}`);
        return oldDataResult;
      }

      return {
        success: true,
        message: "Guardians export completed successfully",
        data: {
          oldGuardians: oldDataResult.oldGuardians || {}
        }
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting guardians data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      // Use sheet type-based naming for parallel execution support
      var newSpreadsheet = spreadsheets("Guardians newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      var oldGuardians = data.oldGuardians || {};

      // Batch fetch required sheet data
      var requiredRanges = ["Master Sheet", "IDS"];
      var batchResult = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (
        !batchResult ||
        batchResult.length === 0 ||
        !batchResult[0].values
      ) {
        console.log("Error getting guardians sheet data");
        return {
          success: false,
          message: "Error getting guardians sheet data",
        };
      }

      var masterSheetData = batchResult[0].values;
      var idsData = batchResult[1].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(newSheetID, "IDS", "IDS Master's", idsData);
      if (!newSheetInfo || !newSheetInfo.importStatus || !newSheetInfo.importStatus.range) {
        console.log(`Could not find import status range in IDS sheet`);
        return {
          success: false,
          message: "Could not find import status range in IDS sheet",
        };
      }

      var guardiansResult = this.updateGuardianLevels(
        "Master Sheet",
        oldGuardians,
        masterSheetData
      );
      if (!guardiansResult || !guardiansResult.success) {
        console.log(`Error updating guardians: ${guardiansResult.message}`);
        return guardiansResult;
      }

      var batchUpdate = guardiansResult.batchUpdate || [];

      // Add import status update to batch
      batchUpdate.push({
        range: newSheetInfo.importStatus.range,
        values: [["✅"]],
      });

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
        message: `Guardians import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing guardians data: ${error.message}`,
      };
    }
  },

  updateGuardianLevels: function (
    sheetName,
    oldGuardians,
    masterSheetData
  ) {
    var targetGuardians = ["Attack", "Ally", "Steal", "Fetch"];
    if (!masterSheetData || masterSheetData.length < 2) {
      return {
        success: false,
        message: "Not enough data in Master Sheet™",
      };
    }

    var headerRow = masterSheetData[0];
    var guardianCol = headerRow.indexOf("Guardians") + 1;

    if (guardianCol === 0) {
      console.log(`Guardian Weapon column not found`);
      return {
        success: false,
        message: `Guardian Weapon column not found`,
      };
    }

    var startCol = guardianCol + 1;
    var endCol = guardianCol + 5;

    var newGuardianDataValues = masterSheetData
      .slice(1) // Skip header row
      .map(function (row) {
        return row.slice(startCol - 1, endCol);
      })
      .filter(function (row) {
        return row.some(function (cell) {
          return (
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
          );
        });
      });

    if (!newGuardianDataValues || newGuardianDataValues.length === 0) {
      return {
        success: false,
        message: "Could not read guardian data",
      };
    }

    // Filter out empty rows
    var newGuardianData = newGuardianDataValues.filter((row) =>
      row.some(
        (cell) =>
          cell !== null &&
          cell !== undefined &&
          String(cell || "").trim() !== ""
      )
    );

    var newGuardianUnlocked = [];
    var newGuardianLevel = [];

    for (var row = 0; row < newGuardianData.length; row++) {
      var rowData = newGuardianData[row];
      if (oldGuardians.hasOwnProperty(rowData[0])) {
        var oldGuardian = oldGuardians[rowData[0]];
        newGuardianUnlocked.push([rowData[0]]);
        newGuardianUnlocked.push([""]);
        newGuardianUnlocked.push([oldGuardian.unlocked]);

        for (var nextRow = row; nextRow < newGuardianData.length; nextRow++) {
          var nextRowData = newGuardianData[nextRow];
          if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
            row = nextRow - 1;
            break;
          }
          var newGuardianProp = nextRowData[2];
          if (oldGuardian.props.hasOwnProperty(newGuardianProp)) {
            newGuardianLevel.push([oldGuardian.props[newGuardianProp]]);
          } else {
            newGuardianLevel.push([nextRowData[4]]);
          }
          if (nextRow == newGuardianData.length - 1) {
            row = nextRow;
          }
        }
      } else {
        newGuardianUnlocked.push([rowData[0]]);
      }
    }

    var batchUpdate = [];
    if (newGuardianUnlocked.length > 0) {
      var unlockedCol = shared.columnToLetter(guardianCol + 1);
      var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
        newGuardianUnlocked.length + 1
      }`;
      batchUpdate.push({
        range: unlockedRange,
        values: newGuardianUnlocked,
      });
    }

    if (newGuardianLevel.length > 0) {
      var levelCol = shared.columnToLetter(guardianCol + 5);
      var levelRange = `${sheetName}!${levelCol}2:${levelCol}${
        newGuardianLevel.length + 1
      }`;
      batchUpdate.push({
        range: levelRange,
        values: newGuardianLevel,
      });
    }
    if (batchUpdate.length !== 0) {
      return {
        success: true,
        message: `Guardians updated successfully`,
        batchUpdate: batchUpdate,
      };
    }
    return {
      success: true,
      message: `No updates needed for guardians`,
    };
  },

  version10: function () {
    try {
      var oldSpreadsheet = spreadsheets("Guardians oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }

      var guardianLevelsRange = "EXPORT!B5:F";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion10Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10Guardians: function (oldGuardianLevelsData) {
    try {
      var targetGuardians = ["Attack", "Ally", "Steal", "Fetch"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {};
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        // Only proceed if guardianName is in targetGuardians
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked = oldGuardianLevels[row + 2][0];
          var guardian = {
            unlocked: unlocked,
            props: {},
          };

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              guardian.props[key] = value;
            }
          }
          oldGuardians[guardianName] = guardian;
        }
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      console.log("Error in getVersion10Guardians: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10Guardians: " + error.message,
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
