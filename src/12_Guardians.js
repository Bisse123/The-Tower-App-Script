const guardians = {
  importData: function (versionDifference) {
    function importGuardiansData(versionDifference) {
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

        var getVersionFunction = convertVersionFunctions[versionDifference];
        if (!getVersionFunction) {
          console.log(`Unsupported version difference: ${versionDifference}`);
          return {
            success: false,
            message: `Unsupported version difference: ${versionDifference}`,
          };
        }
        var oldDataResult = getVersionFunction();
        if (!oldDataResult || !oldDataResult.success) {
          console.log(
            `Error processing guardians data: ${oldDataResult.message}`
          );
          return oldDataResult;
        }

        var targetGuardians = oldDataResult.targetGuardians || [];
        var oldGuardians = oldDataResult.oldGuardians || {};

        // Batch fetch required sheet data
        var requiredRanges = ["Master Sheet"];
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

        var guardiansResult = updateGuardianLevels(
          targetGuardians,
          "Master Sheet",
          oldGuardians,
          masterSheetData
        );
        if (!guardiansResult || !guardiansResult.success) {
          console.log(`Error updating guardians: ${guardiansResult.message}`);
          return guardiansResult;
        }

        var batchUpdate = guardiansResult.batchUpdate || [];

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
            message: guardiansResult.message,
          };
        }

        return {
          success: true,
          message: `No updates needed for Guardians`,
        };
      } catch (error) {
        console.log(`Error in importGuardiansData: ${error.toString()}`);
        return {
          success: false,
          message: `Error in importGuardiansData: ${error.message}`,
        };
      }
    }

    function updateGuardianLevels(
      targetGuardians,
      sheetName,
      oldGuardians,
      masterSheetData
    ) {
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

      // Extract current guardian data from pre-fetched data
      var startCol = guardianCol + 1; // Column after "Guardian Weapon" (1-based)
      var endCol = guardianCol + 5; // 5 columns after "Guardian Weapon"

      var newGuardianDataValues = masterSheetData
        .slice(1) // Skip header row
        .map(function (row) {
          return row.slice(startCol - 1, endCol); // Extract columns (convert to 0-based)
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
      // Update the data using SheetsAPI
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
        // Return batch update data instead of calling API directly
        return {
          success: true,
          message: `Guardians updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      // console.log(`No updates needed for guardians`);
      return {
        success: true,
        message: `No updates needed for guardians`,
      };
    }

    function version10() {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        var newSheetID = newSpreadsheet.spreadsheetId;

        var targetGuardians = ["Attack", "Ally", "Steal", "Fetch"];

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
          console.log(`_IDS sheet not found in new guardians spreadsheet`);
          return {
            success: false,
            message: `_IDS sheet not found in new guardians spreadsheet™`,
          };
        }

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
          console.log(`Master Sheet not found in new guardians spreadsheet`);
          return {
            success: false,
            message: `Master Sheet™ not found in new guardians spreadsheet™`,
          };
        }

        // Get header row to find Guardians column
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
            message: `Could not read header row from _IDS sheet`,
          };
        }

        var headerValues = headerBatchResult[0].values;
        var headerRow = headerValues[0];
        var importGuardianColStart = headerRow.indexOf("Guardians");
        if (importGuardianColStart === -1) {
          console.log(`Guardians column not found in header`);
          return {
            success: false,
            message: `Guardians column not found in header`,
          };
        }

        var oldGuardianLevelsData;
        try {
          var colStart = shared.columnToLetter(importGuardianColStart + 1);
          var colEnd = shared.columnToLetter(importGuardianColStart + 5);
          var guardianBatchResult = SheetsAPI.batchGetValues(newSheetID, [
            `_IDS!${colStart}2:${colEnd}`,
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
          oldGuardianLevelsData = guardianBatchResult[0].values;
        } catch (error) {
          console.log(`Error getting old guardian levels: ${error.toString()}`);
          return {
            success: false,
            message: `Error getting old guardian levels: ${error.message}`,
          };
        }

        // Filter out empty rows
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
          targetGuardians: targetGuardians,
          oldGuardians: oldGuardians,
        };
      } catch (error) {
        console.log("Error in version10: " + error.toString());
        return {
          success: false,
          message: "Error in version10: " + error.message,
        };
      }
    }

    var convertVersionFunctions = {
      "v1.0": version10,
    };

    return importGuardiansData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = ["v1.0"];

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
