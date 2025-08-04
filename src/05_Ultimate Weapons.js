const ultimate = {
  importData: function (versionDifference) {
    try {
      // Use sheet type-based naming for parallel execution support
      var newSpreadsheet = spreadsheets("Ultimate Weapon newSpreadsheet");
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }
      var newSheetID = newSpreadsheet.spreadsheetId;

      var oldSpreadsheet = spreadsheets("Ultimate Weapon oldSpreadsheet");
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
        console.log(`Unsupported version: ${versionDifference}`);
        return {
          success: false,
          message: `Unsupported version: ${versionDifference}`,
        };
      }
      var oldDataResult = getVersionFunction();
      if (!oldDataResult || !oldDataResult.success) {
        console.log(
          `Error processing ultimate weapons data: ${oldDataResult.message}`
        );
        return oldDataResult;
      }

      var targetWeapons = oldDataResult.targetWeapons || [];
      var oldUltimate = oldDataResult.oldUltimate || {};

      // Batch get required data for update function only
      var requiredRanges = ["Master Sheet", "IDS"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var idsData = batchResults[1].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(newSheetID, "IDS", "IDS Master's", idsData);
      if (!newSheetInfo || !newSheetInfo.importStatus || !newSheetInfo.importStatus.range) {
        console.log(`Could not find import status range in IDS sheet`);
        return {
          success: false,
          message: "Could not find import status range in IDS sheet",
        };
      }

      var ultimateResult = this.updateUltimateLevels(
        targetWeapons,
        "Master Sheet",
        oldUltimate,
        masterSheetData
      );
      if (!ultimateResult || !ultimateResult.success) {
        console.log(
          `Error updating ultimate weapon levels: ${ultimateResult.message}`
        );
        return ultimateResult;
      }

      var batchUpdate = ultimateResult.batchUpdate || [];

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
        message: `Ultimate Weapons import completed successfully`,
      };
    } catch (error) {
      console.log("Error in importUltimateData: " + error.toString());
      return {
        success: false,
        message: "Error importing ultimate weapons data: " + error.message,
      };
    }
  },

  updateUltimateLevels: function (
    targetWeapons,
    sheetName,
    oldUltimate,
    masterSheetData
  ) {
    try {
      if (!masterSheetData || masterSheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: `Not enough data in Master Sheet`,
        };
      }

      var headerRow = masterSheetData[0];
      var ultimateCol = headerRow.indexOf("Ultimate Weapon") + 1;

      if (ultimateCol === 0) {
        console.log(`Ultimate Weapon column not found`);
        return {
          success: false,
          message: `Ultimate Weapon column not found`,
        };
      }

      // Extract current ultimate weapons data from pre-fetched data
      var startCol = ultimateCol + 1; // Column after "Ultimate Weapon" (1-based)
      var endCol = ultimateCol + 5; // 5 columns after "Ultimate Weapon"

      var newUltimateDataValues = masterSheetData
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

      if (!newUltimateDataValues || newUltimateDataValues.length === 0) {
        console.log(`Could not read ultimate weapons data from Master Sheet`);
        return {
          success: false,
          message: `Could not read ultimate weapons data from Master Sheet`,
        };
      }

      // Filter out empty rows
      var newUltimateData = newUltimateDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var newUltimateUnlocked = [];
      var newUltimateLevel = [];

      for (var row = 0; row < newUltimateData.length; row++) {
        var rowData = newUltimateData[row];
        if (oldUltimate.hasOwnProperty(rowData[0])) {
          var oldWeapon = oldUltimate[rowData[0]];
          newUltimateUnlocked.push([rowData[0]]);
          newUltimateUnlocked.push([""]);
          newUltimateUnlocked.push([oldWeapon.unlocked]);

          for (
            var nextRow = row;
            nextRow < newUltimateData.length;
            nextRow++
          ) {
            var nextRowData = newUltimateData[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 2;
              break;
            }
            var newWeaponProp = nextRowData[2];
            if (oldWeapon.props.hasOwnProperty(newWeaponProp)) {
              newUltimateLevel.push([oldWeapon.props[newWeaponProp]]);
            } else {
              newUltimateLevel.push([nextRowData[4]]);
            }
            if (nextRow == newUltimateData.length - 1) {
              row = nextRow;
            }
          }
        } else {
          newUltimateUnlocked.push([rowData[0]]);
        }
      }

      var batchUpdate = [];
      // Update the unlocked column (column after Ultimate Weapon)
      if (newUltimateUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(ultimateCol + 1);
        var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
          newUltimateUnlocked.length + 1
        }`;
        batchUpdate.push({
          range: unlockedRange,
          values: newUltimateUnlocked,
        });
      }

      // Update the level column (5 columns after Ultimate Weapon)
      if (newUltimateLevel.length > 0) {
        var levelCol = shared.columnToLetter(ultimateCol + 5);
        var levelRange = `${sheetName}!${levelCol}2:${levelCol}${
          newUltimateLevel.length + 1
        }`;
        batchUpdate.push({
          range: levelRange,
          values: newUltimateLevel,
        });
      }

      if (batchUpdate.length !== 0) {
        // Return batch update data instead of calling API directly
        return {
          success: true,
          message: `Ultimate weapons levels updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      // console.log(`No updates needed for ultimate weapons levels`);
      return {
        success: true,
        message: `No updates needed for ultimate weapons levels`,
      };
    } catch (error) {
      console.log(`Error in updateUltimateLevels: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating ultimate weapons levels: ${error.message}`,
      };
    }
  },

  version10: function () {
    try {
      var oldSpreadsheet = spreadsheets("Ultimate Weapon oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old ultimate weapons spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old ultimate weapons spreadsheet",
        };
      }

      var ultimateLevelsRange = "EXPORT!C5:G";
      var ultimateBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        ultimateLevelsRange
      ]);
      if (
        !ultimateBatchResult ||
        ultimateBatchResult.length === 0 ||
        !ultimateBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons data`,
        };
      }
      var oldUltimateDataValues = ultimateBatchResult[0].values;

      return this.getVersion10Values(oldUltimateDataValues);
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10Values: function (oldUltimateDataValues) {
    try {
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      var oldUltimateLevels = oldUltimateDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldUltimate = {};
      for (var row = 0; row < oldUltimateLevels.length; row++) {
        var weaponName = oldUltimateLevels[row][0];
        // Only proceed if weaponName is in targetWeapons
        if (weaponName && targetWeapons.includes(weaponName)) {
          var unlocked = oldUltimateLevels[row + 2][0];
          var weapon = {
            unlocked: unlocked,
            props: {},
          };

          for (nextRow = row; nextRow < oldUltimateLevels.length; nextRow++) {
            var nextRowData = oldUltimateLevels[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              weapon.props[key] = value;
            }
          }
          oldUltimate[weaponName] = weapon;
        }
      }

      return {
        success: true,
        targetWeapons: targetWeapons,
        oldUltimate: oldUltimate,
      };
    } catch (error) {
      console.log("Error in getVersion10Values: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10Values: " + error.message,
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
