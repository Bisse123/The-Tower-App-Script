const ultimate = {
  importData: function (versionDifference) {
    function importUltimateData(versionDifference) {
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
        var result = getVersionFunction();
        if (!result || !result.success) {
          console.log(`Error processing ultimate weapons data: ${result.message}`);
          return result;
        }

        var targetWeapons = result.targetWeapons || [];
        var oldUltimate = result.oldUltimate || {};
        return updateUltimateLevels(targetWeapons, newSheetID, oldUltimate);
      } catch (error) {
        console.log("Error in importUltimateData: " + error.toString());
        return {
          success: false,
          message: "Error importing ultimate weapons data: " + error.message,
        };
      }
    }

    function version10() {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        var newSheetID = newSpreadsheet.spreadsheetId;
        
        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        var oldSheetID = oldSpreadsheet.spreadsheetId;
        
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

        if (!SheetsAPI.getSheetByName(oldSpreadsheet, "Master Sheet")) {
          console.log(`Master Sheet not found in old ultimate weapons spreadsheet`);
          return {
            success: false,
            message: `Master Sheet™ not found in old ultimate weapons spreadsheet™`,
          };
        }

        // Get all data from Master Sheet to determine range and find columns
        var sheetBatchResult = SheetsAPI.batchGetValues(oldSheetID, ["Master Sheet"]);
        if (!sheetBatchResult || sheetBatchResult.length === 0 || !sheetBatchResult[0].values) {
          console.log(`Could not read Master Sheet data`);
          return {
            success: false,
            message: `Could not read Master Sheet data`,
          };
        }
        var sheetData = sheetBatchResult[0].values;
        if (!sheetData || sheetData.length < 2) {
          console.log(`Not enough data in Master Sheet`);
          return {
            success: false,
            message: `Not enough data in Master Sheet`,
          };
        }

        var headerRow = sheetData[0];
        var ultimateCol = headerRow.indexOf("Ultimate Weapon") + 1;

        if (ultimateCol === 0) {
          console.log(`Ultimate Weapon column not found`);
          return {
            success: false,
            message: `Ultimate Weapon column not found`,
          };
        }

        // Get current ultimate weapons data starting from column after "Ultimate Weapon"
        var ultimateDataRange =
          "Master Sheet!" +
          shared.columnToLetter(ultimateCol + 1) +
          "2:" +
          shared.columnToLetter(ultimateCol + 5);

        var ultimateBatchResult = SheetsAPI.batchGetValues(
          oldSheetID,
          [ultimateDataRange]
        );
        if (!ultimateBatchResult || ultimateBatchResult.length === 0 || !ultimateBatchResult[0].values) {
          console.log(`Could not read ultimate weapons data from Master Sheet`);
          return {
            success: false,
            message: `Could not read ultimate weapons data from Master Sheet`,
          };
        }
        var oldUltimateDataValues = ultimateBatchResult[0].values;

        // Filter out empty rows
        var oldUltimateData = oldUltimateDataValues.filter((row) =>
          row.some(
            (cell) =>
              cell !== null &&
              cell !== undefined &&
              String(cell || "").trim() !== ""
          )
        );

        var oldUltimate = getOldUltimateWeapons(targetWeapons, oldUltimateData);

        return {
          success: true,
          targetWeapons: targetWeapons,
          oldUltimate: oldUltimate,
        };
      } catch (error) {
        console.log("Error in version10: " + error.toString());
        return {
          success: false,
          message: "Error in version10: " + error.message,
        };
      }
    }
    function updateUltimateLevels(targetWeapons, newSheetID, oldUltimate) {
      try {
        // Get all data from Master Sheet to determine range and find columns
        var newSheetBatchResult = SheetsAPI.batchGetValues(newSheetID, ["Master Sheet"]);
        if (!newSheetBatchResult || newSheetBatchResult.length === 0 || !newSheetBatchResult[0].values) {
          console.log(`Could not read Master Sheet data`);
          return {
            success: false,
            message: `Could not read Master Sheet data`,
          };
        }
        var sheetData = newSheetBatchResult[0].values;
        if (!sheetData || sheetData.length < 2) {
          console.log(`Not enough data in Master Sheet`);
          return {
            success: false,
            message: `Not enough data in Master Sheet`,
          };
        }

        var headerRow = sheetData[0];
        var ultimateCol = headerRow.indexOf("Ultimate Weapon") + 1;

        if (ultimateCol === 0) {
          console.log(`Ultimate Weapon column not found`);
          return {
            success: false,
            message: `Ultimate Weapon column not found`,
          };
        }

        // Get current ultimate weapons data starting from column after "Ultimate Weapon"
        var ultimateDataRange =
          "Master Sheet!" +
          shared.columnToLetter(ultimateCol + 1) +
          "2:" +
          shared.columnToLetter(ultimateCol + 5);

        var newUltimateBatchResult = SheetsAPI.batchGetValues(
          newSheetID,
          [ultimateDataRange]
        );
        if (!newUltimateBatchResult || newUltimateBatchResult.length === 0 || !newUltimateBatchResult[0].values) {
          console.log(`Could not read ultimate weapons data from new Master Sheet`);
          return {
            success: false,
            message: `Could not read ultimate weapons data from new Master Sheet`,
          };
        }
        var newUltimateDataValues = newUltimateBatchResult[0].values;
        if (!newUltimateDataValues) {
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
          var unlockedRange =
            "Master Sheet!" +
            unlockedCol +
            "2:" +
            unlockedCol +
            (newUltimateUnlocked.length + 1);
          batchUpdate.push({
            range: unlockedRange,
            values: newUltimateUnlocked,
          });
        }

        // Update the level column (5 columns after Ultimate Weapon)
        if (newUltimateLevel.length > 0) {
          var levelCol = shared.columnToLetter(ultimateCol + 5);
          var levelRange =
            "Master Sheet!" +
            levelCol +
            "2:" +
            levelCol +
            (newUltimateLevel.length + 1);
          batchUpdate.push({
            range: levelRange,
            values: newUltimateLevel,
          });
        }

        if (batchUpdate.length !== 0) {
          SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
          // console.log(`Ultimate weapons levels updated successfully`);
          return {
            success: true,
            message: `Ultimate weapons levels updated successfully`,
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
    }

    function getOldUltimateWeapons(targetWeapons, oldUltimateLevels) {
      var weapons = {};
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
          weapons[weaponName] = weapon;
        }
      }
      return weapons;
    }
    var convertVersionFunctions = {
      "v1.0": version10,
    };

    return importUltimateData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = [
      "v1.0"
    ];
    
    var sortedThresholds = versionCompatibility.slice().sort(function(a, b) {
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