const ultimate = {
  convertVersionFunctions: {},

  importData: function (versionDifference) {
    function importUltimateData(versionDifference) {
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

        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!newSpreadsheet) {
          console.log(`New spreadsheet not found`);
          return {
            success: false,
            message: "New spreadsheet not found",
          };
        }
        var newSheetID = newSpreadsheet.spreadsheetId;

        if (versionDifference === 0) {
          // console.log(`Same Version - proceeding with ultimate weapons data import`);
          if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
            console.log(
              `_IDS sheet not found in new ultimate weapons spreadsheet`
            );
            return {
              success: false,
              message: `_IDS sheet not found in new ultimate weapons spreadsheet`,
            };
          }

          if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
            console.log(
              `Master Sheet not found in new ultimate weapons spreadsheet`
            );
            return {
              success: false,
              message: `Master Sheet not found in new ultimate weapons spreadsheet`,
            };
          }

          // Get header row to find UWs column
          var headerValues = SheetsAPI.getValues(newSheetID, "_IDS!1:1");
          if (!headerValues || headerValues.length === 0) {
            console.log(`Could not read header row from _IDS sheet`);
            return {
              success: false,
              message: `Could not read header row from _IDS sheet`,
            };
          }

          var headerRow = headerValues[0];
          var importUltimateColStart = headerRow.indexOf("UWs") + 1;

          if (importUltimateColStart === 0) {
            console.log(`UWs column not found in _IDS sheet`);
            return {
              success: false,
              message: `UWs column not found in _IDS sheet`,
            };
          }

          // Get ultimate weapons levels data (5 columns starting from UWs column)
          var ultimateLevelsRange =
            "_IDS!" +
            shared.columnToLetter(importUltimateColStart) +
            "2:" +
            shared.columnToLetter(importUltimateColStart + 4);

          var oldUltimateLevelsValues = SheetsAPI.getValues(
            newSheetID,
            ultimateLevelsRange
          );
          if (!oldUltimateLevelsValues) {
            console.log(`Could not read ultimate weapons levels data`);
            return {
              success: false,
              message: `Could not read ultimate weapons levels data`,
            };
          }

          // Filter out empty rows
          var oldUltimateLevels = oldUltimateLevelsValues.filter((row) =>
            row.some(
              (cell) =>
                cell !== null &&
                cell !== undefined &&
                String(cell || "").trim() !== ""
            )
          );

          var oldUltimate = getOldUltimateWeapons(
            targetWeapons,
            oldUltimateLevels
          );

          return updateUltimateLevels(targetWeapons, newSheetID, oldUltimate);
        } else {
          console.log(
            `Version mismatch - skipping ultimate weapons data import`
          );
          return {
            success: false,
            message: `Ultimate weapons version mismatch`,
          };
        }
      } catch (error) {
        console.log("Error in importUltimateData: " + error.toString());
        return {
          success: false,
          message: "Error importing ultimate weapons data: " + error.message,
        };
      }
    }

    function updateUltimateLevels(targetWeapons, newSheetID, oldUltimate) {
      try {
        // Get all data from Master Sheet to determine range and find columns
        var sheetData = SheetsAPI.getValues(newSheetID, "Master Sheet");
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

        var newUltimateDataValues = SheetsAPI.getValues(
          newSheetID,
          ultimateDataRange
        );
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
    return importUltimateData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
