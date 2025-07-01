const ultimate = {
  convertVersionFunctions: {},

  importData: function (sheetType, newUltimateSpreadsheetId) {
    function importUltimateData(sheetType, newUltimateSpreadsheetId) {
      try {
        var idType = sheetType + " ID";
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

        // Check if required sheets exist
        if (!SheetsAPI.hasSheet(newUltimateSpreadsheetId, "IDS")) {
          console.log("IDS sheet not found in new ultimate weapons spreadsheet");
          return {
            success: false,
            message: "IDS sheet not found in new ultimate weapons spreadsheet"
          };
        }
        if (!SheetsAPI.hasSheet(newUltimateSpreadsheetId, "EXPORT")) {
          console.log(
            "EXPORT sheet not found in new ultimate weapons spreadsheet"
          );
          return {
            success: false,
            message: "EXPORT sheet not found in new ultimate weapons spreadsheet"
          };
        }

        // Get version from EXPORT sheet
        var newUltimateVersion = SheetsAPI.getValue(
          newUltimateSpreadsheetId,
          "EXPORT!A1"
        );

        var idMasterSpreadsheetInfo = shared.findSheetTypeID(
          newUltimateSpreadsheetId,
          "IDS"
        );
        var idMasterSpreadsheetId = shared.extractSheetId(
          idMasterSpreadsheetInfo.id
        );
        if (!idMasterSpreadsheetId) {
          console.log("Could not find ID Master spreadsheet");
          return {
            success: false,
            message: "Could not find ID Master spreadsheet"
          };
        }

        // Get ID Master spreadsheet info
        var oldUltimateSpreadsheetInfo = shared.findSheetTypeID(
          idMasterSpreadsheetId,
          "IDS",
          idType
        );
        var oldUltimateSpreadsheetId = shared.extractSheetId(
          oldUltimateSpreadsheetInfo.id
        );
        if (!oldUltimateSpreadsheetId) {
          console.log("Could not find old ultimate spreadsheet");
          return {
            success: false,
            message: "Could not find old ultimate spreadsheet"
          };
        }

        var oldUltimateVersion = SheetsAPI.getValue(
          oldUltimateSpreadsheetId,
          "EXPORT!A1"
        );
        var versionCheck = shared.compareVersions(
          oldUltimateVersion,
          newUltimateVersion
        );

        if (versionCheck === 0) {
          console.log(
            "Same Version - proceeding with ultimate weapons data import"
          );

          // Check if _IDS sheet exists
          if (!SheetsAPI.hasSheet(newUltimateSpreadsheetId, "_IDS")) {
          console.log(
            "_IDS sheet not found in new ultimate weapons spreadsheet"
          );
          return {
            success: false,
            message: "_IDS sheet not found in new ultimate weapons spreadsheet"
          };
          }

          // Get header row to find UWs column
          var headerValues = SheetsAPI.getValues(
            newUltimateSpreadsheetId,
            "_IDS!1:1"
          );
          if (!headerValues || headerValues.length === 0) {
          console.log("Could not read header row from _IDS sheet");
          return {
            success: false,
            message: "Could not read header row from _IDS sheet"
          };
          }

          var headerRow = headerValues[0];
          var importUltimateColStart = headerRow.indexOf("UWs") + 1;

          if (importUltimateColStart === 0) {
          console.log("UWs column not found in _IDS sheet");
          return {
            success: false,
            message: "UWs column not found in _IDS sheet"
          };
          }

          // Get ultimate weapons levels data (5 columns starting from UWs column)
          var ultimateLevelsRange =
            "_IDS!" +
            shared.columnToLetter(importUltimateColStart) +
            "2:" +
            shared.columnToLetter(importUltimateColStart + 4);

          var oldUltimateLevelsValues = SheetsAPI.getValues(
            newUltimateSpreadsheetId,
            ultimateLevelsRange
          );
          if (!oldUltimateLevelsValues) {
          console.log("Could not read ultimate weapons levels data");
          return {
            success: false,
            message: "Could not read ultimate weapons levels data"
          };
          }

          // Filter out empty rows
          var oldUltimateLevels = oldUltimateLevelsValues.filter((row) =>
            row.some(
              (cell) =>
                cell !== null &&
                cell !== undefined &&
                String(cell || '').trim() !== ""
            )
          );

          var oldUltimate = getOldUltimateWeapons(
            targetWeapons,
            oldUltimateLevels
          );

          // Check if Master Sheet exists
          if (!SheetsAPI.hasSheet(newUltimateSpreadsheetId, "Master Sheet")) {
          console.log(
            "Master Sheet not found in new ultimate weapons spreadsheet"
          );
          return {
            success: false,
            message: "Master Sheet not found in new ultimate weapons spreadsheet"
          };
          }

          return updateUltimateLevels(
            targetWeapons,
            newUltimateSpreadsheetId,
            oldUltimate
          );
        } else {
          console.log(
            "Version mismatch - skipping ultimate weapons data import"
          );
        }
      } catch (error) {
        console.log("Error in importUltimateData: " + error.toString());
        return {
          success: false,
          message: "Error importing ultimate weapons data: " + error.message,
        };
      }
    }

    function updateUltimateLevels(targetWeapons, spreadsheetId, oldUltimate) {
      try {
        // Get all data from Master Sheet to determine range and find columns
        var allData = SheetsAPI.getValues(spreadsheetId, "Master Sheet");
        if (!allData || allData.length < 2) {
          console.log("Not enough data in Master Sheet");
          return {
            success: false,
            message: "Not enough data in Master Sheet"
          };
        }

        var headerRow = allData[0];
        var ultimateCol = headerRow.indexOf("Ultimate Weapon") + 1;

        if (ultimateCol === 0) {
          console.log("Ultimate Weapon column not found");
          return {
            success: false,
            message: "Ultimate Weapon column not found"
          };
        }

        // Get current ultimate weapons data starting from column after "Ultimate Weapon"
        var ultimateDataRange =
          "Master Sheet!" +
          shared.columnToLetter(ultimateCol + 1) +
          "2:" +
          shared.columnToLetter(ultimateCol + 5);

        var newUltimateDataValues = SheetsAPI.getValues(
          spreadsheetId,
          ultimateDataRange
        );
        if (!newUltimateDataValues) {
          console.log("Could not read ultimate weapons data from Master Sheet");
          return {
            success: false,
            message: "Could not read ultimate weapons data from Master Sheet"
          };
        }

        // Filter out empty rows
        var newUltimateData = newUltimateDataValues.filter((row) =>
          row.some(
            (cell) =>
              cell !== null && cell !== undefined && String(cell || '').trim() !== ""
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

        // Update the unlocked column (column after Ultimate Weapon)
        if (newUltimateUnlocked.length > 0) {
          var unlockedCol = shared.columnToLetter(ultimateCol + 1);
          var unlockedRange =
            "Master Sheet!" +
            unlockedCol +
            "2:" +
            unlockedCol +
            (newUltimateUnlocked.length + 1);
          SheetsAPI.setValues(
            spreadsheetId,
            unlockedRange,
            newUltimateUnlocked
          );
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
          SheetsAPI.setValues(
            spreadsheetId,
            levelRange,
            newUltimateLevel
          );
        }

        console.log("Ultimate weapons levels updated successfully");
        return {
          success: true,
          message: "Ultimate weapons levels updated successfully",
        };
      } catch (error) {
        console.log("Error in updateUltimateLevels: " + error.toString());
        return {
          success: false,
          message: "Error updating ultimate weapons levels: " + error.message,
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
    return importUltimateData(sheetType, newUltimateSpreadsheetId);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
