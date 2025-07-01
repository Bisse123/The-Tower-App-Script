const bots = {
  convertVersionFunctions: {},

  importData: function (sheetType, newBotSpreadsheetId) {
    function importBotsData(sheetType, newBotSpreadsheetId) {
      var idType = sheetType + " ID";
      var targetBots = [
        "Flame Bot",
        "Thunder Bot",
        "Golden Bot",
        "Amplify Bot",
      ];

      // Get new bot version using SheetsAPI
      var newBotVersion = SheetsAPI.getValue(newBotSpreadsheetId, "EXPORT!A1");
      if (!newBotVersion) {
      console.log("Error getting new bot version");
      return {
        success: false,
        message: "Error getting new bot version"
      };
      }

      // Get ID Master spreadsheet info
      var idMasterSpreadsheetInfo = shared.findSheetTypeID(
        newBotSpreadsheetId,
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

      var oldBotSpreadsheetInfo = shared.findSheetTypeID(
        idMasterSpreadsheetId,
        "IDS",
        idType
      );
      var oldBotSpreadsheetId = shared.extractSheetId(oldBotSpreadsheetInfo.id);
      if (!oldBotSpreadsheetId) {
      console.log("Could not find old bot spreadsheet");
      return {
        success: false,
        message: "Could not find old bot spreadsheet"
      };
      }

      // Get old bot version using SheetsAPI
      var oldBotVersion = SheetsAPI.getValue(oldBotSpreadsheetId, "EXPORT!A1");
      if (!oldBotVersion) {
      console.log("Error getting old bot version");
      return {
        success: false,
        message: "Error getting old bot version"
      };
      }

      var versionCheck = shared.compareVersions(oldBotVersion, newBotVersion);
      if (versionCheck === 0) {
        console.log("Same Version");

        // Get header row from _IDS sheet using Sheets API
        var headerRowData;
        try {
        var headerRowData = SheetsAPI.getValues(newBotSpreadsheetId, "_IDS!1:1")[0] || [];
        } catch (error) {
          console.log("Error getting header row: " + error.toString());
          return {
            success: false,
            message: "Error getting header row: " + error.message
          };
        }

        var importbotColStart = headerRowData.indexOf("Bots");
        if (importbotColStart === -1) {
        console.log("Bots column not found in header");
        return {
          success: false,
          message: "Bots column not found in header"
        };
        }

        // Get old bot levels data using Sheets API
        var oldBotLevelsData;
        try {
        var colStart = shared.columnToLetter(importbotColStart + 1);
        var colEnd = shared.columnToLetter(importbotColStart + 5);
        var oldBotLevelsData = SheetsAPI.getValues(
          newBotSpreadsheetId,
          "_IDS!" + colStart + "2:" + colEnd
        );
        } catch (error) {
          console.log("Error getting old bot levels: " + error.toString());
          return {
            success: false,
            message: "Error getting old bot levels: " + error.message
          };
        }

        // Filter out empty rows
        var oldBotLevels = oldBotLevelsData.filter((row) =>
          row.some(
            (cell) =>
              cell !== null && cell !== undefined && String(cell || '').trim() !== ""
          )
        );

        var oldBots = getOldBots(targetBots, oldBotLevels);
        return updateBotLevels(
          targetBots,
          newBotSpreadsheetId,
          "Master Sheet",
          oldBots
        );
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateBotLevels(targetBots, spreadsheetId, sheetName, oldBots) {
      // Get all data from Master Sheet using Sheets API
      var sheetData = SheetsAPI.getDataRange(spreadsheetId, sheetName);
      if (!sheetData) {
      console.log("Error getting bot master sheet data");
      return {
        success: false,
        message: "Error getting bot master sheet data"
      };
      }

      if (sheetData.length < 2) return; // Need at least header and one row
      if (sheetData.length < 2) {
        return {
          success: false,
          message: "Not enough data in Master Sheet"
        };
      }

      var headerRow = sheetData[0];
      var botCol = headerRow.indexOf("Bot");
      if (botCol === -1) {
      console.log("Bot column not found in Master Sheet");
      return {
        success: false,
        message: "Bot column not found in Master Sheet"
      };
      }

      // Get bot data starting from row 2
      var botDataRange = [];
      for (var i = 1; i < sheetData.length; i++) {
        var row = sheetData[i];
        if (row.length > botCol + 4) {
          botDataRange.push([
            row[botCol + 1] || "",
            row[botCol + 2] || "",
            row[botCol + 3] || "",
            row[botCol + 4] || "",
            row[botCol + 5] || "",
          ]);
        }
      }

      // Filter out empty rows
      var newBotData = botDataRange.filter((row) =>
        row.some(
          (cell) =>
            cell !== null && cell !== undefined && String(cell || '').trim() !== ""
        )
      );

      var newBotUnlocked = [];
      var newBotLevel = [];

      for (var row = 0; row < newBotData.length; row++) {
        var rowData = newBotData[row];
        if (oldBots.hasOwnProperty(rowData[0])) {
          var oldWeapon = oldBots[rowData[0]];
          newBotUnlocked.push([rowData[0]]);
          newBotUnlocked.push([""]);
          newBotUnlocked.push([""]);
          newBotUnlocked.push([oldWeapon.unlocked]);
          for (var nextRow = row; nextRow < newBotData.length; nextRow++) {
            var nextRowData = newBotData[nextRow];
            if (nextRow !== row && targetBots.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var newWeaponProp = nextRowData[2];
            if (oldWeapon.props.hasOwnProperty(newWeaponProp)) {
              newBotLevel.push([oldWeapon.props[newWeaponProp]]);
            } else {
              newBotLevel.push([nextRowData[4]]);
            }
            if (nextRow == newBotData.length - 1) {
              row = nextRow;
            }
          }
        } else {
          newBotUnlocked.push([rowData[0]]);
        }
      }

      // Update the data using Sheets API
      if (newBotUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(botCol + 2);
        var unlockedRange =
          sheetName +
          "!" +
          unlockedCol +
          "2:" +
          unlockedCol +
          (1 + newBotUnlocked.length);
        try {
          SheetsAPI.setValues(spreadsheetId, unlockedRange, newBotUnlocked);
        } catch (error) {
          console.log("Error updating bot unlocked data: " + error.toString());
          return {
            success: false,
            message: "Error updating bot unlocked data: " + error.message
          };
        }
      }

      if (newBotLevel.length > 0) {
        var levelCol = shared.columnToLetter(botCol + 6);
        var levelRange =
          sheetName +
          "!" +
          levelCol +
          "2:" +
          levelCol +
          (1 + newBotLevel.length);
        try {
          SheetsAPI.setValues(spreadsheetId, levelRange, newBotLevel);
        } catch (error) {
          console.log("Error updating bot level data: " + error.toString());
          return {
            success: false,
            message: "Error updating bot level data: " + error.message
          };
        }
      }
      console.log("Bot levels updated successfully");
      return {
        success: true,
        message: "Bot levels updated successfully",
      };
    }

    function getOldBots(targetBots, oldBotLevels) {
      var bots = {};
      for (var row = 0; row < oldBotLevels.length; row++) {
        var weaponName = oldBotLevels[row][0];
        // Only proceed if weaponName is in targetBots
        if (weaponName && targetBots.includes(weaponName)) {
          var unlocked = oldBotLevels[row + 3][0];
          var weapon = {
            unlocked: unlocked,
            props: {},
          };

          for (nextRow = row; nextRow < oldBotLevels.length; nextRow++) {
            var nextRowData = oldBotLevels[nextRow];
            if (nextRow !== row && targetBots.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              weapon.props[key] = value;
            }
          }
          bots[weaponName] = weapon;
        }
      }
      return bots;
    }
    return importBotsData(sheetType, newBotSpreadsheetId);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
