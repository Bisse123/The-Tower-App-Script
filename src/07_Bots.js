const bots = {
  importData: function (versionDifference) {
    function importBotsData(versionDifference) {
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
        var result = getVersionFunction(newSheetID, oldSheetID);
        if (!result || !result.success) {
          console.log(`Error processing bots data: ${result.message}`);
          return result;
        }

        var targetBots = result.targetBots || [];
        var oldBots = result.oldBots || {};
        return updateBotLevels(targetBots, newSheetID, "Master Sheet", oldBots);
      } catch (error) {
        console.log(`Error importing bots data: ${error.toString()}`);
        return {
          success: false,
          message: `Error importing bots data: ${error.message}`,
        };
      }
    }

    function version10(newSheetID, oldSheetID) {
      try {
        var targetBots = [
          "Flame Bot",
          "Thunder Bot",
          "Golden Bot",
          "Amplify Bot",
        ];

        var newSpreadsheet = SpreadsheetApp.openById(newSheetID);
        if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
          console.log(`_IDS sheet not found in new bots spreadsheet`);
          return {
            success: false,
            message: `_IDS sheet not found in new bots spreadsheet`,
          };
        }

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
          console.log(`Master Sheet not found in new bots spreadsheet`);
          return {
            success: false,
            message: `Master Sheet not found in new bots spreadsheet`,
          };
        }

        // Get header row to find Bots column
        var headerValues = SheetsAPI.getValues(newSheetID, "_IDS!1:1");
        if (!headerValues || headerValues.length === 0) {
          console.log(`Could not read header row from _IDS sheet`);
          return {
            success: false,
            message: `Could not read header row from _IDS sheet`,
          };
        }

        var headerRow = headerValues[0];
        var importbotColStart = headerRow.indexOf("Bots");
        if (importbotColStart === -1) {
          console.log(`Bots column not found in header`);
          return {
            success: false,
            message: `Bots column not found in header`,
          };
        }

        // Get old bot levels data using Sheets API
        var oldBotLevelsData;
        try {
          var colStart = shared.columnToLetter(importbotColStart + 1);
          var colEnd = shared.columnToLetter(importbotColStart + 5);
          oldBotLevelsData = SheetsAPI.getValues(
            newSheetID,
            `_IDS!${colStart}2:${colEnd}`
          );
        } catch (error) {
          console.log(`Error getting old bot levels: ${error.toString()}`);
          return {
            success: false,
            message: `Error getting old bot levels: ${error.message}`,
          };
        }

        // Filter out empty rows
        var oldBotLevels = oldBotLevelsData.filter((row) =>
          row.some(
            (cell) =>
              cell !== null &&
              cell !== undefined &&
              String(cell || "").trim() !== ""
          )
        );

        var oldBots = getOldBots(targetBots, oldBotLevels);
        
        return {
          success: true,
          message: "Bots processed successfully",
          targetBots: targetBots,
          oldBots: oldBots,
        };
      } catch (error) {
        console.log(`Error processing bots data: ${error.toString()}`);
        return {
          success: false,
          message: `Error processing bots data: ${error.message}`,
        };
      }
    }

    function updateBotLevels(targetBots, newSheetID, sheetName, oldBots) {
      // Get all data from Master Sheet using Sheets API
      var sheetData = SheetsAPI.getDataRange(newSheetID, sheetName);
      if (!sheetData) {
        console.log(`Error getting bot master sheet data`);
        return {
          success: false,
          message: `Error getting bot master sheet data`,
        };
      }

      if (sheetData.length < 2) return; // Need at least header and one row
      if (sheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: `Not enough data in Master Sheet`,
        };
      }

      var headerRow = sheetData[0];
      var botCol = headerRow.indexOf("Bot") + 1;
      if (botCol === 0) {
        console.log(`Bot column not found in Master Sheet`);
        return {
          success: false,
          message: `Bot column not found in Master Sheet`,
        };
      }

      // Get bot data starting from row 2
      var botDataRange =
        "Master Sheet!" +
        shared.columnToLetter(botCol + 1) +
        "2:" +
        shared.columnToLetter(botCol + 5);

      var newBotDataValues = SheetsAPI.getValues(newSheetID, botDataRange);
      if (!newBotDataValues) {
        console.log(`Could not read bot data from Master Sheet`);
        return {
          success: false,
          message: `Could not read bot data from Master Sheet`,
        };
      }
      // Filter out empty rows
      var newBotData = newBotDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var newBotUnlocked = [];
      var newBotLevel = [];

      for (var row = 0; row < newBotData.length; row++) {
        var rowData = newBotData[row];
        if (oldBots.hasOwnProperty(rowData[0])) {
          var oldBot = oldBots[rowData[0]];
          newBotUnlocked.push([rowData[0]]);
          newBotUnlocked.push([""]);
          newBotUnlocked.push([""]);
          newBotUnlocked.push([oldBot.unlocked]);
          for (var nextRow = row; nextRow < newBotData.length; nextRow++) {
            var nextRowData = newBotData[nextRow];
            if (nextRow !== row && targetBots.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var newBotProp = nextRowData[2];
            if (oldBot.props.hasOwnProperty(newBotProp)) {
              newBotLevel.push([oldBot.props[newBotProp]]);
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

      var batchUpdate = [];
      // Update the data using Sheets API
      if (newBotUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(botCol + 1);
        var unlockedRange =
          sheetName +
          "!" +
          unlockedCol +
          "2:" +
          unlockedCol +
          (newBotUnlocked.length + 1);
        batchUpdate.push({
          range: unlockedRange,
          values: newBotUnlocked,
        });
      }

      if (newBotLevel.length > 0) {
        var levelCol = shared.columnToLetter(botCol + 5);
        var levelRange =
          sheetName +
          "!" +
          levelCol +
          "2:" +
          levelCol +
          (newBotLevel.length + 1);
        batchUpdate.push({
          range: levelRange,
          values: newBotLevel,
        });
      }

      if (batchUpdate.length !== 0) {
        SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        // console.log(`Bot levels updated successfully`);
        return {
          success: true,
          message: `Bot levels updated successfully`,
        };
      }
      // console.log(`No updates needed for bot levels`);
      return {
        success: true,
        message: `No updates needed for bot levels`,
      };
    }

    function getOldBots(targetBots, oldBotLevels) {
      var bots = {};
      for (var row = 0; row < oldBotLevels.length; row++) {
        var botName = oldBotLevels[row][0];
        // Only proceed if botName is in targetBots
        if (botName && targetBots.includes(botName)) {
          var unlocked = oldBotLevels[row + 3][0];
          var bot = {
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
              bot.props[key] = value;
            }
          }
          bots[botName] = bot;
        }
      }
      return bots;
    }
    var convertVersionFunctions = {
      "v1.0": version10,
    };

    return importBotsData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = {
      "v1.0": true,
    };
    
    var highestThreshold = null;
    var compatibleThreshold = null;
    
    for (var threshold in versionCompatibility) {
      var compareResult = shared.compareVersions(oldVersion, threshold);
      
      if (compareResult === "older" || compareResult === "same") {
        if (versionCompatibility[threshold]) {
          compatibleThreshold = threshold;
        }
        break;
      }
      
      if (!highestThreshold || shared.compareVersions(highestThreshold, threshold) === "older") {
        highestThreshold = threshold;
      }
    }
    
    if (!compatibleThreshold && highestThreshold) {
      var compareWithHighest = shared.compareVersions(highestThreshold, oldVersion);
      if (compareWithHighest === "older") {
        compatibleThreshold = highestThreshold;
      }
    }
    
    return compatibleThreshold;
  },
};
