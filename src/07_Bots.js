const bots = {
  importData: function (versionDifference) {
    function importBotsData(versionDifference) {
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
          console.log(`Error processing bots data: ${oldDataResult.message}`);
          return oldDataResult;
        }

        var targetBots = oldDataResult.targetBots || [];
        var oldBots = oldDataResult.oldBots || {};

        // Batch get required data for update function only
        var requiredRanges = ["Master Sheet"];
        var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
        if (!batchResults || batchResults.length === 0) {
          console.log(`Could not read required data from spreadsheet`);
          return {
            success: false,
            message: "Could not read required data from spreadsheet",
          };
        }

        var masterSheetData = batchResults[0].values;

        var botsResult = updateBotLevels(
          targetBots,
          "Master Sheet",
          oldBots,
          masterSheetData
        );
        if (!botsResult || !botsResult.success) {
          console.log(`Error updating bots: ${botsResult.message}`);
          return botsResult;
        }

        var batchUpdate = botsResult.batchUpdate || [];

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
            message: botsResult.message,
          };
        }

        return {
          success: true,
          message: `No updates needed for Bots`,
        };
      } catch (error) {
        console.log(`Error importing bots data: ${error.toString()}`);
        return {
          success: false,
          message: `Error importing bots data: ${error.message}`,
        };
      }
    }

    function updateBotLevels(targetBots, sheetName, oldBots, masterSheetData) {
      try {
        if (!masterSheetData) {
          console.log(`Error getting bot master sheet data`);
          return {
            success: false,
            message: `Error getting bot master sheet data`,
          };
        }

        if (masterSheetData.length < 2) {
          console.log(`Not enough data in Master Sheet`);
          return {
            success: false,
            message: `Not enough data in Master Sheet`,
          };
        }

        var headerRow = masterSheetData[0];
        var botCol = headerRow.indexOf("Bot") + 1;
        if (botCol === 0) {
          console.log(`Bot column not found in Master Sheet`);
          return {
            success: false,
            message: `Bot column not found in Master Sheet`,
          };
        }

        // Get bot data starting from row 2
        // Extract current bot data from pre-fetched data
        var startCol = botCol + 1; // Column after "Bot" (1-based)
        var endCol = botCol + 5; // 5 columns after "Bot"

        var newBotDataValues = masterSheetData
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

        if (!newBotDataValues || newBotDataValues.length === 0) {
          console.log(`No bot data found`);
          return {
            success: false,
            message: `No bot data found`,
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
          var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
            newBotUnlocked.length + 1
          }`;
          batchUpdate.push({
            range: unlockedRange,
            values: newBotUnlocked,
          });
        }

        if (newBotLevel.length > 0) {
          var levelCol = shared.columnToLetter(botCol + 5);
          var levelRange = `${sheetName}!${levelCol}2:${levelCol}${
            newBotLevel.length + 1
          }`;
          batchUpdate.push({
            range: levelRange,
            values: newBotLevel,
          });
        }

        if (batchUpdate.length !== 0) {
          return {
            success: true,
            message: `Bot levels updated successfully`,
            batchUpdate: batchUpdate,
          };
        }
        // console.log(`No updates needed for bot levels`);
        return {
          success: true,
          message: `No updates needed for bot levels`,
        };
      } catch (error) {
        console.log(`Error in updateBotLevels: ${error.toString()}`);
        return {
          success: false,
          message: `Error updating bot levels: ${error.message}`,
        };
      }
    }

    function version10() {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        var newSheetID = newSpreadsheet.spreadsheetId;

        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        var oldSheetID = oldSpreadsheet.spreadsheetId;

        var targetBots = [
          "Flame Bot",
          "Thunder Bot",
          "Golden Bot",
          "Amplify Bot",
        ];

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
          console.log(`_IDS sheet not found in new bots spreadsheet`);
          return {
            success: false,
            message: `_IDS sheet not found in new bots spreadsheet™`,
          };
        }

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
          console.log(`Master Sheet not found in new bots spreadsheet`);
          return {
            success: false,
            message: `Master Sheet™ not found in new bots spreadsheet™`,
          };
        }

        // Get header row to find Bots column
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
        var importbotColStart = headerRow.indexOf("Bots");
        if (importbotColStart === -1) {
          console.log(`Bots column not found in header`);
          return {
            success: false,
            message: `Bots column not found in header`,
          };
        }

        var oldBotLevelsData;
        try {
          var colStart = shared.columnToLetter(importbotColStart + 1);
          var colEnd = shared.columnToLetter(importbotColStart + 5);
          var botBatchResult = SheetsAPI.batchGetValues(newSheetID, [
            `_IDS!${colStart}2:${colEnd}`,
          ]);
          if (
            !botBatchResult ||
            botBatchResult.length === 0 ||
            !botBatchResult[0].values
          ) {
            console.log(`Could not read old bot levels data`);
            return {
              success: false,
              message: `Could not read old bot levels data`,
            };
          }
          oldBotLevelsData = botBatchResult[0].values;
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

        var oldBots = {};
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
            oldBots[botName] = bot;
          }
        }

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

    function getOldBots(targetBots, oldBotLevels) {}
    var convertVersionFunctions = {
      "v1.0": version10,
    };

    return importBotsData(versionDifference);
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
