const bots = {
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
        message: "Bots export completed successfully",
        data: {
          oldBots: oldDataResult.oldBots || {}
        }
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting bots data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      // Use sheet type-based naming for parallel execution support
      var newSpreadsheet = spreadsheets("Bots newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      var oldBots = data.oldBots || {};

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

      var botsResult = this.updateBotLevels(
        "Master Sheet",
        oldBots,
        masterSheetData
      );
      if (!botsResult || !botsResult.success) {
        console.log(`Error updating bots: ${botsResult.message}`);
        return botsResult;
      }

      var batchUpdate = botsResult.batchUpdate || [];

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
        message: `Bots import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing bots data: ${error.message}`,
      };
    }
  },

  updateBotLevels: function (sheetName, oldBots, masterSheetData) {
    try {
      var targetBots = [
        "Flame Bot",
        "Thunder Bot",
        "Golden Bot",
        "Amplify Bot",
      ];

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
  },

  version20: function () {
    try {
      var oldSpreadsheet = spreadsheets("Bots oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }

      var botsLevelsRange = "EXPORT!C5:G";
      var botBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        botsLevelsRange
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
      var oldBotLevelsData = botBatchResult[0].values;

      var botsData = this.getVersion20Bots(oldBotLevelsData);
      return botsData;
    } catch (error) {
      console.log("Error in version20: " + error.toString());
      return {
        success: false,
        message: "Error in version20: " + error.message,
      };
    }
  },

  getVersion20Bots: function (oldBotLevelsData) {
    try {
      var targetBots = [
        "Flame Bot",
        "Thunder Bot",
        "Golden Bot",
        "Amplify Bot",
      ];

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
        oldBots: oldBots,
      };
    } catch (error) {
      console.log("Error in getVersion20Bots: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion20Bots: " + error.message,
      };
    }
  },

  version10: function () {
    try {
      var oldSpreadsheet = spreadsheets("Bots oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }

      var botsLevelsRange = "EXPORT!C5:G";
      var botBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        botsLevelsRange
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
      var oldBotLevelsData = botBatchResult[0].values;

      var botsData = this.getVersion10Bots(oldBotLevelsData);
      return botsData;
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10Bots: function (oldBotLevelsData) {
    try {
      var targetBots = [
        "Flame Bot",
        "Thunder Bot",
        "Golden Bot",
        "Amplify Bot",
      ];

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
              var valueStr = value.toString();
              if (valueStr.length >= 2 && /^\d{2}/.test(valueStr)) {
                var firstTwoDigits = parseInt(valueStr.substring(0, 2));
                var modifiedFirstTwo = (firstTwoDigits - 1).toString().padStart(2, '0');
                value = modifiedFirstTwo + valueStr.substring(2);
              }
              bot.props[key] = value;
            }
          }
          oldBots[botName] = bot;
        }
      }

      return {
        success: true,
        message: "Bots processed successfully",
        oldBots: oldBots,
      };
    } catch (error) {
      console.log("Error in getVersion10Bots: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10Bots: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
      "v2.0": this.version20.bind(this),
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
