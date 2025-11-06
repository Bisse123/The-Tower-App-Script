const bots = {
  exportData: function (versionDifference) {
    try {
      console.log("Called: bots.exportData");
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
        data: oldDataResult,
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
      console.log("Called: bots.importData");
      var newSpreadsheet = spreadsheets("Bots newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      // Batch get required data for update function only
      var requiredRanges = ["Master Sheet", "IDS"];
      var dvtIndex = requiredRanges.length;
      var dvtNamedRanges = {
        "Flame Bot": {
          "Damage R.": "DVT_BOT_UG_FB_DMGR",
          Cooldown: "DVT_BOT_UG_FB_CD",
          Damage: "DVT_BOT_UG_FB_DMG",
          Range: "DVT_BOT_UG_FB_RANGE",
        },
        "Thunder Bot": {
          Duration: "DVT_BOT_UG_TB_DUR",
          Cooldown: "DVT_BOT_UG_TB_CD",
          Linger: "DVT_BOT_UG_TB_LINGER",
          Range: "DVT_BOT_UG_TB_RANGE",
        },
        "Golden Bot": {
          Duration: "DVT_BOT_UG_GB_DUR",
          Cooldown: "DVT_BOT_UG_GB_CD",
          Bonus: "DVT_BOT_UG_GB_BONUS",
          Range: "DVT_BOT_UG_GB_RANGE",
        },
        "Amplify Bot": {
          Duration: "DVT_BOT_UG_AB_DUR",
          Cooldown: "DVT_BOT_UG_AB_CD",
          Bonus: "DVT_BOT_UG_AB_BONUS",
          Range: "DVT_BOT_UG_AB_RANGE",
        },
      };

      Object.keys(dvtNamedRanges).forEach(function (bot) {
        Object.keys(dvtNamedRanges[bot]).forEach(function (prop) {
          requiredRanges.push(dvtNamedRanges[bot][prop]);
        });
      });

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

      var dvtNamedRangesData = {};
      Object.keys(dvtNamedRanges).forEach(function (bot) {
        dvtNamedRangesData[bot] = {};
        Object.keys(dvtNamedRanges[bot]).forEach(function (prop) {
          if (batchResults[dvtIndex]) {
            dvtNamedRangesData[bot][prop] = batchResults[dvtIndex].values;
          } else {
            dvtNamedRangesData[bot][prop] = [];
          }
          dvtIndex++;
        });
      });

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(
        newSheetID,
        "IDS",
        "IDS Master's",
        idsData
      );
      if (
        !newSheetInfo ||
        !newSheetInfo.importStatus ||
        !newSheetInfo.importStatus.range
      ) {
        console.log(`Could not find import status range in IDS sheet`);
        return {
          success: false,
          message: "Could not find import status range in IDS sheet",
        };
      }

      var batchUpdate = [];

      // Only update bots if key exists
      if (data.hasOwnProperty("oldBots")) {
        var oldBots = data.oldBots;
        var botsResult = this.updateBotLevels(
          "Master Sheet",
          oldBots,
          masterSheetData,
          dvtNamedRangesData
        );
        if (!botsResult || !botsResult.success) {
          console.log(`Error updating bots: ${botsResult.message}`);
          return botsResult;
        }
        batchUpdate = batchUpdate.concat(botsResult.batchUpdate || []);
      }

      // Add import status update to batch if there were data updates
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Bots",
        newSheetID,
        idsData,
        data.idMasterID
      );

      // Apply all updates (including ID setting and import status)
      var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
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

  updateBotLevels: function (
    sheetName,
    oldBots,
    masterSheetData,
    dvtNamedRangesData
  ) {
    try {
      console.log("Called: bots.updateBotLevels");
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
        var botName = rowData[0];
        if (oldBots.hasOwnProperty(botName)) {
          var oldBot = oldBots[botName];
          newBotUnlocked.push([botName]);
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
              var dvtPropValue = shared.getDVTValue(
                oldBot.props[newBotProp],
                dvtNamedRangesData[botName][newBotProp]
              );
              newBotLevel.push([dvtPropValue]);
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
      console.log("Called: bots.version20");
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
        botsLevelsRange,
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

  version10: function () {
    try {
      console.log("Called: bots.version10");
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
        botsLevelsRange,
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
      console.log("Called: bots.getVersion10Bots");
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
                var modifiedFirstTwo = (firstTwoDigits - 1)
                  .toString()
                  .padStart(2, "0");
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

  getVersion20Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: bots.getVersion20Bots");
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
