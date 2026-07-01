const bots = {
  // #region Export Functions
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

  // #endregion
  // #region Import Functions
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
          Wildfire: "DVT_BOT_UG_FB_WILDFIRE",
        },
        "Thunder Bot": {
          Duration: "DVT_BOT_UG_TB_DUR",
          Cooldown: "DVT_BOT_UG_TB_CD",
          Linger: "DVT_BOT_UG_TB_LINGER",
          Range: "DVT_BOT_UG_TB_RANGE",
          "Titan Shock": "DVT_BOT_UG_TB_TITANSHOCK",
        },
        "Golden Bot": {
          Duration: "DVT_BOT_UG_GB_DUR",
          Cooldown: "DVT_BOT_UG_GB_CD",
          Bonus: "DVT_BOT_UG_GB_BONUS",
          Range: "DVT_BOT_UG_GB_RANGE",
          "Bonus Cell": "DVT_BOT_UG_GB_BONUSCELL",
        },
        "Amplify Bot": {
          Duration: "DVT_BOT_UG_AB_DUR",
          Cooldown: "DVT_BOT_UG_AB_CD",
          Bonus: "DVT_BOT_UG_AB_BONUS",
          Range: "DVT_BOT_UG_AB_RANGE",
          "Echoing Shot": "DVT_BOT_UG_AB_ECHOINGSHOT",
        },
        "Bot Bot": {
          Duration: "DVT_BOT_UG_BB_DUR",
          Cooldown: "DVT_BOT_UG_BB_CD",
          Bonus: "DVT_BOT_UG_BB_BONUS",
          Range: "DVT_BOT_UG_BB_RANGE",
          "Maximum Power": "DVT_BOT_UG_BB_MAXIMUMPOWER",
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
        idsData,
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
          dvtNamedRangesData,
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
        data.idMasterID,
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

  // #endregion
  // #region Update Functions
  updateBotLevels: function (
    sheetName,
    oldBots,
    masterSheetData,
    dvtNamedRangesData,
  ) {
    try {
      console.log("Called: bots.updateBotLevels");
      var targetBots = [
        "Flame Bot",
        "Thunder Bot",
        "Golden Bot",
        "Amplify Bot",
        "Bot Bot",
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

      var startCol = botCol + 1; // Column after "Bot" (1-based)
      var endCol = startCol;

      var presetColumnMapping = [];
      var firstPresetIndex = headerRow.indexOf("Farming");

      var batchUpdate = [];

      headerRow.forEach(function (header, index) {
        if (index < firstPresetIndex) {
          return;
        }
        var presetName = header && header.trim() !== "" ? header.trim() : null;
        if (presetName) {
          if (oldBots.presetNames.includes(presetName)) {
            endCol = index + 1;
            presetColumnMapping.push({
              presetName: presetName,
              levelColIndex: index,
              syncColIndex: index + 1,
            });
            return;
          }
          if (presetName.startsWith("Preset")) {
            var presetNumber = presetName.substring(6).trim();
            if (presetNumber) {
              var oldPresetName = oldBots.presetNames[presetNumber - 1];
              if (oldPresetName) {
                endCol = index + 1;
                presetColumnMapping.push({
                  presetName: oldPresetName,
                  levelColIndex: index,
                  syncColIndex: index + 1,
                });
                batchUpdate.push({
                  range: `${sheetName}!${shared.columnToLetter(index + 1)}1`,
                  values: [[oldPresetName]],
                });
                return;
              }
            }
          }
        }
      });

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
            String(cell || "").trim() !== "",
        ),
      );

      var newBotUnlocked = [];
      var newBotLevels = {};
      var newBotSync = {};

      for (var row = 0; row < newBotData.length; row++) {
        var rowData = newBotData[row];
        var botName = rowData[0];
        if (!botName || !oldBots.data.hasOwnProperty(botName)) {
          continue;
        }

        var oldBot = oldBots.data[botName];
        newBotUnlocked.push([botName]);
        newBotUnlocked.push([null]);
        newBotUnlocked.push([null]);
        newBotUnlocked.push([oldBot.unlocked]);
        newBotUnlocked.push([null]);

        for (var nextRow = row; nextRow < newBotData.length; nextRow++) {
          var nextRowData = newBotData[nextRow];
          if (nextRow !== row && targetBots.includes(nextRowData[0])) {
            row = nextRow - 1;
            break;
          }
          var newBotProp = nextRowData[2];
          if (!newBotProp) {
            continue;
          }

          for (
            var presetIdx = 0;
            presetIdx < presetColumnMapping.length;
            presetIdx++
          ) {
            var presetMap = presetColumnMapping[presetIdx];
            var presetName = presetMap.presetName;
            var presetData =
              oldBot.presets && oldBot.presets[presetName]
                ? oldBot.presets[presetName]
                : null;

            if (!presetData) {
              continue;
            }

            if (
              presetData.hasOwnProperty("sync") &&
              presetData.sync !== null &&
              nextRow === row
            ) {
              newBotSync[presetName] = newBotSync[presetName] || [];
              newBotSync[presetName].push([presetData.sync]);
            } else {
              newBotSync[presetName] = newBotSync[presetName] || [];
              newBotSync[presetName].push([null]);
            }

            if (!presetData.props) {
              newBotLevels[presetName].push([null]);
              continue;
            }

            newBotLevels[presetName] = newBotLevels[presetName] || [];
            if (!presetData.props.hasOwnProperty(newBotProp)) {
              newBotLevels[presetName].push([null]);
              continue;
            }

            var oldPropValue = presetData.props[newBotProp];
            var dvtPropValue = shared.getDVTValue(
              oldPropValue,
              dvtNamedRangesData[botName][newBotProp],
            );
            newBotLevels[presetName].push([dvtPropValue]);
          }
        }
      }

      if (newBotUnlocked.length > 0) {
        var unlockedColLetter = shared.columnToLetter(botCol + 1);
        batchUpdate.push({
          range: `${sheetName}!${unlockedColLetter}3:${unlockedColLetter}${newBotUnlocked.length + 2}`,
          values: newBotUnlocked,
        });
      }

      if (Object.keys(newBotLevels).length > 0) {
        Object.keys(newBotLevels).forEach(function (presetName) {
          var presetMap = presetColumnMapping.find(function (map) {
            return map.presetName === presetName;
          });
          if (presetMap) {
            var levelColLetter = shared.columnToLetter(
              presetMap.levelColIndex + 1,
            );
            batchUpdate.push({
              range: `${sheetName}!${levelColLetter}3:${levelColLetter}${newBotLevels[presetName].length + 2}`,
              values: newBotLevels[presetName],
            });
          }
        });
      }

      if (Object.keys(newBotSync).length > 0) {
        Object.keys(newBotSync).forEach(function (presetName) {
          var presetMap = presetColumnMapping.find(function (map) {
            return map.presetName === presetName;
          });
          if (presetMap) {
            var syncColLetter = shared.columnToLetter(
              presetMap.syncColIndex + 1,
            );
            batchUpdate.push({
              range: `${sheetName}!${syncColLetter}3:${syncColLetter}${newBotSync[presetName].length + 2}`,
              values: newBotSync[presetName],
            });
          }
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

  // #endregion
  // #region Convert Versions
  version3_0: function () {
    try {
      console.log("Called: bots.version3_0");
      var oldSpreadsheet = spreadsheets("Bots oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }

      var botsLevelsRange = "EXPORT!C4:L";
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

      var botsData = this.getVersion3_0Bots(oldBotLevelsData);
      return botsData;
    } catch (error) {
      console.log("Error in version3_0: " + error.toString());
      return {
        success: false,
        message: "Error in version3_0: " + error.message,
      };
    }
  },

  version2_0: function () {
    try {
      console.log("Called: bots.version2_0");
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

      var botsData = this.getVersion2_0Bots(oldBotLevelsData);
      return botsData;
    } catch (error) {
      console.log("Error in version2_0: " + error.toString());
      return {
        success: false,
        message: "Error in version2_0: " + error.message,
      };
    }
  },

  version1_0: function () {
    try {
      console.log("Called: bots.version1_0");
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

      var botsData = this.getVersion1_0Bots(oldBotLevelsData);
      return botsData;
    } catch (error) {
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Bots
  getVersion3_0Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: bots.getVersion3_0Bots");
      var targetBots = [
        "Flame Bot",
        "Thunder Bot",
        "Golden Bot",
        "Amplify Bot",
        "Bot Bot",
      ];

      var oldBotLevels = oldBotLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== "",
        ),
      );

      if (!oldBotLevels || oldBotLevels.length === 0) {
        return {
          success: false,
          message: "No bot levels data found",
        };
      }

      var oldBotsHeaderRow = oldBotLevels[0] || [];
      var oldBotsPresetNames = [];
      var presetColumnMapping = [];

      var firstPresetIndex = oldBotsHeaderRow.indexOf("Farming");
      for (
        var colIdx = firstPresetIndex;
        colIdx < oldBotsHeaderRow.length;
        colIdx++
      ) {
        var presetName = String(oldBotsHeaderRow[colIdx] || "").trim();
        if (!presetName) {
          continue;
        }

        oldBotsPresetNames.push(presetName);
        presetColumnMapping.push({
          presetName: presetName,
          levelColIndex: colIdx,
          syncColIndex: colIdx + 1,
        });
      }

      var oldBots = {
        presetNames: oldBotsPresetNames,
        data: {},
      };
      for (var row = 0; row < oldBotLevels.length; row++) {
        var botName = String(oldBotLevels[row][0] || "").trim();
        if (!botName || !targetBots.includes(botName)) {
          continue;
        }

        var unlocked =
          oldBotLevels[row + 3] && oldBotLevels[row + 3][0]
            ? oldBotLevels[row + 3][0]
            : null;

        var bot = {
          unlocked: unlocked,
          presets: {},
        };

        presetColumnMapping.forEach(function (presetMap) {
          bot.presets[presetMap.presetName] = {
            props: {},
            sync: oldBotLevels[row][presetMap.syncColIndex],
          };
        });

        for (var nextRow = row; nextRow < oldBotLevels.length; nextRow++) {
          var nextRowData = oldBotLevels[nextRow];

          if (nextRow !== row && targetBots.includes(nextRowData[0])) {
            row = nextRow - 1;
            break;
          }
          var key = String(nextRowData[2] || "").trim();
          if (!key) {
            continue;
          }
          var defaultLevelValue = nextRowData[firstPresetIndex];
          presetColumnMapping.forEach(function (presetMap) {
            var levelValue = nextRowData[presetMap.levelColIndex];
            if (
              presetMap.presetName !== "Farming" &&
              levelValue === defaultLevelValue
            ) {
              levelValue = null;
            }
            bot.presets[presetMap.presetName].props[key] = levelValue;
          });
        }

        oldBots.data[botName] = bot;
      }

      return {
        success: true,
        message: "Bots processed successfully",
        oldBots: oldBots,
      };
    } catch (error) {
      console.log("Error in getVersion3_0Bots: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion3_0Bots: " + error.message,
      };
    }
  },

  getVersion2_0Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: bots.getVersion2_0Bots");
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
            String(cell || "").trim() !== "",
        ),
      );

      var oldBots = {
        presetNames: ["Farming"],
        data: {},
      };
      for (var row = 0; row < oldBotLevels.length; row++) {
        var botName = oldBotLevels[row][0];
        // Only proceed if botName is in targetBots
        if (botName && targetBots.includes(botName)) {
          var unlocked = oldBotLevels[row + 3][0];
          var bot = {
            unlocked: unlocked,
            presets: {
              Farming: {
                props: {},
                sync: null,
              },
            },
          };

          for (var nextRow = row; nextRow < oldBotLevels.length; nextRow++) {
            var nextRowData = oldBotLevels[nextRow];
            if (nextRow !== row && targetBots.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              bot.presets.Farming.props[key] = value;
            }
          }
          oldBots.data[botName] = bot;
        }
      }

      return {
        success: true,
        message: "Bots processed successfully",
        oldBots: oldBots,
      };
    } catch (error) {
      console.log("Error in getVersion2_0Bots: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion2_0Bots: " + error.message,
      };
    }
  },

  getVersion1_0Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: bots.getVersion1_0Bots");
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
            String(cell || "").trim() !== "",
        ),
      );

      var oldBots = {
        presetNames: ["Farming"],
        data: {},
      };
      for (var row = 0; row < oldBotLevels.length; row++) {
        var botName = oldBotLevels[row][0];
        // Only proceed if botName is in targetBots
        if (botName && targetBots.includes(botName)) {
          var unlocked = oldBotLevels[row + 3][0];
          var bot = {
            unlocked: unlocked,
            presets: {
              Farming: {
                props: {},
                sync: null,
              },
            },
          };

          for (var nextRow = row; nextRow < oldBotLevels.length; nextRow++) {
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
              bot.presets.Farming.props[key] = value;
            }
          }
          oldBots.data[botName] = bot;
        }
      }

      return {
        success: true,
        message: "Bots processed successfully",
        oldBots: oldBots,
      };
    } catch (error) {
      console.log("Error in getVersion1_0Bots: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0Bots: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseBotsData: function (data) {
    const targetBots = {
      "Flame Bot": {
        upgrades: ["Damage", "Range", "Cooldown", "Damage R."],
        plusUpgrade: "Wildfire",
      },
      "Thunder Bot": {
        upgrades: ["Linger", "Range", "Cooldown", "Duration"],
        plusUpgrade: "Titan Shock",
      },
      "Golden Bot": {
        upgrades: ["Bonus", "Range", "Cooldown", "Duration"],
        plusUpgrade: "Bonus Cell",
      },
      "Amplify Bot": {
        upgrades: ["Bonus", "Range", "Cooldown", "Duration"],
        plusUpgrade: "Echoing Shot",
      },
      "Bot Bot": {
        upgrades: ["Bonus", "Range", "Cooldown", "Duration"],
        plusUpgrade: "Maximum Power",
      },
    };

    const presetNameOverride = ["Farming", "Tourney"];
    
    const presetNames = (data.presetNames || []).map((name, index) => (presetNameOverride[index] || name || `Preset ${index + 1}`));
    const flameBotData = data.flameBotPresets || {};
    const thunderBotData = data.thunderBotPresets || {};
    const goldenBotData = data.goldenBotPresets || {};
    const amplifyBotData = data.amplifyBotPresets || {};
    const botBotData = data.botBotPresets || {};

    const syncPresets = data.synchronicityPresets || {};
    var oldBots = {
      presetNames: presetNames,
      data: {},
    };

    presetNames.forEach(function (presetName, index) {
      const allBotPresets = [
        flameBotData[index] || {},
        thunderBotData[index] || {},
        goldenBotData[index] || {},
        amplifyBotData[index] || {},
        botBotData[index] || {},
      ];
      allBotPresets.forEach(function (botPreset, botIndex) {
        var botName = Object.keys(targetBots)[botIndex];
        var botLevels = botPreset.levels || [];
        var props = targetBots[botName].upgrades.reduce(function (
          acc,
          upgrade,
          idx,
        ) {
          var level = botLevels[idx];
          acc[upgrade] = level ? String(level).padStart(2, "0") : null ;
          return acc;
        }, {});
        props[targetBots[botName].plusUpgrade] = botPreset.plusUnlocked
          ? (botPreset.plusLevel ? String(botPreset.plusLevel).padStart(2, "0") : null)
          : "Lo";
        if (!oldBots.data.hasOwnProperty(botName)) {
          oldBots.data[botName] = {
            unlocked: false,
            presets: {},
          };
        }
        if (botPreset.unlocked) {
          oldBots.data[botName].unlocked = botPreset.unlocked;
        }
        oldBots.data[botName].presets[presetName] = {
          props: props,
          sync: syncPresets[index][botIndex] || null,
        };
      });
    });

    return {
      oldBots: oldBots,
      targetBots: targetBots,
      botOrder: Object.keys(targetBots),
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v2.0": this.version2_0.bind(this),
      "v3.0": this.version3_0.bind(this),
    };
  },

  // #endregion
  // #region Compatibility Check
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

  // #endregion
};
