const bots = {
  // #region Export Functions
  exportData: function (versionDifference, oldSheetID) {
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

      var oldDataResult = getVersionFunction(oldSheetID);
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
      var errorReport = errors.report("bots.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data, newSheetID) {
    try {
      console.log("Called: bots.importData");

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
      var errorReport = errors.report("bots.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
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

      var startCol = botCol + 1;
      var endCol = startCol;

      var presetColumnMapping = [];
      var firstPresetIndex = startCol + 1;

      if (firstPresetIndex === -1) {
        console.log(`Preset columns not found in Master Sheet`);
        return {
          success: false,
          message: `Preset columns not found in Master Sheet™`,
        };
      }

      var batchUpdate = [];

      var presetSlots = [];
      headerRow.forEach(function (header, index) {
        if (index < firstPresetIndex || !String(header).trim() || presetSlots.length >= oldBots.presetNames.length) {
          return;
        }
        presetSlots.push({ header: String(header).trim(), colIndex: index });
      });

      oldBots.presetNames.forEach(function (presetName, slot) {
        var presetSlot = presetSlots[slot];
        if (!presetName || !presetSlot) {
          return;
        }
        var colIndex = presetSlot.colIndex;
        endCol = colIndex + 2;
        presetColumnMapping.push({
          presetName: presetName,
          levelColIndex: colIndex,
          toggleColIndex: endCol,
        });
        if (presetSlot.header !== presetName) {
          batchUpdate.push({
            range: `${sheetName}!${shared.columnToLetter(colIndex + 1)}1`,
            values: [[presetName]],
          });
        }
      });

      var newBotDataValues = masterSheetData
        .slice(1)
        .map(function (row) {
          return row.slice(startCol - 1, endCol);
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

      var newBotLevels = {};
      var newBotToggle = {};

      for (var row = 0; row < newBotData.length; row++) {
        var rowData = newBotData[row];
        var botName = rowData[0];
        if (!botName || !oldBots.data.hasOwnProperty(botName)) {
          continue;
        }

        var oldBot = oldBots.data[botName];
        var botStartRow = row;

        for (var nextRow = row; nextRow < newBotData.length; nextRow++) {
          var nextRowData = newBotData[nextRow];
          if (nextRow !== botStartRow && targetBots.includes(nextRowData[0])) {
            row = nextRow - 1;
            break;
          }
          var newBotProp = nextRowData[1];
          var botRowOffset = nextRow - botStartRow;

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
            
            if (!newBotToggle.hasOwnProperty(presetName)) {
              newBotToggle[presetName] = [];
            }

            // One entry per sheet row so the toggle column stays aligned with the levels column
            if (botRowOffset === 1 && presetData.hasOwnProperty("active")) {
              newBotToggle[presetName].push([presetData.active]);
            } else if (botRowOffset === 4 && presetData.hasOwnProperty("sync")) {
              newBotToggle[presetName].push([presetData.sync]);
            } else {
              newBotToggle[presetName].push([null]);
            }

            if (!newBotLevels.hasOwnProperty(presetName)) {
              newBotLevels[presetName] = [];
            }

            if (
              !presetData.hasOwnProperty("props") ||
              !newBotProp ||
              !presetData.props.hasOwnProperty(newBotProp)
            ) {
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

      if (Object.keys(newBotToggle).length > 0) {
        Object.keys(newBotToggle).forEach(function (presetName) {
          var presetMap = presetColumnMapping.find(function (map) {
            return map.presetName === presetName;
          });
          if (presetMap) {
            var toggleColLetter = shared.columnToLetter(
              presetMap.toggleColIndex + 1,
            );
            batchUpdate.push({
              range: `${sheetName}!${toggleColLetter}3:${toggleColLetter}${newBotToggle[presetName].length + 2}`,
              values: newBotToggle[presetName],
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
      var errorReport = errors.report("dvtNamedRanges.updateBotLevels", error, {
        sheetName: sheetName,
        oldBots: oldBots,
        masterSheetData: masterSheetData,
        dvtNamedRangesData: dvtNamedRangesData,
      });
      return errors.fail(errorReport);
    }
  },

  // #endregion
  // #region Convert Versions
  version3_2: function (oldSheetID) {
    try {
      console.log("Called: bots.version3_2");

      var botsLevelsRange = "EXPORT!C4:N";
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

      var botsData = this.getVersion3_2Bots(oldBotLevelsData);
      return botsData;
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version3_2", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  version3_0: function (oldSheetID) {
    try {
      console.log("Called: bots.version3_0");

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
      var errorReport = errors.report("dvtNamedRanges.version3_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  version2_0: function (oldSheetID) {
    try {
      console.log("Called: bots.version2_0");

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
      var errorReport = errors.report("dvtNamedRanges.version2_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  version1_0: function (oldSheetID) {
    try {
      console.log("Called: bots.version1_0");

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
      var errorReport = errors.report("dvtNamedRanges.version1_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  // #endregion
  // #region Get Bots
  getVersion3_2Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: bots.getVersion3_2Bots");
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

      var firstPresetIndex = 2;
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
          toggleColIndex: colIdx + 1,
        });
      }

      var oldBots = {
        presetNames: shared.resolvePresetOrder(
          oldBotsPresetNames,
          shared.templatePresetNames,
        ).order,
        data: {},
      };
      for (var row = 0; row < oldBotLevels.length; row++) {
        var botName = String(oldBotLevels[row][0] || "").trim();
        if (!botName || !targetBots.includes(botName)) {
          continue;
        }

        var bot = {
          presets: {},
        };

        presetColumnMapping.forEach(function (presetMap) {
          bot.presets[presetMap.presetName] = {
            props: {},
            active: oldBotLevels[row + 1][presetMap.toggleColIndex],
            sync: oldBotLevels[row + 4][presetMap.toggleColIndex],
          };
        });

        for (var nextRow = row; nextRow < oldBotLevels.length; nextRow++) {
          var nextRowData = oldBotLevels[nextRow];

          if (nextRow !== row && targetBots.includes(nextRowData[0])) {
            row = nextRow - 1;
            break;
          }
          var key = String(nextRowData[1] || "").trim();
          if (!key) {
            continue;
          }
          
          // var defaultLevelColIndex = presetColumnMapping[0].levelColIndex;
          // var defaultLevelValue = nextRowData[defaultLevelColIndex];
          presetColumnMapping.forEach(function (presetMap) {
            var levelValue = nextRowData[presetMap.levelColIndex];
            // if (
            //   presetMap.levelColIndex !== defaultLevelColIndex &&
            //   levelValue === defaultLevelValue
            // ) {
            //   levelValue = null;
            // }
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
      var errorReport = errors.report("dvtNamedRanges.getVersion3_2Bots", error, {
        oldBotLevelsData: oldBotLevelsData,
      });
      return errors.fail(errorReport);
    }
  },

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

      var firstPresetIndex = 4;
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
          toggleColIndex: colIdx + 1,
        });
      }

      var oldBots = {
        presetNames: shared.resolvePresetOrder(
          oldBotsPresetNames,
          shared.templatePresetNames,
        ).order,
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
          presets: {},
        };

        presetColumnMapping.forEach(function (presetMap) {
          bot.presets[presetMap.presetName] = {
            props: {},
            sync: oldBotLevels[row][presetMap.toggleColIndex],
            active:
              presetMap.levelColIndex === firstPresetIndex ? unlocked : null,
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

          // var defaultLevelColIndex = presetColumnMapping[0].levelColIndex;
          // var defaultLevelValue = nextRowData[defaultLevelColIndex];
          presetColumnMapping.forEach(function (presetMap) {
            var levelValue = nextRowData[presetMap.levelColIndex];
            // if (
            //   presetMap.levelColIndex !== defaultLevelColIndex &&
            //   levelValue === defaultLevelValue
            // ) {
            //   levelValue = null;
            // }
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
      var errorReport = errors.report("bot.getVersion3_0Bots", error, {
        oldBotLevelsData: oldBotLevelsData,
      });
      return errors.fail(errorReport);
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
            presets: {
              Farming: {
                props: {},
                sync: null,
                active: unlocked,
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
      var errorReport = errors.report("bot.getVersion2_0Bots", error, {
        oldBotLevelsData: oldBotLevelsData,
      });
      return errors.fail(errorReport);
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
            presets: {
              Farming: {
                props: {},
                sync: null,
                active: unlocked,
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
      var errorReport = errors.report("bot.getVersion1_0Bots", error, {
        oldBotLevelsData: oldBotLevelsData,
      });
      return errors.fail(errorReport);
    }
  },

  // #endregion
  // #region Parse Saved File
  parseBotsData: function (data) {
    try {
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

      const presetOrder = shared.resolvePresetOrder(
        data.presetNames || [],
        shared.templatePresetNames,
      );
      const presetNames = presetOrder.order;
      const presetIndices = presetOrder.indices;

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

      presetNames.forEach(function (presetName, slot) {
        const index = presetIndices[slot];
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
            acc[upgrade] = level ? String(level).padStart(2, "0") : "00";
            return acc;
          }, {});
        
          props[targetBots[botName].plusUpgrade] = "Lo";
          if (botPreset.plusUnlocked) {
            props[targetBots[botName].plusUpgrade] =
              botPreset.plusLevel === null ||
              botPreset.plusLevel === undefined ||
              botPreset.plusLevel === ""
                ? "Lo"
                : String(botPreset.plusLevel).padStart(2, "0");
          }
          if (!oldBots.data.hasOwnProperty(botName)) {
            oldBots.data[botName] = {
              presets: {},
            };
          }
          const active = botPreset.unlocked !== undefined ? botPreset.unlocked : null;
          oldBots.data[botName].presets[presetName] = {
            props: props,
            sync: syncPresets[index][botIndex] || null,
            active: active,
          };
        });
      });

      return {
        success: true,
        oldBots: oldBots,
        targetBots: targetBots,
        botOrder: Object.keys(targetBots),
      };
    } catch (error) {
      var errorReport = errors.report("bots.parseBotsData", error, {
        data: data,
        oldBots: oldBots,
      });
      return errors.fail(errorReport);
    }
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v2.0": this.version2_0.bind(this),
      "v3.0": this.version3_0.bind(this),
      "v3.2": this.version3_2.bind(this),
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
