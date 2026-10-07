const botsReader = {

  /**
   * Reads Bots data from a v3.2 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_2: function (oldSheetID) {
    try {
      console.log("Called: botsReader.version3_2");

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

  /**
   * Reads Bots data from a v3.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_0: function (oldSheetID) {
    try {
      console.log("Called: botsReader.version3_0");

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

  /**
   * Reads Bots data from a v2.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_0: function (oldSheetID) {
    try {
      console.log("Called: botsReader.version2_0");

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

  /**
   * Reads Bots data from a v1.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: botsReader.version1_0");

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

  /**
   * Extracts Bots from a v3.2 sheet's values.
   * @param {Array<Array<*>>} oldBotLevelsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion3_2Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: botsReader.getVersion3_2Bots");
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
        presetNames: presetUtils.resolvePresetOrder(
          oldBotsPresetNames,
          presetUtils.templatePresetNames,
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

          presetColumnMapping.forEach(function (presetMap) {
            var levelValue = nextRowData[presetMap.levelColIndex];

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

  /**
   * Extracts Bots from a v3.0 sheet's values.
   * @param {Array<Array<*>>} oldBotLevelsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion3_0Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: botsReader.getVersion3_0Bots");
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
        presetNames: presetUtils.resolvePresetOrder(
          oldBotsPresetNames,
          presetUtils.templatePresetNames,
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

          presetColumnMapping.forEach(function (presetMap) {
            var levelValue = nextRowData[presetMap.levelColIndex];

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

  /**
   * Extracts Bots from a v2.0 sheet's values.
   * @param {Array<Array<*>>} oldBotLevelsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_0Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: botsReader.getVersion2_0Bots");
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

  /**
   * Extracts Bots from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldBotLevelsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0Bots: function (oldBotLevelsData) {
    try {
      console.log("Called: botsReader.getVersion1_0Bots");
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
};
