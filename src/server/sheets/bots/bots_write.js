const botsWriter = {

  /**
   * Builds the batch update that writes BotLevels into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldBots
   * @param {Object} masterSheetData
   * @param {Object} dvtNamedRangesData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateBotLevels: function (
    sheetName,
    oldBots,
    masterSheetData,
    dvtNamedRangesData,
  ) {
    try {
      console.log("Called: botsWriter.updateBotLevels");
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
            range: `${sheetName}!${sheetRefs.columnToLetter(colIndex + 1)}1`,
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
            var dvtPropValue = dropdownUtils.getDVTValue(
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
            var levelColLetter = sheetRefs.columnToLetter(
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
            var toggleColLetter = sheetRefs.columnToLetter(
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
};
