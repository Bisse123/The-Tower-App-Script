const masterWriter = {

  /**
   * Builds the batch update that writes IDSData into the new sheet.
   * @param {Array<Array<*>>} oldIDSValues
   * @param {Object} newIDSData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateIDSData: function (oldIDSValues, newIDSData) {
    try {
      console.log("Called: masterWriter.updateIDSData");
      var batchUpdate = [];

      Object.keys(oldIDSValues).forEach(function (sheetType) {
        var sheetID = oldIDSValues[sheetType];

        var sheetInfo = labelUtils.findSheetTypeID(
          null,
          "IDS",
          sheetType,
          newIDSData,
        );

        if (sheetInfo) {

          if (sheetID && sheetInfo.cell && sheetInfo.cell.range) {
            batchUpdate.push({
              range: sheetInfo.cell.range,
              values: [[sheetID]],
            });
          }
        }
      });

      return {
        success: true,
        batchUpdate: batchUpdate,
        message: `Updated ${batchUpdate.length} IDS references`,
      };
    } catch (error) {
      var errorReport = errors.report("masterWriter.updateIDSData", error, {
        oldIDSValues: oldIDSValues,
        newIDSData: newIDSData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes PresetsData into the new sheet.
   * @param {string} sheetName The tab the ranges are written to.
   * @param {Array<Array<*>>} oldPresetsValues
   * @param {Object} newPresetsData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updatePresetsData: function (sheetName, oldPresetsValues, newPresetsData) {
    try {
      console.log("Called: masterWriter.updatePresetsData");

      const sliders = ["Range", "Shockwave Size"];
      const oldPresetNames =
        oldPresetsValues.presetNames || Object.keys(oldPresetsValues.data);

      var presetIndex = 0;
      var batchUpdate = [];

      for (var row = 0; row < newPresetsData.length; row++) {
        var rowData = newPresetsData[row];
        if (
          !rowData ||
          rowData.length === 0 ||
          !rowData.some((cell) => cell === "Sliders")
        ) {
          continue;
        }
        var finalRow = row;
        for (var col = 0; col < rowData.length; col++) {
          var cellValue = rowData[col];
          if (!cellValue || cellValue !== "Sliders") {
            continue;
          }
          var presetName = oldPresetNames[presetIndex++];
          if (!presetName || !oldPresetsValues.data.hasOwnProperty(presetName)) {
            continue;
          }

          batchUpdate.push({
            range: `${sheetName}!${sheetRefs.columnToLetter(col)}${row - 1}`,
            values: [[presetName]],
          });

          var oldPresetData = oldPresetsValues.data[presetName];
          for (
            var nextRow = row + 1;
            nextRow < newPresetsData.length;
            nextRow++
          ) {
            var nextRowData = newPresetsData[nextRow];
            if (!nextRowData || nextRowData.length <= col) {
              continue;
            }
            var key = nextRowData[col];

            const isSlider = sliders.includes(key);
            nextRow = isSlider ? nextRow : nextRow + 1;
            var colIndex = isSlider ? col + 2 : col + 1;
            var levelValue = newPresetsData[nextRow][colIndex];
            if (!key && !levelValue) {
              finalRow = nextRow;
              break;
            }
            if (!oldPresetData.hasOwnProperty(key)) {
              continue;
            }
            var oldLevelValue = oldPresetData[key];
            if (oldLevelValue !== levelValue && oldLevelValue !== presetName) {
              batchUpdate.push({
                range: `${sheetName}!${sheetRefs.columnToLetter(
                  colIndex + 1,
                )}${nextRow + 1}`,
                values: [[oldLevelValue]],
              });
            }
          }
        }
        row = finalRow;
      }

      return {
        success: true,
        batchUpdate: batchUpdate,
        message: `Updated ${batchUpdate.length} Presets entries`,
      };
    } catch (error) {
      var errorReport = errors.report("masterWriter.updatePresetsData", error, {
        oldPresetsValues: oldPresetsValues,
        newPresetsData: newPresetsData,
      });
      return errors.fail(errorReport);
    }
  },
};
