const masterReader = {

  /**
   * Reads IDS_Master data from a v4.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_0: function (oldSheetID) {
    try {
      console.log("Called: masterReader.version4_0");

      var requiredRanges = ["IDS", "Presets Presets"];
      var idsResult = SheetsAPI.batchGetValues(oldSheetID, requiredRanges);
      if (!idsResult || !idsResult[0] || !idsResult[0].values) {
        console.log(`Could not read IDS data from old spreadsheet`);
        return {
          success: false,
          message: "Could not read IDS data from old spreadsheet",
        };
      }

      var idsValues = idsResult[0].values;
      var presetsValues = idsResult[1].values;

      var idsData = this.getVersion4_0IDSData(idsValues);

      var presetsData = this.getVersion4_0PresetsData(presetsValues);

      var success = idsData.success && presetsData.success;
      var message = `IDS Master export completed successfully. IDS: ${idsData.message}, Presets: ${presetsData.message}`;

      return {
        success: success,
        message: message,
        oldIdsData: idsData.oldIdsData || {},
        oldPresetsData: presetsData.oldPresetsData || {},
      };
    } catch (error) {
      var errorReport = errors.report("masterReader.version4_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads IDS_Master data from a v2.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_0: function (oldSheetID) {
    try {
      console.log("Called: masterReader.version2_0");

      var idsResult = SheetsAPI.batchGetValues(oldSheetID, ["IDS"]);
      if (!idsResult || !idsResult[0] || !idsResult[0].values) {
        console.log(`Could not read IDS data from old spreadsheet`);
        return {
          success: false,
          message: "Could not read IDS data from old spreadsheet",
        };
      }

      var idsValues = idsResult[0].values;
      var oldIdsData = this.getVersion2_0IDSData(idsValues);
      return oldIdsData;
    } catch (error) {
      var errorReport = errors.report("masterReader.version2_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts IDSData from a v4.0 sheet's values.
   * @param {Array<Array<*>>} idsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0IDSData: function (idsValues) {
    try {
      console.log("Called: masterReader.getVersion4_0IDSData");

      var sheetReferences = {};
      var sheetTypes = [
        "Laboratory",
        "Workshop",
        "Ultimate Weapon",
        "Themes, Songs & Relics",
        "Bots",
        "Vault",
        "Cards",
        "Modules",
        "Guardians",
        "Player & Stuff",
      ];

      for (var i = 0; i < sheetTypes.length; i++) {
        var sheetType = sheetTypes[i];
        var sheetInfo = labelUtils.findSheetTypeID(
          null,
          "IDS",
          sheetType,
          idsValues,
        );
        if (sheetInfo && sheetInfo.id) {
          var sheetID = sheetRefs.extractSheetId(sheetInfo.id);
          sheetReferences[sheetType] = sheetID;
        }
      }

      return {
        success: true,
        oldIdsData: sheetReferences,
        message: "IDS Master data extracted successfully",
      };
    } catch (error) {
      var errorReport = errors.report("masterReader.getVersion4_0IDSData", error, {
        idsValues: idsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts IDSData from a v2.0 sheet's values.
   * @param {Array<Array<*>>} idsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_0IDSData: function (idsValues) {
    try {
      console.log("Called: masterReader.getVersion2_0IDSData");

      var sheetReferences = {};
      var sheetTypes = [
        "Laboratory",
        "Workshop",
        "Ultimate Weapon",
        "Themes & Songs",
        "Bots",
        "Relics",
        "Vault",
        "Cards",
        "Modules",
        "Guardians",
        "Player & Stuff",
      ];

      for (var i = 0; i < sheetTypes.length; i++) {
        var sheetType = sheetTypes[i];
        var sheetInfo = labelUtils.findSheetTypeID(
          null,
          "IDS",
          sheetType,
          idsValues,
        );
        if (sheetInfo && sheetInfo.id) {
          var sheetID = sheetRefs.extractSheetId(sheetInfo.id);
          sheetReferences[sheetType] = sheetID;
        }
      }

      return {
        success: true,
        oldIdsData: sheetReferences,
        message: "IDS Master data extracted successfully",
      };
    } catch (error) {
      var errorReport = errors.report("masterReader.getVersion2_0IDSData", error, {
        idsValues: idsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts PresetsData from a v4.0 sheet's values.
   * @param {Array<Array<*>>} presetsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0PresetsData: function (presetsValues) {
    try {
      console.log("Called: masterReader.getVersion4_0PresetsData");

      const sliders = ["Range", "Shockwave Size"];
      var presetsData = {
        data: {},
      };
      for (var row = 1; row < presetsValues.length; row++) {
        var rowData = presetsValues[row];
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
          var presetName = presetsValues[row - 2][col - 1];
          for (
            var nextRow = row + 1;
            nextRow < presetsValues.length;
            nextRow++
          ) {
            var nextRowData = presetsValues[nextRow];
            if (!nextRowData || nextRowData.length <= col) {
              continue;
            }
            var presetType = nextRowData[col];

            const isSlider = sliders.includes(presetType);
            nextRow = isSlider ? nextRow : nextRow + 1;
            var colIndex = isSlider ? col + 2 : col + 1;
            var levelValue = presetsValues[nextRow][colIndex];
            if (!presetType && !levelValue) {
              finalRow = nextRow;
              break;
            }
            if (!presetsData.data[presetName]) {
              presetsData.data[presetName] = {};
            }
            presetsData.data[presetName][presetType] = levelValue;
          }
        }
        row = finalRow;
      }

      presetsData.presetNames = presetUtils.resolvePresetOrder(
        Object.keys(presetsData.data),
        presetUtils.templatePresetNames,
      ).order;

      return {
        success: true,
        oldPresetsData: presetsData,
        message: "Presets data extracted successfully",
      };
    } catch (error) {
      var errorReport = errors.report("masterReader.getVersion4_0PresetsData", error, {
        presetsValues: presetsValues,
      });
      return errors.fail(errorReport);
    }
  },
};
