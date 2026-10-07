const workshopReader = {

  /**
   * Reads Workshop data from a v2.2.8 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_2_8: function (oldSheetID) {
    try {
      console.log("Called: workshopReader.version2_2_8");

      var workshopLevelsRange = "EXPORT!B2:M";
      var workshopPlusLevelsRange = "EXPORT!P2:V";
      var workshopPlusRatioRange = "Desired Ratios";
      var valuesRanges = [
        workshopLevelsRange,
        workshopPlusLevelsRange,
        workshopPlusRatioRange,
      ];

      var updateWorkshopValuesBatchResult = SheetsAPI.batchGetValues(
        oldSheetID,
        valuesRanges,
      );
      if (
        !updateWorkshopValuesBatchResult ||
        updateWorkshopValuesBatchResult.length === 0
      ) {
        console.log(`Could not read workshop levels data`);
        return {
          success: false,
          message: `Could not read workshop levels data`,
        };
      }
      var oldWorkshopLevelsValues = updateWorkshopValuesBatchResult[0].values;
      var oldWorkshopPlusLevelsValues =
        updateWorkshopValuesBatchResult[1].values;
      var oldWorkshopPlusRatiosValues =
        updateWorkshopValuesBatchResult[2].values;

      var workshopLevelsResult = this.getVersion2_0WorkshopLevels(
        oldWorkshopLevelsValues,
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      var workshopPlusLevelsResult = this.getVersion2_0WorkshopPlusLevels(
        oldWorkshopPlusLevelsValues,
      );
      if (!workshopPlusLevelsResult || !workshopPlusLevelsResult.success) {
        return workshopPlusLevelsResult;
      }

      var workshopPlusRatiosResult = this.getVersion2_2_8WorkshopPlusRatios(
        workshopPlusLevelsResult.oldWorkshopPlusLevels.presetNames,
        oldWorkshopPlusRatiosValues,
      );
      if (!workshopPlusRatiosResult || !workshopPlusRatiosResult.success) {
        return workshopPlusRatiosResult;
      }

      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: workshopLevelsResult.oldWorkshopLevels,
        oldWorkshopPlusLevels: workshopPlusLevelsResult.oldWorkshopPlusLevels,
        oldWorkshopPlusRatios: workshopPlusRatiosResult.oldWorkshopPlusRatios,
      };
    } catch (error) {
      var errorReport = errors.report("workshopReader.version2_2_8", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Workshop data from a v2.1 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_1: function (oldSheetID) {
    try {
      console.log("Called: workshopReader.version2_1");

      var workshopLevelsRange = "EXPORT!B2:M";
      var workshopPlusLevelsRange = "EXPORT!P2:V";
      var valuesRanges = [workshopLevelsRange, workshopPlusLevelsRange];

      var updateWorkshopValuesBatchResult = SheetsAPI.batchGetValues(
        oldSheetID,
        valuesRanges,
      );
      if (
        !updateWorkshopValuesBatchResult ||
        updateWorkshopValuesBatchResult.length < 2 ||
        !updateWorkshopValuesBatchResult[0].values ||
        !updateWorkshopValuesBatchResult[1].values
      ) {
        console.log(`Could not read workshop levels data`);
        return {
          success: false,
          message: `Could not read workshop levels data`,
        };
      }
      var oldWorkshopLevelsValues = updateWorkshopValuesBatchResult[0].values;
      var oldWorkshopPlusLevelsValues =
        updateWorkshopValuesBatchResult[1].values;

      var workshopPlusRatioRange = "Desired Ratios";
      var formulasRanges = [workshopPlusRatioRange];
      var updateWorkshopFormulasBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulasRanges,
      );
      if (
        !updateWorkshopFormulasBatchResult ||
        updateWorkshopFormulasBatchResult.length < 1 ||
        !updateWorkshopFormulasBatchResult[0].values
      ) {
        console.log(`Could not read workshop plus ratios data`);
        return {
          success: false,
          message: `Could not read workshop plus ratios data`,
        };
      }
      var oldWorkshopPlusRatiosValues =
        updateWorkshopFormulasBatchResult[0].values;

      var workshopLevelsResult = this.getVersion2_0WorkshopLevels(
        oldWorkshopLevelsValues,
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      var workshopPlusLevelsResult = this.getVersion2_0WorkshopPlusLevels(
        oldWorkshopPlusLevelsValues,
      );
      if (!workshopPlusLevelsResult || !workshopPlusLevelsResult.success) {
        return workshopPlusLevelsResult;
      }

      var workshopPlusRatiosResult = this.getVersion2_1WorkshopPlusRatios(
        workshopPlusLevelsResult.oldWorkshopPlusLevels.presetNames,
        oldWorkshopPlusRatiosValues,
      );
      if (!workshopPlusRatiosResult || !workshopPlusRatiosResult.success) {
        return workshopPlusRatiosResult;
      }

      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: workshopLevelsResult.oldWorkshopLevels,
        oldWorkshopPlusLevels: workshopPlusLevelsResult.oldWorkshopPlusLevels,
        oldWorkshopPlusRatios: workshopPlusRatiosResult.oldWorkshopPlusRatios,
      };
    } catch (error) {
      var errorReport = errors.report("workshopReader.version2_1", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Workshop data from a v2.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_0: function (oldSheetID) {
    try {
      console.log("Called: workshopReader.version2_0");

      var workshopLevelsRange = "EXPORT!B2:M";
      var workshopPlusLevelsRange = "EXPORT!P2:V";
      var valuesRanges = [workshopLevelsRange, workshopPlusLevelsRange];

      var updateWorkshopValuesBatchResult = SheetsAPI.batchGetValues(
        oldSheetID,
        valuesRanges,
      );
      if (
        !updateWorkshopValuesBatchResult ||
        updateWorkshopValuesBatchResult.length < 2 ||
        !updateWorkshopValuesBatchResult[0].values ||
        !updateWorkshopValuesBatchResult[1].values
      ) {
        console.log(`Could not read workshop levels data`);
        return {
          success: false,
          message: `Could not read workshop levels data`,
        };
      }
      var oldWorkshopLevelsValues = updateWorkshopValuesBatchResult[0].values;
      var oldWorkshopPlusLevelsValues =
        updateWorkshopValuesBatchResult[1].values;

      var workshopLevelsResult = this.getVersion2_0WorkshopLevels(
        oldWorkshopLevelsValues,
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      var workshopPlusLevelsResult = this.getVersion2_0WorkshopPlusLevels(
        oldWorkshopPlusLevelsValues,
      );
      if (!workshopPlusLevelsResult || !workshopPlusLevelsResult.success) {
        return workshopPlusLevelsResult;
      }
      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: workshopLevelsResult.oldWorkshopLevels,
        oldWorkshopPlusLevels: workshopPlusLevelsResult.oldWorkshopPlusLevels,
      };
    } catch (error) {
      var errorReport = errors.report("workshopReader.version2_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Workshop data from a v1.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: workshopReader.version1_0");

      var workshopLevelsRange = "EXPORT!B3:F";
      var workshopPlusLevelsRange = "EXPORT!H3:K";

      var updateWorkshopBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        workshopLevelsRange,
        workshopPlusLevelsRange,
      ]);
      if (
        !updateWorkshopBatchResult ||
        updateWorkshopBatchResult.length < 2 ||
        !updateWorkshopBatchResult[0].values ||
        !updateWorkshopBatchResult[1].values
      ) {
        console.log(`Could not read workshop levels data`);
        return {
          success: false,
          message: `Could not read workshop levels data`,
        };
      }
      var oldWorkshopLevelsValues = updateWorkshopBatchResult[0].values;
      var oldWorkshopPlusLevelsValues = updateWorkshopBatchResult[1].values;

      var workshopLevelsResult = this.getVersion1_0WorkshopLevels(
        oldWorkshopLevelsValues,
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      var workshopPlusLevelsResult = this.getVersion1_0WorkshopPlusLevels(
        oldWorkshopPlusLevelsValues,
      );
      if (!workshopPlusLevelsResult || !workshopPlusLevelsResult.success) {
        return workshopPlusLevelsResult;
      }

      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: workshopLevelsResult.oldWorkshopLevels,
        oldWorkshopPlusLevels: workshopPlusLevelsResult.oldWorkshopPlusLevels,
      };
    } catch (error) {
      var errorReport = errors.report("workshopReader.version1_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts WorkshopLevels from a v2.0 sheet's values.
   * @param {Array<Array<*>>} oldWorkshopLevelsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_0WorkshopLevels: function (oldWorkshopLevelsValues) {
    try {
      console.log("Called: workshopReader.getVersion2_0WorkshopLevels");
      var oldWorkshopLevelsHeaders = oldWorkshopLevelsValues[0];
      var presetColumns = [];
      oldWorkshopLevelsHeaders.forEach(function (name, index) {
        if (index < 2) {
          return;
        }
        var presetName = name && name.trim() !== "" ? name.trim() : null;
        if (presetName) {
          presetColumns.push({ presetName: presetName, colIndex: index });
        }
      });

      var presetOrder = presetUtils.resolvePresetOrder(
        presetColumns.map(function (column) {
          return column.presetName;
        }),
        presetUtils.templatePresetNames,
      );
      var orderedColumns = presetOrder.indices.map(function (sourceIndex) {
        return presetColumns[sourceIndex];
      });

      var oldWorkshopLevels = {
        presetNames: presetOrder.order,
        data: {},
      };
      oldWorkshopLevelsValues.splice(0, 2);
      oldWorkshopLevelsValues.forEach(function (row) {
        var hasData = row.some(function (cell) {
          return (
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
          );
        });
        if (hasData && row[1]) {
          var levels = [];
          orderedColumns.forEach(function (column) {
            var index = column.colIndex;
            levels.push(row[index] || null, row[index + 1] || null);
          });
          oldWorkshopLevels.data[row[1]] = {
            unlocked: row[0] || null,
            levels: levels,
          };
        }
      });
      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: oldWorkshopLevels,
      };
    } catch (error) {
      var errorReport = errors.report("workshopReader.getVersion2_0WorkshopLevels", error, {
        oldWorkshopLevelsValues: oldWorkshopLevelsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts WorkshopLevels from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldWorkshopLevelsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0WorkshopLevels: function (oldWorkshopLevelsValues) {
    try {
      console.log("Called: workshopReader.getVersion1_0WorkshopLevels");
      var oldWorkshopLevels = { presetNames: [null], data: {} };
      oldWorkshopLevelsValues.forEach(function (row) {
        var hasData = row.some(function (cell) {
          return (
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
          );
        });
        if (hasData && row[1]) {
          oldWorkshopLevels.data[row[1]] = {
            unlocked: row[0] || null,
            levels: [row[2] || "", row[3] || null],
          };
        }
      });

      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: oldWorkshopLevels,
      };
    } catch (error) {
      var errorReport = errors.report("oldWorkshopLevels.getVersion1_0WorkshopLevels", error, {
        oldWorkshopLevelsValues: oldWorkshopLevelsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts WorkshopPlusLevels from a v2.0 sheet's values.
   * @param {Array<Array<*>>} oldWorkshopPlusLevelsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_0WorkshopPlusLevels: function (oldWorkshopPlusLevelsValues) {
    try {
      console.log("Called: workshopReader.getVersion2_0WorkshopPlusLevels");
      var oldWorkshopPlusLevelsHeaders = oldWorkshopPlusLevelsValues[0];
      var presetColumns = [];
      oldWorkshopPlusLevelsHeaders.forEach(function (name, index) {
        if (index < 2) {
          return;
        }
        var presetName = name && name.trim() !== "" ? name.trim() : null;
        if (presetName) {
          presetColumns.push({ presetName: presetName, colIndex: index });
        }
      });

      var presetOrder = presetUtils.resolvePresetOrder(
        presetColumns.map(function (column) {
          return column.presetName;
        }),
        presetUtils.templatePresetNames,
      );
      var orderedColumns = presetOrder.indices.map(function (sourceIndex) {
        return presetColumns[sourceIndex];
      });

      var oldWorkshopPlusLevels = {
        presetNames: presetOrder.order,
        data: {},
      };
      oldWorkshopPlusLevelsValues.splice(0, 2);
      oldWorkshopPlusLevelsValues.forEach(function (row) {
        var hasData = row.some(function (cell) {
          return (
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
          );
        });
        if (hasData && row[0]) {
          var levels = [];
          orderedColumns.forEach(function (column) {
            var index = column.colIndex;
            levels.push(row[index] || null);
          });
          oldWorkshopPlusLevels.data[row[0]] = levels;
        }
      });
      return {
        success: true,
        message: "Workshop plus levels processed successfully",
        oldWorkshopPlusLevels: oldWorkshopPlusLevels,
      };
    } catch (error) {
      var errorReport = errors.report("oldWorkshopLevels.getVersion2_0WorkshopPlusLevels", error, {
        oldWorkshopPlusLevelsValues: oldWorkshopPlusLevelsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts WorkshopPlusLevels from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldWorkshopPlusLevelsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0WorkshopPlusLevels: function (oldWorkshopPlusLevelsValues) {
    try {
      console.log("Called: workshopReader.getVersion1_0WorkshopPlusLevels");
      var oldWorkshopPlusLevels = { presetNames: [null], data: {} };
      oldWorkshopPlusLevelsValues.forEach(function (row) {
        var hasData = row.some(function (cell) {
          return (
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
          );
        });
        if (hasData && row[0]) {
          oldWorkshopPlusLevels.data[row[0]] = [row[2] || null];
        }
      });

      return {
        success: true,
        message: "Workshop plus levels processed successfully",
        oldWorkshopPlusLevels: oldWorkshopPlusLevels,
      };
    } catch (error) {
      var errorReport = errors.report("oldWorkshopPlusLevels.getVersion1_0WorkshopPlusLevels", error, {
        oldWorkshopPlusLevelsValues: oldWorkshopPlusLevelsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts WorkshopPlusRatios from a v2.2.8 sheet's values.
   * @param {*} presetNames
   * @param {Array<Array<*>>} oldWorkshopPlusRatiosValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_2_8WorkshopPlusRatios: function (
    presetNames,
    oldWorkshopPlusRatiosValues,
  ) {
    try {
      console.log("Called: workshopReader.getVersion2_2_8WorkshopPlusRatios");
      var oldWorkshopPlusRatios = {};
      var oldWorkshopPlusRatiosHeaders = oldWorkshopPlusRatiosValues[0];
      var workshopEnhancementNameCol = oldWorkshopPlusRatiosHeaders.indexOf(
        "Workshop Enhancement",
      );
      oldWorkshopPlusRatiosValues.splice(0, 1);
      for (i = 0; i < oldWorkshopPlusRatiosValues.length; i++) {
        var row = oldWorkshopPlusRatiosValues[i];
        var workshopPresetIndex = row.indexOf("Workshop preset");
        if (workshopPresetIndex !== -1) {
          var presetValue = row[workshopPresetIndex + 1];
          oldWorkshopPlusRatios["Workshop preset"] = presetNames.includes(
            presetValue,
          )
            ? presetValue
            : null;
        }
        var enhancementName = row[workshopEnhancementNameCol];
        if (enhancementName.includes("Order in between")) {
          enhancementName = "Order in between";
        }
        if (enhancementName) {
          var firstIndex = workshopEnhancementNameCol + 1;
          var lastIndex = workshopEnhancementNameCol + 10;
          oldWorkshopPlusRatios[enhancementName] = row
            .slice(firstIndex, lastIndex + 1)
            .map(function (cell) {
              return cell || "";
            });
        }
        if (enhancementName === "Order in between") {
          break;
        }
      }
      return {
        success: true,
        message: "Workshop plus ratios processed successfully",
        oldWorkshopPlusRatios: oldWorkshopPlusRatios,
      };
    } catch (error) {
      var errorReport = errors.report("oldWorkshopPlusLevels.getVersion2_2_8WorkshopPlusRatios", error, {
        presetNames: presetNames,
        oldWorkshopPlusRatiosValues: oldWorkshopPlusRatiosValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts WorkshopPlusRatios from a v2.1 sheet's values.
   * @param {*} presetNames
   * @param {Array<Array<*>>} oldWorkshopPlusRatiosValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_1WorkshopPlusRatios: function (
    presetNames,
    oldWorkshopPlusRatiosValues,
  ) {
    try {
      console.log("Called: workshopReader.getVersion2_1WorkshopPlusRatios");
      var oldWorkshopPlusRatios = {};
      var oldWorkshopPlusRatiosHeaders = oldWorkshopPlusRatiosValues[0];
      var workshopEnhancementNameCol = oldWorkshopPlusRatiosHeaders.indexOf(
        "Workshop Enhancement",
      );
      oldWorkshopPlusRatiosValues.splice(0, 1);
      for (i = 0; i < oldWorkshopPlusRatiosValues.length; i++) {
        var row = oldWorkshopPlusRatiosValues[i];
        var enhancementName = row[workshopEnhancementNameCol];
        if (enhancementName.includes("Order in between")) {
          enhancementName = "Order in between";
        }
        if (enhancementName && enhancementName.startsWith("=")) {
          var parts = enhancementName.split(",");
          if (parts.length === 4) {
            enhancementName = parts[2].trim();
          } else {
            enhancementName = parts[parts.length - 1].trim();
          }
          enhancementName = enhancementName.replace(/['"()]/g, "");
        }
        if (enhancementName) {
          enhancementName = enhancementName.replace(/\+/g, "").trim();
          var ratioValue = row[workshopEnhancementNameCol + 2];
          if (enhancementName === "Workshop preset") {
            if (!presetNames.includes(ratioValue) && /^\d+$/.test(ratioValue)) {
              ratioValue =
                presetNames[Number(ratioValue) - 1] || "Preset " + ratioValue;
            }
            oldWorkshopPlusRatios[enhancementName] = ratioValue || "";
            continue;
          } else if (enhancementName === "Base ratio") {
            ratioValue = ratioValue.replace(/\+/g, "").trim();
          }
          oldWorkshopPlusRatios[enhancementName] = [
            ratioValue || "",
            row[workshopEnhancementNameCol + 1] || "",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
          ];
        }
        if (enhancementName === "Order in between") {
          break;
        }
      }
      return {
        success: true,
        message: "Workshop plus ratios processed successfully",
        oldWorkshopPlusRatios: oldWorkshopPlusRatios,
      };
    } catch (error) {
      var errorReport = errors.report("oldWorkshopPlusLevels.getVersion2_1WorkshopPlusRatios", error, {
        presetNames: presetNames,
        oldWorkshopPlusRatiosValues: oldWorkshopPlusRatiosValues,
      });
      return errors.fail(errorReport);
    }
  },
};
