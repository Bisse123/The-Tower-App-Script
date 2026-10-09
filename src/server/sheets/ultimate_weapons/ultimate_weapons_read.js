const ultimateReader = {

  /**
   * Reads Ultimate_Weapons data from a v3.1.1 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_1_1: function (oldSheetID) {
    try {
      console.log("Called: ultimateReader.version3_1_1");

      var ultimateLevelsRange = "EXPORT!C5:H";
      var ultimateBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        ultimateLevelsRange,
      ]);
      if (
        !ultimateBatchResult ||
        ultimateBatchResult.length === 0 ||
        !ultimateBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons data`,
        };
      }
      var oldUltimateDataValues = ultimateBatchResult[0].values;

      var ultimateCostCalculatorRange = "UW Cost Calculator v3";
      var costCalculatorBatchResult = SheetsAPI.batchGetFormulas(oldSheetID, [
        ultimateCostCalculatorRange,
      ]);
      if (
        !costCalculatorBatchResult ||
        costCalculatorBatchResult.length === 0 ||
        !costCalculatorBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons cost calculator data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons cost calculator data`,
        };
      }
      var oldUltimateCostCalculatorValues = costCalculatorBatchResult[0].values;

      var ultimateWeaponsData = this.getVersion3_1_1UltimateWeapons(
        oldUltimateDataValues,
      );
      if (!ultimateWeaponsData || !ultimateWeaponsData.success) {
        return ultimateWeaponsData;
      }

      var costCalculatorData = this.getVersion1_0CostCalculator(
        oldUltimateCostCalculatorValues,
      );
      if (!costCalculatorData || !costCalculatorData.success) {
        return costCalculatorData;
      }

      return {
        success: true,
        message: "Ultimate weapons processed successfully",
        oldUltimate: ultimateWeaponsData.oldUltimate,
        oldUltimateCostCalculator: costCalculatorData["UW Cost Calculator"],
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version3_1_1", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Ultimate_Weapons data from a v2.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_0: function (oldSheetID) {
    try {
      console.log("Called: ultimateReader.version2_0");

      var ultimateLevelsRange = "EXPORT!C5:G";
      var ultimateBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        ultimateLevelsRange,
      ]);
      if (
        !ultimateBatchResult ||
        ultimateBatchResult.length === 0 ||
        !ultimateBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons data`,
        };
      }
      var oldUltimateDataValues = ultimateBatchResult[0].values;

      var ultimateCostCalculatorRange = "UW Cost Calculator v3";
      var costCalculatorBatchResult = SheetsAPI.batchGetFormulas(oldSheetID, [
        ultimateCostCalculatorRange,
      ]);
      if (
        !costCalculatorBatchResult ||
        costCalculatorBatchResult.length === 0 ||
        !costCalculatorBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons cost calculator data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons cost calculator data`,
        };
      }
      var oldUltimateCostCalculatorValues = costCalculatorBatchResult[0].values;

      var ultimateWeaponsData = this.getVersion2_0UltimateWeapons(
        oldUltimateDataValues,
      );
      if (!ultimateWeaponsData || !ultimateWeaponsData.success) {
        return ultimateWeaponsData;
      }

      var costCalculatorData = this.getVersion1_0CostCalculator(
        oldUltimateCostCalculatorValues,
      );
      if (!costCalculatorData || !costCalculatorData.success) {
        return costCalculatorData;
      }

      return {
        success: true,
        message: "Ultimate weapons processed successfully",
        oldUltimate: ultimateWeaponsData.oldUltimate,
        oldUltimateCostCalculator: costCalculatorData["UW Cost Calculator"],
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version2_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Ultimate_Weapons data from a v1.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: ultimateReader.version1_0");

      var ultimateLevelsRange = "EXPORT!C5:G";
      var ultimateBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        ultimateLevelsRange,
      ]);
      if (
        !ultimateBatchResult ||
        ultimateBatchResult.length === 0 ||
        !ultimateBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons data`,
        };
      }
      var oldUltimateDataValues = ultimateBatchResult[0].values;

      var ultimateCostCalculatorRange = "UW Cost Calculator v3";
      var costCalculatorBatchResult = SheetsAPI.batchGetFormulas(oldSheetID, [
        ultimateCostCalculatorRange,
      ]);
      if (
        !costCalculatorBatchResult ||
        costCalculatorBatchResult.length === 0 ||
        !costCalculatorBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons cost calculator data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons cost calculator data`,
        };
      }
      var oldUltimateCostCalculatorValues = costCalculatorBatchResult[0].values;

      var ultimateWeaponsData = this.getVersion1_0UltimateWeapons(
        oldUltimateDataValues,
      );
      if (!ultimateWeaponsData || !ultimateWeaponsData.success) {
        return ultimateWeaponsData;
      }

      var costCalculatorData = this.getVersion1_0CostCalculator(
        oldUltimateCostCalculatorValues,
      );
      if (!costCalculatorData || !costCalculatorData.success) {
        return costCalculatorData;
      }

      return {
        success: true,
        message: "Ultimate weapons processed successfully",
        oldUltimate: ultimateWeaponsData.oldUltimate,
        oldUltimateCostCalculator: costCalculatorData["UW Cost Calculator"],
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version1_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts UltimateWeapons from a v3.1.1 sheet's values.
   * @param {Array<Array<*>>} oldUltimateDataValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion3_1_1UltimateWeapons: function (oldUltimateDataValues) {
    try {
      console.log("Called: ultimateReader.getVersion3_1_1UltimateWeapons");
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      var oldUltimateLevels = oldUltimateDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== "",
        ),
      );

      var oldUltimate = {};
      for (var row = 0; row < oldUltimateLevels.length; row++) {
        var weaponName = oldUltimateLevels[row][0];

        if (weaponName && targetWeapons.includes(weaponName)) {
          var unlocked = oldUltimateLevels[row + 2][0];
          var weapon = {
            unlocked: unlocked,
          };

          for (nextRow = row; nextRow < oldUltimateLevels.length; nextRow++) {
            var nextRowData = oldUltimateLevels[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            if (!key) {
              continue;
            }
            var level = nextRowData[4];
            var target = nextRowData[5];
            if (level) {
              if (!weapon.hasOwnProperty("levels")) {
                weapon.levels = {};
              }
              weapon.levels[key] = level;
            }
            if (target) {
              if (!weapon.hasOwnProperty("targets")) {
                weapon.targets = {};
              }
              weapon.targets[key] = target;
            }
          }
          oldUltimate[weaponName] = weapon;
        }
      }

      return {
        success: true,
        oldUltimate: oldUltimate,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion3_1_1UltimateWeapons", error, {
        oldUltimateDataValues: oldUltimateDataValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts UltimateWeapons from a v2.0 sheet's values.
   * @param {Array<Array<*>>} oldUltimateDataValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_0UltimateWeapons: function (oldUltimateDataValues) {
    try {
      console.log("Called: ultimateReader.getVersion2_0UltimateWeapons");
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      var oldUltimateLevels = oldUltimateDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== "",
        ),
      );

      var oldUltimate = {};
      for (var row = 0; row < oldUltimateLevels.length; row++) {
        var weaponName = oldUltimateLevels[row][0];

        if (weaponName && targetWeapons.includes(weaponName)) {
          var unlocked = oldUltimateLevels[row + 2][0];
          var weapon = {
            unlocked: unlocked,
          };

          for (nextRow = row; nextRow < oldUltimateLevels.length; nextRow++) {
            var nextRowData = oldUltimateLevels[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              if (!weapon.hasOwnProperty("levels")) {
                weapon.levels = {};
              }
              weapon.levels[key] = value;
            }
          }
          oldUltimate[weaponName] = weapon;
        }
      }

      return {
        success: true,
        oldUltimate: oldUltimate,
      };
    } catch (error) {
      var errorReport = errors.report("weapon.getVersion2_0UltimateWeapons", error, {
        oldUltimateDataValues: oldUltimateDataValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts UltimateWeapons from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldUltimateDataValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0UltimateWeapons: function (oldUltimateDataValues) {
    try {
      console.log("Called: ultimateReader.getVersion1_0UltimateWeapons");
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      var oldUltimateLevels = oldUltimateDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== "",
        ),
      );

      var oldUltimate = {};
      for (var row = 0; row < oldUltimateLevels.length; row++) {
        var weaponName = oldUltimateLevels[row][0];

        if (weaponName && targetWeapons.includes(weaponName)) {
          var unlocked = oldUltimateLevels[row + 2][0];
          var weapon = {
            unlocked: unlocked,
          };

          for (nextRow = row; nextRow < oldUltimateLevels.length; nextRow++) {
            var nextRowData = oldUltimateLevels[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];

            if (key && value) {
              var valueStr = value.toString();
              if (valueStr.indexOf("Locked") !== -1) {
                value = "Lo | Locked";
              } else if (valueStr.length >= 2 && /^\d{2}/.test(valueStr)) {
                var firstTwoDigits = parseInt(valueStr.substring(0, 2));
                var subtractAmount = nextRow === row + 3 ? 2 : 1;
                var modifiedFirstTwo = (firstTwoDigits - subtractAmount)
                  .toString()
                  .padStart(2, "0");
                value = modifiedFirstTwo + valueStr.substring(2);
              }
              if (!weapon.hasOwnProperty("levels")) {
                weapon.levels = {};
              }
              weapon.levels[key] = value;
            }
          }
          oldUltimate[weaponName] = weapon;
        }
      }

      return {
        success: true,
        oldUltimate: oldUltimate,
      };
    } catch (error) {
      var errorReport = errors.report("weapon.getVersion1_0UltimateWeapons", error, {
        oldUltimateDataValues: oldUltimateDataValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts CostCalculator from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldUltimateCostCalculatorValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0CostCalculator: function (oldUltimateCostCalculatorValues) {
    try {
      console.log("Called: ultimateReader.getVersion1_0CostCalculator");
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      var oldUltimateCostCalculator = {};
      var processedWeapons = {};

      for (var row = 0; row < oldUltimateCostCalculatorValues.length; row++) {
        var rowData = oldUltimateCostCalculatorValues[row];
        var uwsWantedIndex = rowData.indexOf("# Of UWs Wanted");
        var uwPlusWantedIndex = rowData.indexOf("# Of UW+ Wanted");
        if (uwsWantedIndex !== -1) {
          oldUltimateCostCalculator["# Of UWs Wanted"] =
            rowData[uwsWantedIndex + 2];
        }
        if (uwPlusWantedIndex !== -1) {
          oldUltimateCostCalculator["# Of UW+ Wanted"] =
            rowData[uwPlusWantedIndex + 2];
        }

        for (var colIndex = 0; colIndex < rowData.length; colIndex++) {
          var cellValue = rowData[colIndex];
          if (
            cellValue &&
            typeof cellValue === "string" &&
            cellValue.trim() !== ""
          ) {
            var weapon = cellValue.trim();

            if (
              processedWeapons[weapon] ||
              weapon === "# Of UWs Wanted" ||
              weapon === "# Of UW+ Wanted" ||
              weapon === "Current Value" ||
              weapon === "Target Value" ||
              weapon === "Sub"
            ) {
              break;
            }

            if (!targetWeapons.includes(weapon)) {
              continue;
            }

            processedWeapons[weapon] = true;
            oldUltimateCostCalculator[weapon] = {};
            var weaponColIndex = colIndex;
            var unlockedValue = rowData[weaponColIndex + 3];
            var unlocked =
              typeof unlockedValue === "string" && unlockedValue.startsWith("=")
                ? null
                : unlockedValue;
            if (unlocked) {
              oldUltimateCostCalculator[weapon].unlocked = unlocked;
            }

            if (row + 1 < oldUltimateCostCalculatorValues.length) {
              var subData = oldUltimateCostCalculatorValues[row + 1];
              var currentValueIndex = subData.indexOf("Current Value");
              var subNameIndex = currentValueIndex - 1;
              var targetValueIndex = subData.indexOf("Target Value");
              var modSubvalue = subData.indexOf("Sub");
              var weaponValues = {};

              for (
                var subRow = row + 2;
                subRow < oldUltimateCostCalculatorValues.length;
                subRow++
              ) {
                var subRowData = oldUltimateCostCalculatorValues[subRow];
                var subName = subRowData[subNameIndex]
                  ? subRowData[subNameIndex].toString().trim()
                  : "";
                if (
                  subName === "" ||
                  (subRowData[weaponColIndex] &&
                    targetWeapons.includes(
                      subRowData[weaponColIndex].toString().trim(),
                    ))
                ) {
                  break;
                }
                if (!weaponValues.hasOwnProperty(subName)) {
                  weaponValues[subName] = {};
                }
                var currentValue = subRowData[currentValueIndex];
                var targetValue = subRowData[targetValueIndex];
                var modSub = subRowData[modSubvalue];
                if (
                  currentValue &&
                  (typeof currentValue !== "string" ||
                    !currentValue.startsWith("="))
                ) {
                  weaponValues[subName].currentValue = currentValue;
                }
                if (
                  targetValue &&
                  (typeof targetValue !== "string" ||
                    !targetValue.startsWith("="))
                ) {
                  weaponValues[subName].targetValue = targetValue;
                }
                if (
                  modSub &&
                  (typeof modSub !== "string" || !modSub.startsWith("="))
                ) {
                  weaponValues[subName].modSub = modSub;
                }
                if (
                  weaponValues[subName] &&
                  Object.keys(weaponValues[subName]).length === 0
                ) {
                  delete weaponValues[subName];
                }
                row = subRow;
              }
              if (weaponValues && Object.keys(weaponValues).length !== 0) {
                oldUltimateCostCalculator[weapon].values = weaponValues;
              }
            }

            if (
              oldUltimateCostCalculator[weapon] &&
              Object.keys(oldUltimateCostCalculator[weapon]).length === 0
            ) {
              delete oldUltimateCostCalculator[weapon];
            }

            break;
          }
        }
      }

      return {
        success: true,
        "UW Cost Calculator": oldUltimateCostCalculator,
      };
    } catch (error) {
      var errorReport = errors.report("weapon.getVersion1_0CostCalculator", error, {
        oldUltimateCostCalculatorValues: oldUltimateCostCalculatorValues,
      });
      return errors.fail(errorReport);
    }
  },
};
