const ultimate = {

  /**
   * Reads Ultimate_Weapons data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: ultimate.exportData");
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
        message: "Ultimate weapons export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("ultimate.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Ultimate_Weapons data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: ultimate.importData");

      var requiredRanges = ["Master Sheet", "UW Cost Calculator v3", "IDS"];
      var dvtIndex = requiredRanges.length;
      var dvtNamedRanges = {
        "Chain Lightning": {
          Damage: "DVT_UW_UG_CL_DMG",
          Quantity: "DVT_UW_UG_CL_QNT",
          Chance: "DVT_UW_UG_CL_CH",
          Smite: "DVT_UW_UG_CL_SM",
        },
        "Smart Missiles": {
          Damage: "DVT_UW_UG_SM_DMG",
          Quantity: "DVT_UW_UG_SM_QNT",
          Cooldown: "DVT_UW_UG_SM_CD",
          "Cover Fire": "DVT_UW_UG_SM_CF",
        },
        "Death Wave": {
          Damage: "DVT_UW_UG_DW_DMG",
          Quantity: "DVT_UW_UG_DW_QNT",
          Cooldown: "DVT_UW_UG_DW_CD",
          "Kill Wall": "DVT_UW_UG_DW_KW",
        },
        "Chrono Field": {
          Duration: "DVT_UW_UG_CF_DU",
          "Speed Reduction": "DVT_UW_UG_CF_SP",
          Cooldown: "DVT_UW_UG_CF_CD",
          "Chrono Loop": "DVT_UW_UG_CF_CL",
        },
        "Inner Land Mines": {
          Damage: "DVT_UW_UG_ILM_DMG",
          Quantity: "DVT_UW_UG_ILM_QNT",
          Cooldown: "DVT_UW_UG_ILM_CD",
          "Charged Mines": "DVT_UW_UG_ILM_CM",
        },
        "Golden Tower": {
          Multiplier: "DVT_UW_UG_GT_M",
          Duration: "DVT_UW_UG_GT_DU",
          Cooldown: "DVT_UW_UG_GT_CD",
          "Golden Combo": "DVT_UW_UG_GT_GC",
        },
        "Poison Swamp": {
          Damage: "DVT_UW_UG_PS_DMG",
          Duration: "DVT_UW_UG_PS_DU",
          Cooldown: "DVT_UW_UG_PS_CH",
          "Death Creep": "DVT_UW_UG_PS_DC",
        },
        "Black Hole": {
          Size: "DVT_UW_UG_BH_SZ",
          Duration: "DVT_UW_UG_BH_DU",
          Cooldown: "DVT_UW_UG_BH_CD",
          Consume: "DVT_UW_UG_BH_C",
        },
        Spotlight: {
          Multiplier: "DVT_UW_UG_SL_MU",
          Angle: "DVT_UW_UG_SL_AN",
          Quantity: "DVT_UW_UG_SL_QNT",
          "Light Range": "DVT_UW_UG_SL_LR",
        },
      };
      Object.keys(dvtNamedRanges).forEach(function (weapon) {
        Object.keys(dvtNamedRanges[weapon]).forEach(function (level) {
          requiredRanges.push(dvtNamedRanges[weapon][level]);
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
      var ultimateCostCalculatorData = batchResults[1].values;
      var idsData = batchResults[2].values;

      var dvtNamedRangesData = {};
      Object.keys(dvtNamedRanges).forEach(function (weapon) {
        dvtNamedRangesData[weapon] = {};
        Object.keys(dvtNamedRanges[weapon]).forEach(function (level) {
          if (batchResults[dvtIndex]) {
            dvtNamedRangesData[weapon][level] = batchResults[dvtIndex].values;
          } else {
            dvtNamedRangesData[weapon][level] = [];
          }
          dvtIndex++;
        });
      });

      var batchUpdate = [];

      if (data.hasOwnProperty("oldUltimate")) {
        var oldUltimate = data.oldUltimate;
        var ultimateResult = this.updateUltimateLevels(
          "Master Sheet",
          oldUltimate,
          masterSheetData,
          dvtNamedRangesData,
        );
        if (!ultimateResult || !ultimateResult.success) {
          console.log(
            `Error updating ultimate weapon levels: ${ultimateResult.message}`,
          );
          return ultimateResult;
        }
        batchUpdate = batchUpdate.concat(ultimateResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldUltimateCostCalculator")) {
        var oldUltimateCostCalculator = data.oldUltimateCostCalculator;
        var ultimateCostCalculatorResult = this.updateUltimateCostCalculator(
          "UW Cost Calculator v3",
          oldUltimateCostCalculator,
          ultimateCostCalculatorData,
        );
        if (
          !ultimateCostCalculatorResult ||
          !ultimateCostCalculatorResult.success
        ) {
          console.log(
            `Error updating ultimate cost calculator: ${ultimateCostCalculatorResult.message}`,
          );
          return ultimateCostCalculatorResult;
        }
        batchUpdate = batchUpdate.concat(
          ultimateCostCalculatorResult.batchUpdate || [],
        );
      }

      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Ultimate Weapon",
        newSheetID,
        idsData,
        data.idMasterID,
      );

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
        message: `Ultimate Weapons import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("ultimate.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes UltimateLevels into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldUltimate
   * @param {Object} masterSheetData
   * @param {Object} dvtNamedRangesData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateUltimateLevels: function (
    sheetName,
    oldUltimate,
    masterSheetData,
    dvtNamedRangesData,
  ) {
    try {
      console.log("Called: ultimate.updateUltimateLevels");
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

      if (!masterSheetData || masterSheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: `Not enough data in Master Sheet`,
        };
      }

      var headerRow = masterSheetData[0];
      var ultimateCol = headerRow.indexOf("Ultimate Weapon") + 1;

      if (ultimateCol === 0) {
        console.log(`Ultimate Weapon column not found`);
        return {
          success: false,
          message: `Ultimate Weapon column not found`,
        };
      }

      var startCol = ultimateCol + 1;
      var endCol = ultimateCol + 6;

      var newUltimateData = masterSheetData
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

      if (!newUltimateData || newUltimateData.length === 0) {
        console.log(`Could not read ultimate weapons data from Master Sheet`);
        return {
          success: false,
          message: `Could not read ultimate weapons data from Master Sheet`,
        };
      }

      var newUltimateUnlocked = [];
      var newUltimateLevel = [];

      for (var row = 0; row < newUltimateData.length; row++) {
        var rowData = newUltimateData[row];
        var weaponName = rowData[0];
        if (oldUltimate.hasOwnProperty(weaponName)) {
          var oldWeapon = oldUltimate[weaponName];
          newUltimateUnlocked.push([weaponName]);
          newUltimateUnlocked.push([null]);
          newUltimateUnlocked.push([oldWeapon.unlocked]);

          for (var nextRow = row; nextRow < newUltimateData.length; nextRow++) {
            var nextRowData = newUltimateData[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 2;
              break;
            }
            var newWeaponAttribute = nextRowData[2];
            var newValues = [];

            if (
              oldWeapon.hasOwnProperty("levels") &&
              newWeaponAttribute &&
              oldWeapon.levels.hasOwnProperty(newWeaponAttribute)
            ) {
              var dvtLevelValue = shared.getDVTValue(
                oldWeapon.levels[newWeaponAttribute],
                dvtNamedRangesData[weaponName][newWeaponAttribute],
              );
              newValues.push(dvtLevelValue);
            } else {
              newValues.push(null);
            }

            if (
              oldWeapon.hasOwnProperty("targets") &&
              newWeaponAttribute &&
              oldWeapon.targets.hasOwnProperty(newWeaponAttribute)
            ) {
              var dvtTargetValue = shared.getDVTValue(
                oldWeapon.targets[newWeaponAttribute],
                dvtNamedRangesData[weaponName][newWeaponAttribute],
              );
              newValues.push(dvtTargetValue);
            } else {
              newValues.push(null);
            }
            newUltimateLevel.push(newValues);

            if (nextRow == newUltimateData.length - 1) {
              row = nextRow;
            }
          }
        } else {
          newUltimateUnlocked.push([null]);
        }
      }

      var batchUpdate = [];

      if (newUltimateUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(ultimateCol + 1);
        var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
          newUltimateUnlocked.length + 1
        }`;
        batchUpdate.push({
          range: unlockedRange,
          values: newUltimateUnlocked,
        });
      }

      if (newUltimateLevel.length > 0) {
        var levelCol = shared.columnToLetter(ultimateCol + 5);
        var targetCol = shared.columnToLetter(ultimateCol + 6);
        var levelRange = `${sheetName}!${levelCol}2:${targetCol}${
          newUltimateLevel.length + 1
        }`;
        batchUpdate.push({
          range: levelRange,
          values: newUltimateLevel,
        });
      }

      if (batchUpdate.length !== 0) {

        return {
          success: true,
          message: `Ultimate weapons levels updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for ultimate weapons levels`,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.updateUltimateLevels", error, {
        sheetName: sheetName,
        oldUltimate: oldUltimate,
        masterSheetData: masterSheetData,
        dvtNamedRangesData: dvtNamedRangesData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes UltimateCostCalculator into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldUltimateCostCalculator
   * @param {Object} ultimateCostCalculatorData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateUltimateCostCalculator: function (
    sheetName,
    oldUltimateCostCalculator,
    ultimateCostCalculatorData,
  ) {
    try {
      console.log("Called: ultimate.updateUltimateCostCalculator");
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

      if (
        !ultimateCostCalculatorData ||
        ultimateCostCalculatorData.length === 0
      ) {
        console.log(`No data in UW Cost Calculator sheet`);
        return {
          success: false,
          message: `No data in UW Cost Calculator sheet`,
        };
      }

      var batchUpdate = [];
      var missingWeapons = [...targetWeapons];

      for (var row = 0; row < ultimateCostCalculatorData.length; row++) {

        if (missingWeapons.length === 0) {
          break;
        }

        var rowData = ultimateCostCalculatorData[row];

        var uwsWantedIndex = rowData.indexOf("# Of UWs Wanted");
        var uwPlusWantedIndex = rowData.indexOf("# Of UW+ Wanted");
        if (
          uwsWantedIndex !== -1 &&
          oldUltimateCostCalculator["# Of UWs Wanted"]
        ) {
          batchUpdate.push({
            range: `${sheetName}!${shared.columnToLetter(uwsWantedIndex + 3)}${
              row + 1
            }`,
            values: [[oldUltimateCostCalculator["# Of UWs Wanted"]]],
          });
        }
        if (
          uwPlusWantedIndex !== -1 &&
          oldUltimateCostCalculator["# Of UW+ Wanted"]
        ) {
          batchUpdate.push({
            range: `${sheetName}!${shared.columnToLetter(
              uwPlusWantedIndex + 3,
            )}${row + 1}`,
            values: [[oldUltimateCostCalculator["# Of UW+ Wanted"]]],
          });
        }

        for (
          var weaponIndex = 0;
          weaponIndex < missingWeapons.length;
          weaponIndex++
        ) {
          var weapon = missingWeapons[weaponIndex];

          if (
            rowData.includes(weapon) &&
            oldUltimateCostCalculator.hasOwnProperty(weapon)
          ) {
            var weaponColIndex = rowData.indexOf(weapon);
            var oldWeaponData = oldUltimateCostCalculator[weapon];

            if (oldWeaponData.hasOwnProperty("unlocked")) {
              var unlockedCol = shared.columnToLetter(weaponColIndex + 4);
              var unlockedRange = `${sheetName}!${unlockedCol}${row + 1}`;
              batchUpdate.push({
                range: unlockedRange,
                values: [[oldWeaponData.unlocked]],
              });
            }

            if (
              oldWeaponData.hasOwnProperty("values") &&
              Object.keys(oldWeaponData.values).length > 0
            ) {

              if (row + 1 < ultimateCostCalculatorData.length) {
                var subData = ultimateCostCalculatorData[row + 1];
                var currentValueIndex = subData.indexOf("Current Value");
                var subNameIndex = currentValueIndex - 1;
                var targetValueIndex = subData.indexOf("Target Value");
                var modSubIndex = subData.indexOf("Sub");

                for (
                  var subRow = row + 2;
                  subRow < ultimateCostCalculatorData.length;
                  subRow++
                ) {
                  var subRowData = ultimateCostCalculatorData[subRow];
                  var subName = subRowData[subNameIndex]
                    ? subRowData[subNameIndex].toString().trim()
                    : "";

                  if (
                    subName === "" ||
                    targetWeapons.includes(
                      subRowData[weaponColIndex]
                        ? subRowData[weaponColIndex].toString().trim()
                        : "",
                    )
                  ) {
                    break;
                  }

                  if (oldWeaponData.values.hasOwnProperty(subName)) {
                    var subValues = oldWeaponData.values[subName];

                    if (subValues.hasOwnProperty("currentValue")) {
                      var currentCol = shared.columnToLetter(
                        currentValueIndex + 1,
                      );
                      var currentRange = `${sheetName}!${currentCol}${
                        subRow + 1
                      }`;
                      batchUpdate.push({
                        range: currentRange,
                        values: [[subValues.currentValue]],
                      });
                    }

                    if (subValues.hasOwnProperty("targetValue")) {
                      var targetCol = shared.columnToLetter(
                        targetValueIndex + 1,
                      );
                      var targetRange = `${sheetName}!${targetCol}${
                        subRow + 1
                      }`;
                      batchUpdate.push({
                        range: targetRange,
                        values: [[subValues.targetValue]],
                      });
                    }

                    if (subValues.hasOwnProperty("modSub")) {
                      var modCol = shared.columnToLetter(modSubIndex + 1);
                      var modRange = `${sheetName}!${modCol}${subRow + 1}`;
                      batchUpdate.push({
                        range: modRange,
                        values: [[subValues.modSub]],
                      });
                    }
                  }

                  row = subRow;
                }
              }
            }

            missingWeapons.splice(weaponIndex, 1);

            break;
          }
        }
      }

      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: `Ultimate cost calculator updated successfully`,
          batchUpdate: batchUpdate,
        };
      } else {
        return {
          success: true,
          message: `No updates needed for ultimate cost calculator`,
        };
      }
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.updateUltimateCostCalculator", error, {
        sheetName: sheetName,
        oldUltimateCostCalculator: oldUltimateCostCalculator,
        ultimateCostCalculatorData: ultimateCostCalculatorData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Ultimate_Weapons data from a v3.1.1 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_1_1: function (oldSheetID) {
    try {
      console.log("Called: ultimate.version3_1_1");

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
      console.log("Called: ultimate.version2_0");

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
      console.log("Called: ultimate.version1_0");

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
      console.log("Called: ultimate.getVersion3_1_1UltimateWeapons");
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
      console.log("Called: ultimate.getVersion2_0UltimateWeapons");
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
      console.log("Called: ultimate.getVersion1_0UltimateWeapons");
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
      console.log("Called: ultimate.getVersion1_0CostCalculator");
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

  /**
   * Parses Ultimate_Weapons data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseUltimateWeaponData: function (data) {
    try {
      const targetWeapons = {
        "Chain Lightning": { upgrades: ["Damage", "Quantity", "Chance"], plusUpgrade: "Smite" },
        "Smart Missiles": { upgrades: ["Damage", "Quantity", "Cooldown"], plusUpgrade: "Cover Fire" },
        "Death Wave": { upgrades: ["Damage", "Quantity", "Cooldown"], plusUpgrade: "Kill Wall" },
        "Chrono Field": { upgrades: ["Duration", "Speed Reduction", "Cooldown"], plusUpgrade: "Chrono Loop" },
        "Inner Land Mines": { upgrades: ["Damage", "Quantity", "Cooldown"], plusUpgrade: "Charged Mines" },
        "Golden Tower": { upgrades: ["Multiplier", "Duration", "Cooldown"], plusUpgrade: "Golden Combo" },
        "Poison Swamp": { upgrades: ["Damage", "Duration", "Cooldown"], plusUpgrade: "Death Creep" },
        "Black Hole": { upgrades: ["Size", "Duration", "Cooldown"], plusUpgrade: "Consume" },
        "Spotlight": { upgrades: ["Multiplier", "Angle", "Quantity"], plusUpgrade: "Light Range" },
      };
      const ultimateWeaponUnlocked = data.ultimateWeaponUnlocked || [];
      const ultimateWeaponLevel = data.ultimateWeaponLevel || [];
      const ultimateWeaponPlusUnlocked = data.ultimateWeaponPlusUnlocked || [];
      const ultimateWeaponPlusLevel = data.ultimateWeaponPlusLevel || [];

      var oldUltimate = {};

      Object.keys(targetWeapons).forEach(function (weaponName, i) {
        var { upgrades, plusUpgrade } = targetWeapons[weaponName];
        var weaponLevels = {};
        upgrades.forEach(function (attr, j) {
          var idx = i * upgrades.length + j;
          var level = ultimateWeaponLevel[idx];
          weaponLevels[attr] = level ? String(level).padStart(2, "0") : "00";
        });
        if (!ultimateWeaponPlusUnlocked[i]) {
          weaponLevels[plusUpgrade] = "Lo";
        } else {
          var level = ultimateWeaponPlusLevel[i];
          weaponLevels[plusUpgrade] = level ? String(level).padStart(2, "0") : "00";
        }
        oldUltimate[weaponName] = {
          unlocked: ultimateWeaponUnlocked[i] || null,
          levels: weaponLevels,
        };
      });

      return {
        success: true,
        oldUltimate: oldUltimate,
        targetWeapons: targetWeapons,
        weaponOrder: Object.keys(targetWeapons),
      };
    } catch (error) {
      var errorReport = errors.report("ultimate.parseUltimateWeaponData", error, {
        data: data,
        oldUltimate: oldUltimate,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v2.0": this.version2_0.bind(this),
      "v3.1.1": this.version3_1_1.bind(this),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
  isCompatibleVersion: function (oldVersion) {
    console.log("Called: ultimate.isCompatibleVersion");
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
