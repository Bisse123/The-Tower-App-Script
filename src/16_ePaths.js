const ePaths = {

  /**
   * Reads EPaths data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: ePaths.exportData");
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
        message: "Effective Paths export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported EPaths data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: ePaths.importData");

      var eHPRange = "eHP!AJ1:AY50";
      var eDamageRange = "eDamage!AI1:AY100";
      var eEconRange = "eEcon!AK1:AZ65";

      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!O3:O5";
      var eEconStoneMult = "eEcon!X3:X5";
      var eDiscountLabRange = "eEcon!AH3:AH5";

      var eHPLabColumn = "L";
      var eRegenLabColumn = "AH";
      var eDamageLabColumn = "L";
      var eEconLabColumn = "O";
      var eEconStoneMultColumn = "X";
      var eDiscountLabColumn = "AH";

      var ranges = [
        "IDS",
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eEconStoneMult,
        eDiscountLabRange,
      ];

      var eHPColumnOffset = shared.getColumnOffsetFromRange(eHPRange);
      var eDamageColumnOffset = shared.getColumnOffsetFromRange(eDamageRange);
      var eEconColumnOffset = shared.getColumnOffsetFromRange(eEconRange);

      var batchGetResult = SheetsAPI.batchGetFormulas(newSheetID, ranges);
      if (!batchGetResult || batchGetResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from new spreadsheet™.",
        };
      }

      var idsData = batchGetResult[0].values;
      var eHPValues = batchGetResult[1].values;
      var eDamageValues = batchGetResult[2].values;
      var eEconValues = batchGetResult[3].values;
      var eHPLabValues = batchGetResult[4].values;
      var eRegenLabValues = batchGetResult[5].values;
      var eDamageLabValues = batchGetResult[6].values;
      var eEconLabValues = batchGetResult[7].values;
      var eEconStoneMultValues = batchGetResult[8].values;
      var eDiscountLabValues = batchGetResult[9].values;

      var batchUpdate = [];

      if (data.hasOwnProperty("eHP")) {
        var eHPResult = this.updateEHP(
          "eHP",
          data.eHP.oldData,
          eHPValues,
          eHPColumnOffset,
          eHPLabValues,
          eRegenLabValues,
          eHPLabColumn,
          eRegenLabColumn,
        );
        if (!eHPResult || !eHPResult.success) {
          console.log(`eHP update failed: ${eHPResult.message}`);
          return {
            success: false,
            message: `eHP update failed: ${eHPResult.message}`,
          };
        }
        batchUpdate = batchUpdate.concat(eHPResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("eDamage")) {
        var eDamageResult = this.updateEDamage(
          "eDamage",
          data.eDamage.oldData,
          eDamageValues,
          eDamageColumnOffset,
          eDamageLabValues,
          eDamageLabColumn,
        );
        if (!eDamageResult || !eDamageResult.success) {
          console.log(`eDamage update failed: ${eDamageResult.message}`);
          return {
            success: false,
            message: `eDamage update failed: ${eDamageResult.message}`,
          };
        }
        batchUpdate = batchUpdate.concat(eDamageResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("eEcon")) {
        var eEconResult = this.updateEEcon(
          "eEcon",
          data.eEcon.oldData,
          eEconValues,
          eEconColumnOffset,
          eEconLabValues,
          eEconStoneMultValues,
          eDiscountLabValues,
          eEconLabColumn,
          eEconStoneMultColumn,
          eDiscountLabColumn,
        );
        if (!eEconResult || !eEconResult.success) {
          console.log(`eEcon update failed: ${eEconResult.message}`);
          return {
            success: false,
            message: `eEcon update failed: ${eEconResult.message}`,
          };
        }
        batchUpdate = batchUpdate.concat(eEconResult.batchUpdate || []);
      }

      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Effective Paths",
        newSheetID,
        idsData,
        data.idMasterID,
      );

      var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
      if (!updateResult) {
        return {
          success: false,
          message:
            "Failed to update new spreadsheet™ with Effective Paths data.",
        };
      }
      return {
        success: true,
        message: `Effective Paths import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.importData", error, {
        note: `Error importing ePaths data`,
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes EHP into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldData
   * @param {Object} eHPData
   * @param {*} columnOffset
   * @param {Object} eHPLabData
   * @param {Object} eRegenLabData
   * @param {*} eHPLabColumn
   * @param {*} eRegenLabColumn
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateEHP: function (
    sheetName,
    oldData,
    eHPData,
    columnOffset,
    eHPLabData,
    eRegenLabData,
    eHPLabColumn,
    eRegenLabColumn,
  ) {
    try {
      console.log("Called: ePaths.updateEHP");

      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }

      var batchUpdate = [];

      if (eHPLabData && eHPLabData[1] && eHPLabData[1][0] === "Running Time") {
        if (oldData.hasOwnProperty("eHPLabCost")) {
          var eHPCostValue = oldData.eHPLabCost;
          batchUpdate.push({
            range: `${sheetName}!${eHPLabColumn}3`,
            values: [[eHPCostValue]],
          });
        }
        if (oldData.hasOwnProperty("eHPRunningTime")) {
          var eHPRunningTimeValue = oldData.eHPRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eHPLabColumn}5`,
            values: [[eHPRunningTimeValue]],
          });
        }
      }
      if (
        eRegenLabData &&
        eRegenLabData[1] &&
        eRegenLabData[1][0] === "Running Time"
      ) {
        if (oldData.hasOwnProperty("eRegenLabCost")) {
          var eRegenLabCostValue = oldData.eRegenLabCost;
          batchUpdate.push({
            range: `${sheetName}!${eRegenLabColumn}3`,
            values: [[eRegenLabCostValue]],
          });
        }
        if (oldData.hasOwnProperty("eRegenRunningTime")) {
          var eRegenRunningTimeValue = oldData.eRegenRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eRegenLabColumn}5`,
            values: [[eRegenRunningTimeValue]],
          });
        }
      }

      for (var row = 0; row < eHPData.length; row++) {
        for (var column = 0; column < eHPData[row].length; column++) {
          var cell = eHPData[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < eHPData.length; nextRow++) {
              var customName = eHPData[nextRow][column - 2];
              if (!customName) break;
              if (oldData.Custom && oldData.Custom.hasOwnProperty(customName)) {
                var newValue = oldData.Custom[customName];
                var cellCol = shared.columnToLetter(
                  columnOffset + (column - 1) + 1,
                );
                var cellAddress = `${cellCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${cellAddress}`,
                  values: [[newValue]],
                });
              }
            }

          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < eHPData.length; nextRow++) {
              var guessName = eHPData[nextRow][column];
              if (!guessName) break;
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (
                oldData.UserGuess &&
                oldData.UserGuess.hasOwnProperty(guessName)
              ) {
                var guessValue = oldData.UserGuess[guessName];
                var guessCol = shared.columnToLetter(columnOffset + column + 6);
                var guessCellAddress = `${guessCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${guessCellAddress}`,
                  values: [[guessValue]],
                });
              }
            }
          } else if (oldData.Modules && oldData.Modules.hasOwnProperty(cell)) {
            var moduleValue = oldData.Modules[cell];
            var moduleCol = shared.columnToLetter(columnOffset + column + 2);
            var moduleCellAddress = `${moduleCol}${row + 1}`;
            batchUpdate.push({
              range: `${sheetName}!${moduleCellAddress}`,
              values: [[moduleValue]],
            });
          } else if (cell === "Rows Calculated") {
            if (oldData.hasOwnProperty("rowsCalculated")) {
              var rowsCalculatedValue = oldData.rowsCalculated;
              var rowsCol = shared.columnToLetter(columnOffset + column + 1);
              var rowsCalculatedCellAddress = `${rowsCol}${row + 2}`;
              batchUpdate.push({
                range: `${sheetName}!${rowsCalculatedCellAddress}`,
                values: [[rowsCalculatedValue]],
              });
            }

          }
        }
      }
      return {
        success: true,
        message: "eHP data updated successfully",
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.updateEHP`,
        sheetName: sheetName,
        oldData: oldData,
        eHPData: eHPData,
        columnOffset: columnOffset,
        eHPLabData: eHPLabData,
        eRegenLabData: eRegenLabData,
        eHPLabColumn: eHPLabColumn,
        eRegenLabColumn: eRegenLabColumn,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes EDamage into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldData
   * @param {Object} eDamageData
   * @param {*} columnOffset
   * @param {Object} eDamageLabData
   * @param {*} eDamageLabColumn
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateEDamage: function (
    sheetName,
    oldData,
    eDamageData,
    columnOffset,
    eDamageLabData,
    eDamageLabColumn,
  ) {
    try {
      console.log("Called: ePaths.updateEDamage");

      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }

      var batchUpdate = [];

      if (
        eDamageLabData &&
        eDamageLabData[1] &&
        eDamageLabData[1][0] === "Running Time"
      ) {
        if (oldData.hasOwnProperty("eDamageLabCost")) {
          var eDamageCostValue = oldData.eDamageLabCost;
          batchUpdate.push({
            range: `${sheetName}!${eDamageLabColumn}3`,
            values: [[eDamageCostValue]],
          });
        }
        if (oldData.hasOwnProperty("eDamageRunningTime")) {
          var eDamageRunningTimeValue = oldData.eDamageRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eDamageLabColumn}5`,
            values: [[eDamageRunningTimeValue]],
          });
        }
      }

      for (var row = 0; row < eDamageData.length; row++) {
        for (var column = 0; column < eDamageData[row].length; column++) {
          var cell = eDamageData[row][column];
          if (cell === "Total Value") {

            for (
              var nextRow = row + 1;
              nextRow < eDamageData.length;
              nextRow++
            ) {
              var customName = eDamageData[nextRow][column - 2];
              if (!customName) break;
              if (oldData.Custom && oldData.Custom.hasOwnProperty(customName)) {
                var newValue = oldData.Custom[customName];
                var cellCol = shared.columnToLetter(
                  columnOffset + (column - 1) + 1,
                );
                var cellAddress = `${cellCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${cellAddress}`,
                  values: [[newValue]],
                });
              }
            }

          } else if (cell === "User Specific Guesses") {
            for (
              var nextRow = row + 1;
              nextRow < eDamageData.length;
              nextRow++
            ) {
              var guessName = eDamageData[nextRow][column];
              if (!guessName) break;
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (
                oldData.UserGuess &&
                oldData.UserGuess.hasOwnProperty(guessName)
              ) {
                var guessValue = oldData.UserGuess[guessName];
                var guessValueIndex = 6;
                if (["Run Type", "Simulated Tier"].includes(guessName)) {
                  guessValueIndex -= 1;
                }
                var guessCol = shared.columnToLetter(
                  columnOffset + column + guessValueIndex,
                );
                var guessCellAddress = `${guessCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${guessCellAddress}`,
                  values: [[guessValue]],
                });
              }
            }
          } else if (oldData.Modules && oldData.Modules.hasOwnProperty(cell)) {
            var moduleValue = oldData.Modules[cell];
            var moduleCol = shared.columnToLetter(columnOffset + column + 2);
            var moduleCellAddress = `${moduleCol}${row + 1}`;
            batchUpdate.push({
              range: `${sheetName}!${moduleCellAddress}`,
              values: [[moduleValue]],
            });
          } else if (cell === "Rows Calculated") {
            if (oldData.hasOwnProperty("rowsCalculated")) {
              var rowsCalculatedValue = oldData.rowsCalculated;
              var rowsCol = shared.columnToLetter(columnOffset + column + 1);
              var rowsCalculatedCellAddress = `${rowsCol}${row + 2}`;
              batchUpdate.push({
                range: `${sheetName}!${rowsCalculatedCellAddress}`,
                values: [[rowsCalculatedValue]],
              });
            }

          } else if (cell === "PS Beta Testing") {
            if (oldData.hasOwnProperty("PSBeta")) {
              var psBetaValue = oldData.PSBeta;
              var psBetaCol = shared.columnToLetter(columnOffset + column + 1);
              var psBetaCellAddress = `${psBetaCol}${row}`;
              batchUpdate.push({
                range: `${sheetName}!${psBetaCellAddress}`,
                values: [[psBetaValue]],
              });
            }
          }
        }
      }
      return {
        success: true,
        message: "eDamage data updated successfully",
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.updateEDamage`,
        sheetName: sheetName,
        oldData: oldData,
        eDamageData: eDamageData,
        columnOffset: columnOffset,
        eDamageLabData: eDamageLabData,
        eDamageLabColumn: eDamageLabColumn,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes EEcon into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldData
   * @param {Object} eEconData
   * @param {*} columnOffset
   * @param {Object} eEconLabData
   * @param {Object} eEconStoneMultData
   * @param {Object} eDiscountLabData
   * @param {*} eEconLabColumn
   * @param {*} eEconStoneMultColumn
   * @param {*} eDiscountLabColumn
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateEEcon: function (
    sheetName,
    oldData,
    eEconData,
    columnOffset,
    eEconLabData,
    eEconStoneMultData,
    eDiscountLabData,
    eEconLabColumn,
    eEconStoneMultColumn,
    eDiscountLabColumn,
  ) {
    try {
      console.log("Called: ePaths.updateEEcon");

      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }

      var batchUpdate = [];

      if (
        eEconLabData &&
        eEconLabData[1] &&
        eEconLabData[1][0] === "Speed Up"
      ) {
        if (oldData.hasOwnProperty("eEconLabRoi")) {
          var eEconCostValue = oldData.eEconLabRoi;
          batchUpdate.push({
            range: `${sheetName}!${eEconLabColumn}3`,
            values: [[eEconCostValue]],
          });
        }
        if (oldData.hasOwnProperty("eEconRunningTime")) {
          var eEconRunningTimeValue = oldData.eEconRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eEconLabColumn}5`,
            values: [[eEconRunningTimeValue]],
          });
        }
      }
      if (
        eEconStoneMultData &&
        eEconStoneMultData[1] &&
        eEconStoneMultData[1][0] === "Stone Multiplier"
      ) {
        if (oldData.hasOwnProperty("eEconStoneMult")) {
          var eEconStoneMultValue = oldData.eEconStoneMult;
          batchUpdate.push({
            range: `${sheetName}!${eEconStoneMultColumn}5`,
            values: [[eEconStoneMultValue]],
          });
        }
      }
      if (
        eDiscountLabData &&
        eDiscountLabData[1] &&
        eDiscountLabData[1][0] === "Speed Up"
      ) {
        if (oldData.hasOwnProperty("eDiscountLabRoi")) {
          var eDiscountLabRoi = oldData.eDiscountLabRoi;
          batchUpdate.push({
            range: `${sheetName}!${eDiscountLabColumn}3`,
            values: [[eDiscountLabRoi]],
          });
        }
        if (oldData.hasOwnProperty("eDiscountRunningTime")) {
          var eDiscountRunningTime = oldData.eDiscountRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eDiscountLabColumn}5`,
            values: [[eDiscountRunningTime]],
          });
        }
      }

      for (var row = 0; row < eEconData.length; row++) {
        for (var column = 0; column < eEconData[row].length; column++) {
          var cell = eEconData[row][column];

          if (cell === "User Inputs") {
            var skipPresets = false;
            for (var nextRow = row + 1; nextRow < eEconData.length; nextRow++) {
              var guessName = eEconData[nextRow][column];
              console.log(`Processing guessName: ${guessName}`);
              if (!guessName) break;
              if (String(guessName).toLowerCase().includes("presets")) {
                skipPresets = true;
              } else if (String(guessName).toLowerCase().includes("inputs")) {
                skipPresets = false;
              }
              if (skipPresets) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (
                oldData.UserGuess &&
                oldData.UserGuess.hasOwnProperty(guessName)
              ) {
                var guessValue = oldData.UserGuess[guessName];
                var guessCol = shared.columnToLetter(columnOffset + column + 6);
                var guessCellAddress = `${guessCol}${nextRow + 1}`;
                if (guessName === "GB Sync Desired Ratio:") {
                  var gbFirstCol = shared.columnToLetter(
                    columnOffset + column + 4,
                  );
                  guessCellAddress = `${gbFirstCol}${nextRow + 1}:${guessCol}${nextRow + 1}`;
                  var gbValue = [
                    guessValue["antecedent value"],
                    null,
                    guessValue["consequent value"],
                  ];
                  batchUpdate.push({
                    range: `${sheetName}!${guessCellAddress}`,
                    values: [gbValue],
                  });
                  continue;
                }
                batchUpdate.push({
                  range: `${sheetName}!${guessCellAddress}`,
                  values: [[guessValue]],
                });
              }
            }
          } else if (oldData.Modules && oldData.Modules.hasOwnProperty(cell)) {
            var moduleValue = oldData.Modules[cell];
            if (moduleValue.main !== undefined) {
              var moduleCol = shared.columnToLetter(columnOffset + column + 2);
              var moduleCellAddress = `${moduleCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${moduleCellAddress}`,
                values: [[moduleValue.main]],
              });
            }
            if (moduleValue.assist !== undefined) {
              var assistCol = shared.columnToLetter(columnOffset + column + 6);
              var assistCellAddress = `${assistCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${assistCellAddress}`,
                values: [[moduleValue.assist]],
              });
            }
          } else if (cell === "Calculation Rows") {
            if (oldData.hasOwnProperty("rowsCalculated")) {
              var rowsCalculatedValue = oldData.rowsCalculated;
              var rowsCol = shared.columnToLetter(columnOffset + column + 1);
              var rowsCalculatedCellAddress = `${rowsCol}${row + 2}`;
              batchUpdate.push({
                range: `${sheetName}!${rowsCalculatedCellAddress}`,
                values: [[rowsCalculatedValue]],
              });
            }
          } else if (cell === "Coins Spent") {
            if (oldData.hasOwnProperty("enhancementDiscount")) {
              var enhancementDiscountValue = oldData.enhancementDiscount;
              var enhancementDiscountCol = shared.columnToLetter(
                columnOffset + column,
              );
              var enhancementDiscountCellAddress = `${enhancementDiscountCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${enhancementDiscountCellAddress}`,
                values: [[enhancementDiscountValue]],
              });
            }
          }
        }
      }
      return {
        success: true,
        message: "eEcon data updated successfully",
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.updateEEcon`,
        sheetName: sheetName,
        oldData: oldData,
        eEconData: eEconData,
        columnOffset: columnOffset,
        eEconLabData: eEconLabData,
        eEconStoneMultData: eEconStoneMultData,
        eDiscountLabData: eDiscountLabData,
        eEconLabColumn: eEconLabColumn,
        eEconStoneMultColumn: eEconStoneMultColumn,
        eDiscountLabColumn: eDiscountLabColumn,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.09.00.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_09_00_00: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version5_09_00_00");

      var eHPRange = "eHP!AJ1:AY50";
      var eDamageRange = "eDamage!AI1:AY100";
      var eEconRange = "eEcon!AK1:AZ65";

      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!O3:O5";
      var eEconStoneMult = "eEcon!X3:X5";
      var eDiscountLabRange = "eEcon!AH3:AH5";

      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eEconStoneMult,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eEconStoneMultValues = batchResult[7].values;
      var eDiscountLabValues = batchResult[8].values;

      var eHPData = this.getVersion5_09_00_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion5_09_00_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = this.getVersion5_09_00_00eEcon(
        eEconValues,
        eEconLabValues,
        eEconStoneMultValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version5_09_00_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.08.04.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_08_04_00: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version5_08_04_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AK1:AY65";

      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!O3:O5";
      var eEconStoneMult = "eEcon!X3:X5";
      var eDiscountLabRange = "eEcon!AH3:AH5";

      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eEconStoneMult,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eEconStoneMultValues = batchResult[7].values;
      var eDiscountLabValues = batchResult[8].values;

      var eHPData = this.getVersion5_05_01_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion5_06_02_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = this.getVersion5_08_00_00eEcon(
        eEconValues,
        eEconLabValues,
        eEconStoneMultValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version5_08_04_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.08.00.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_08_00_00: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version5_08_00_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AS1:BG65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!O3:O5";
      var eEconStoneMult = "eEcon!X3:X5";
      var eDiscountLabRange = "eEcon!AQ3:AQ5";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eEconStoneMult,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eEconStoneMultValues = batchResult[7].values;
      var eDiscountLabValues = batchResult[8].values;

      var eHPData = this.getVersion5_05_01_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion5_06_02_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = this.getVersion5_08_00_00eEcon(
        eEconValues,
        eEconLabValues,
        eEconStoneMultValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version5_08_00_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.06.02.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_06_02_00: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version5_06_02_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AI1:AW65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;

      var eHPData = this.getVersion5_05_01_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion5_06_02_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = this.getVersion5_06_02_00eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version5_06_02_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.05.01.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_05_01_00: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version5_05_01_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AI1:AW65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = this.getVersion5_05_01_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion5_05_00_00eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = this.getVersion5_00_01_04eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version5_05_01_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.05.00.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_05_00_00: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version5_05_00_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AI1:AW65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = this.getVersion5_03_00_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion5_05_00_00eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = this.getVersion5_00_01_04eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version5_05_00_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.03.00.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_03_00_00: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version5_03_00_00");

      var eHPRange = "eHP!AJ1:AX50";
      var eDamageRange = "eDamage!AI1:AX100";
      var eEconRange = "eEcon!AI1:AW65";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = this.getVersion5_03_00_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion4_11_03_21eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = this.getVersion5_00_01_04eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version5_03_00_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v5.00.01.04 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_00_01_04: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version5_00_01_04");

      var eHPRange = "eHP!AC1:AQ35";
      var eDamageRange = "eDamage!AI1:AX90";
      var eEconRange = "eEcon!AI1:AW55";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AA3:AA5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = this.getVersion4_11_03_21eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion4_11_03_21eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = this.getVersion5_00_01_04eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version5_00_01_04`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v4.11.03.21 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_11_03_21: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version4_11_03_21");

      var eHPRange = "eHP!AC1:AQ35";
      var eDamageRange = "eDamage!AI1:AX75";
      var eEconRange = "eEcon!AI1:AW55";
      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AA3:AA5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!L3:L5";
      var eDiscountLabRange = "eEcon!AG3:AG5";
      var CLDmgRange = "eDamage!AL149:AM149";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
        CLDmgRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;
      var cLDmgValues = batchResult[8].values;

      var eHPData = this.getVersion4_11_03_21eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion4_11_03_21eDamage(
        eDamageValues,
        eDamageLabValues,
        cLDmgValues,
      );
      var eEconData = this.getVersion4_11_03_21eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version4_11_03_21`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads EPaths data from a v4.11.02.00 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_11_02_00: function (oldSheetID) {
    try {
      console.log("Called: ePaths.version4_11_02_00");

      var eHPRange = "eHP!AC1:AQ35";
      var eDamageRange = "eDamage!AI1:AX75";
      var eEconRange = "eEcon!AI1:AW55";
      var eHPLabRange = "eHP!L4:L5";
      var eRegenLabRange = "eHP!AA4:AA5";
      var eDamageLabRange = "eDamage!L4:L5";
      var eEconLabRange = "eEcon!L4:L5";
      var eDiscountLabRange = "eEcon!AG4:AG5";
      var ranges = [
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eDiscountLabRange,
      ];
      var batchResult = SheetsAPI.batchGetFormulas(oldSheetID, ranges);
      if (!batchResult || !batchResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from old spreadsheet™.",
        };
      }
      var eHPValues = batchResult[0].values;
      var eDamageValues = batchResult[1].values;
      var eEconValues = batchResult[2].values;
      var eHPLabValues = batchResult[3].values;
      var eRegenLabValues = batchResult[4].values;
      var eDamageLabValues = batchResult[5].values;
      var eEconLabValues = batchResult[6].values;
      var eDiscountLabValues = batchResult[7].values;

      var eHPData = this.getVersion4_11_02_00eHP(
        eHPValues,
        eHPLabValues,
        eRegenLabValues,
      );
      var eDamageData = this.getVersion4_11_02_00eDamage(
        eDamageValues,
        eDamageLabValues,
      );
      var eEconData = this.getVersion4_11_02_00eEcon(
        eEconValues,
        eEconLabValues,
        eDiscountLabValues,
      );

      return {
        success: true,
        eHP: eHPData,
        eDamage: eDamageData,
        eEcon: eEconData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePaths.version4_11_02_00`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v5.09.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_09_00_00eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion5_09_00_00eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[1] &&
        oldeHPLabValues[1][0] === "Running Time"
      ) {
        var eHPLabCost = oldeHPLabValues[0][0];
        if (String(eHPLabCost)) {
          oldData.eHPLabCost = eHPLabCost;
        }
        var eHPRunningTime = oldeHPLabValues[2][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[1] &&
        oldeRegenLabValues[1][0] === "Running Time"
      ) {
        var eRegenLabCost = oldeRegenLabValues[0][0];
        if (String(eRegenLabCost)) {
          oldData.eRegenLabCost = eRegenLabCost;
        }
        var eRegenRunningTime = oldeRegenLabValues[2][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 5];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 5];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 5];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Presets") {
            var globalPreset = oldValues[row][column + 4];
            oldData.Presets = {
              Presets: globalPreset,
            };
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              var presetValue = oldValues[nextRow][column + 4];
              if (!presetName) break;
              if (presetValue === globalPreset) {
                continue;
              }
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = presetValue;
            }
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_09_00_00eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v5.05.01.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_05_01_00eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion5_05_01_00eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[1] &&
        oldeHPLabValues[1][0] === "Running Time"
      ) {
        var eHPLabCost = oldeHPLabValues[0][0];
        if (String(eHPLabCost)) {
          oldData.eHPLabCost = eHPLabCost;
        }
        var eHPRunningTime = oldeHPLabValues[2][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[1] &&
        oldeRegenLabValues[1][0] === "Running Time"
      ) {
        var eRegenLabCost = oldeRegenLabValues[0][0];
        if (String(eRegenLabCost)) {
          oldData.eRegenLabCost = eRegenLabCost;
        }
        var eRegenRunningTime = oldeRegenLabValues[2][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Presets") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              if (!presetName) break;
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = oldValues[nextRow][column + 3];
            }
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_05_01_00eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v5.03.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_03_00_00eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion5_03_00_00eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[1] &&
        oldeHPLabValues[1][0] === "Running Time"
      ) {
        var eHPLabCost = oldeHPLabValues[0][0];
        if (String(eHPLabCost)) {
          oldData.eHPLabCost = eHPLabCost;
        }
        var eHPRunningTime = oldeHPLabValues[2][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[1] &&
        oldeRegenLabValues[1][0] === "Running Time"
      ) {
        var eRegenLabCost = oldeRegenLabValues[0][0];
        if (String(eRegenLabCost)) {
          oldData.eRegenLabCost = eRegenLabCost;
        }
        var eRegenRunningTime = oldeRegenLabValues[2][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Total Lab Time") {
            var preset = oldValues[row][column + 2];
            oldData.Presets = {
              Workshop: preset,
              Cards: preset,
              Modules: preset,
            };
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_03_00_00eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v4.11.03.21 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_03_21eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion4_11_03_21eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[1] &&
        oldeHPLabValues[1][0] === "Running Time"
      ) {
        var eHPLabCost = oldeHPLabValues[0][0];
        if (String(eHPLabCost)) {
          oldData.eHPLabCost = eHPLabCost;
        }
        var eHPRunningTime = oldeHPLabValues[2][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[1] &&
        oldeRegenLabValues[1][0] === "Running Time"
      ) {
        var eRegenLabCost = oldeRegenLabValues[0][0];
        if (String(eRegenLabCost)) {
          oldData.eRegenLabCost = eRegenLabCost;
        }
        var eRegenRunningTime = oldeRegenLabValues[2][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion4_11_03_21eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v4.11.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_02_00eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion4_11_02_00eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[0] &&
        oldeHPLabValues[0][0] === "Running Time"
      ) {
        var eHPRunningTime = oldeHPLabValues[1][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[0] &&
        oldeRegenLabValues[0][0] === "Running Time"
      ) {
        var eRegenRunningTime = oldeRegenLabValues[1][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion4_11_02_00eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v5.09.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_09_00_00eDamage: function (oldValues, oldeDamageLabValues) {
    try {
      console.log("Called: ePaths.getVersion5_09_00_00eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[1] &&
        oldeDamageLabValues[1][0] === "Running Time"
      ) {
        var eDamageLabCost = oldeDamageLabValues[0][0];
        if (String(eDamageLabCost)) {
          oldData.eDamageLabCost = eDamageLabCost;
        }
        var eDamageRunningTime = oldeDamageLabValues[2][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 5];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 5];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValueIndex = 5;
              if (["Run Type", "Simulated Tier"].includes(guessName)) {
                guessValueIndex -= 1;
              }
              var guessValue = oldValues[nextRow][column + guessValueIndex];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Presets") {
            var globalPreset = oldValues[row][column + 4];
            oldData.Presets = {
              Presets: globalPreset,
            };
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              var presetValue = oldValues[nextRow][column + 4];
              if (!presetName) break;
              if (presetValue === globalPreset) {
                continue;
              }
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = presetValue;
            }
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_09_00_00eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v5.06.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_06_02_00eDamage: function (oldValues, oldeDamageLabValues) {
    try {
      console.log("Called: ePaths.getVersion5_06_02_00eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[1] &&
        oldeDamageLabValues[1][0] === "Running Time"
      ) {
        var eDamageLabCost = oldeDamageLabValues[0][0];
        if (String(eDamageLabCost)) {
          oldData.eDamageLabCost = eDamageLabCost;
        }
        var eDamageRunningTime = oldeDamageLabValues[2][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValueIndex = 4;
              if (["Run Type", "Simulated Tier"].includes(guessName)) {
                guessValueIndex -= 1;
              }
              var guessValue = oldValues[nextRow][column + guessValueIndex];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Presets") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              if (!presetName) break;
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = oldValues[nextRow][column + 3];
            }
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_06_02_00eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v5.05.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @param {Array<Array<*>>} cLDmgValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_05_00_00eDamage: function (
    oldValues,
    oldeDamageLabValues,
    cLDmgValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion5_05_00_00eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[1] &&
        oldeDamageLabValues[1][0] === "Running Time"
      ) {
        var eDamageLabCost = oldeDamageLabValues[0][0];
        if (String(eDamageLabCost)) {
          oldData.eDamageLabCost = eDamageLabCost;
        }
        var eDamageRunningTime = oldeDamageLabValues[2][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
        }
      }
      if (cLDmgValues && cLDmgValues[0] && cLDmgValues[0][0] !== null) {
        oldData.CLDamage = cLDmgValues[0][1];
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValueIndex = 4;
              if (["Run Type", "Simulated Tier"].includes(guessName)) {
                guessValueIndex -= 1;
              }
              var guessValue = oldValues[nextRow][column + guessValueIndex];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Presets") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              if (!presetName) break;
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = oldValues[nextRow][column + 3];
            }
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_05_00_00eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
        cLDmgValues: cLDmgValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v4.11.03.21 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @param {Array<Array<*>>} cLDmgValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_03_21eDamage: function (
    oldValues,
    oldeDamageLabValues,
    cLDmgValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion4_11_03_21eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[1] &&
        oldeDamageLabValues[1][0] === "Running Time"
      ) {
        var eDamageLabCost = oldeDamageLabValues[0][0];
        if (String(eDamageLabCost)) {
          oldData.eDamageLabCost = eDamageLabCost;
        }
        var eDamageRunningTime = oldeDamageLabValues[2][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
        }
      }
      if (cLDmgValues && cLDmgValues[0] && cLDmgValues[0][0] !== null) {
        oldData.CLDamage = cLDmgValues[0][1];
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Total Lab Time") {
            var preset = oldValues[row][column + 2];
            oldData.Presets = {
              Workshop: preset,
              Cards: preset,
              Modules: preset,
            };
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion4_11_03_21eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
        cLDmgValues: cLDmgValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v4.11.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_02_00eDamage: function (oldValues, oldeDamageLabValues) {
    try {
      console.log("Called: ePaths.getVersion4_11_02_00eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[0] &&
        oldeDamageLabValues[0][0] === "Running Time"
      ) {
        var eDamageRunningTime = oldeDamageLabValues[1][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Total Lab Time") {
            var preset = oldValues[row][column + 2];
            oldData.Presets = {
              Workshop: preset,
              Cards: preset,
              Modules: preset,
            };
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion4_11_02_00eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v5.09.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeEconStoneMultValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_09_00_00eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeEconStoneMultValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion5_09_00_00eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Speed Up"
      ) {
        var eEconLabRoi = oldeEconLabValues[0][0];
        if (String(eEconLabRoi)) {
          oldData.eEconLabRoi = eEconLabRoi;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeEconStoneMultValues &&
        oldeEconStoneMultValues[1] &&
        oldeEconStoneMultValues[1][0] === "Stone Multiplier"
      ) {
        var eEconStoneMult = oldeEconStoneMultValues[2][0];
        if (String(eEconStoneMult)) {
          oldData.eEconStoneMult = eEconStoneMult;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Speed Up"
      ) {
        var eDiscountLabRoi = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabRoi)) {
          oldData.eDiscountLabRoi = eDiscountLabRoi;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];

          if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 5];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 5];
            }
          } else if (cell === "User Inputs") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 5];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "GB Sync Desired Ratio:") {
                guessValue = {
                  "antecedent value": oldValues[nextRow][column + 3],
                  "consequent value": guessValue,
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            var assModLevel = oldValues[row][column + 5];
            var mainIsFormula =
              !moduleLevel || String(moduleLevel).startsWith("=");
            var assistIsFormula =
              !assModLevel || String(assModLevel).startsWith("=");
            if (mainIsFormula && assistIsFormula) {
              continue;
            }
            var moduleObj = {};
            if (!mainIsFormula) {
              moduleObj["main"] = moduleLevel;
            }
            if (!assistIsFormula) {
              moduleObj["assist"] = assModLevel;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleObj;
          } else if (cell === "Calculation Rows") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Coins Spent") {
            var enhancementDiscount = oldValues[row][column - 1];
            oldData.enhancementDiscount = enhancementDiscount;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_09_00_00eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeEconStoneMultValues: oldeEconStoneMultValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v5.08.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeEconStoneMultValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_08_00_00eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeEconStoneMultValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion5_08_00_00eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Speed Up"
      ) {
        var eEconLabRoi = oldeEconLabValues[0][0];
        if (String(eEconLabRoi)) {
          oldData.eEconLabRoi = eEconLabRoi;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeEconStoneMultValues &&
        oldeEconStoneMultValues[1] &&
        oldeEconStoneMultValues[1][0] === "Stone Multiplier"
      ) {
        var eEconStoneMult = oldeEconStoneMultValues[2][0];
        if (String(eEconStoneMult)) {
          oldData.eEconStoneMult = eEconStoneMult;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Speed Up"
      ) {
        var eDiscountLabRoi = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabRoi)) {
          oldData.eDiscountLabRoi = eDiscountLabRoi;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];

          if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Inputs") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "GB Sync Desired Ratio:") {
                guessValue = {
                  "antecedent value": oldValues[nextRow][column + 3],
                  "consequent value": guessValue,
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            var assModLevel = oldValues[row][column + 5];
            var mainIsFormula =
              !moduleLevel || String(moduleLevel).startsWith("=");
            var assistIsFormula =
              !assModLevel || String(assModLevel).startsWith("=");
            if (mainIsFormula && assistIsFormula) {
              continue;
            }
            var moduleObj = {};
            if (!mainIsFormula) {
              moduleObj["main"] = moduleLevel;
            }
            if (!assistIsFormula) {
              moduleObj["assist"] = assModLevel;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleObj;
          } else if (cell === "Calculation Rows") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Coins Spent") {
            var enhancementDiscount = oldValues[row][column - 1];
            oldData.enhancementDiscount = enhancementDiscount;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_08_00_00eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeEconStoneMultValues: oldeEconStoneMultValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v5.06.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_06_02_00eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion5_06_02_00eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Running Time"
      ) {
        var eEconLabCost = oldeEconLabValues[0][0];
        if (String(eEconLabCost)) {
          oldData.eEconLabCost = eEconLabCost;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Running Time"
      ) {
        var eDiscountLabCost = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabCost)) {
          oldData.eDiscountLabCost = eDiscountLabCost;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                if (guessName.includes("GB Sync Current Ratio")) {
                  guessName = "GB Sync Current Ratio";
                } else {
                  var parts = guessName.split(",");
                  guessName = parts[parts.length - 2]
                    .replace(/["\(\)]/g, "")
                    .trim();
                }
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "Ignore Lab Target Levels") {
                oldData.UserGuess["Ignore UW Target Levels"] = guessValue;
                guessValue = oldValues[nextRow][column + 1];
              } else if (guessName === "GB Sync Current Ratio") {
                guessName = "GB Sync Desired Ratio:";
                var guessValues = String(guessValue).split("/");
                guessValue = {
                  "antecedent value": guessValues[0].trim(),
                  "consequent value": guessValues[1].trim(),
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            var assModLevel = oldValues[row][column + 5];
            var mainIsFormula =
              !moduleLevel || String(moduleLevel).startsWith("=");
            var assistIsFormula =
              !assModLevel || String(assModLevel).startsWith("=");
            if (mainIsFormula && assistIsFormula) {
              continue;
            }
            var moduleObj = {};
            if (!mainIsFormula) {
              var moduleValue = String(moduleLevel).split("|")[0].trim();
              moduleObj["main"] = moduleValue;
            }
            if (!assistIsFormula) {
              var assistValue = String(assModLevel).split("|")[0].trim();
              moduleObj["assist"] = assistValue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleObj;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_06_02_00eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v5.00.01.04 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_00_01_04eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion5_00_01_04eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Running Time"
      ) {
        var eEconLabCost = oldeEconLabValues[0][0];
        if (String(eEconLabCost)) {
          oldData.eEconLabCost = eEconLabCost;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Running Time"
      ) {
        var eDiscountLabCost = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabCost)) {
          oldData.eDiscountLabCost = eDiscountLabCost;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                if (guessName.includes("GB Sync Current Ratio")) {
                  guessName = "GB Sync Current Ratio";
                } else {
                  var parts = guessName.split(",");
                  guessName = parts[parts.length - 2]
                    .replace(/["\(\)]/g, "")
                    .trim();
                }
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "Turn off Labs in Coin Path") {
                oldData.UserGuess["Ignore Target Levels"] = guessValue;
                guessValue = oldValues[nextRow][column + 1];
              } else if (guessName === "GB Sync Current Ratio") {
                guessName = "GB Sync Desired Ratio:";
                var guessValues = String(guessValue).split("/");
                guessValue = {
                  "antecedent value": guessValues[0].trim(),
                  "consequent value": guessValues[1].trim(),
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            var assModLevel = oldValues[row][column + 5];
            var mainIsFormula =
              !moduleLevel || String(moduleLevel).startsWith("=");
            var assistIsFormula =
              !assModLevel || String(assModLevel).startsWith("=");
            if (mainIsFormula && assistIsFormula) {
              continue;
            }
            var moduleObj = {};
            if (!mainIsFormula) {
              var moduleValue = String(moduleLevel).split("|")[0].trim();
              moduleObj["main"] = moduleValue;
            }
            if (!assistIsFormula) {
              var assistValue = String(assModLevel).split("|")[0].trim();
              moduleObj["assist"] = assistValue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleObj;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion5_00_01_04eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v4.11.03.21 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_03_21eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion4_11_03_21eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Running Time"
      ) {
        var eEconLabCost = oldeEconLabValues[0][0];
        if (String(eEconLabCost)) {
          oldData.eEconLabCost = eEconLabCost;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Running Time"
      ) {
        var eDiscountLabCost = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabCost)) {
          oldData.eDiscountLabCost = eDiscountLabCost;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                if (guessName.includes("GB Sync Current Ratio")) {
                  guessName = "GB Sync Current Ratio";
                } else {
                  var parts = guessName.split(",");
                  guessName = parts[parts.length - 2]
                    .replace(/["\(\)]/g, "")
                    .trim();
                }
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "GB Sync Current Ratio") {
                guessName = "GB Sync Desired Ratio:";
                var guessValues = String(guessValue).split("/");
                guessValue = {
                  "antecedent value": guessValues[0].trim(),
                  "consequent value": guessValues[1].trim(),
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            var moduleValue = String(moduleLevel).split("|")[0].trim();
            oldData.Modules[cell] = { main: moduleValue };
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion4_11_03_21eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v4.11.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_02_00eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePaths.getVersion4_11_02_00eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[0] &&
        oldeEconLabValues[0][0] === "Running Time"
      ) {
        var eEconRunningTime = oldeEconLabValues[1][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[0] &&
        oldeDiscountLabValues[0][0] === "Running Time"
      ) {
        var eDiscountRunningTime = oldeDiscountLabValues[1][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                if (guessName.includes("GB Sync Current Ratio")) {
                  guessName = "GB Sync Current Ratio";
                } else {
                  var parts = guessName.split(",");
                  guessName = parts[parts.length - 2]
                    .replace(/["\(\)]/g, "")
                    .trim();
                }
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "GB Sync Current Ratio") {
                guessName = "GB Sync Desired Ratio:";
                var guessValues = String(guessValue).split("/");
                guessValue = {
                  "antecedent value": guessValues[0].trim(),
                  "consequent value": guessValues[1].trim(),
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            var moduleValue = String(moduleLevel).split("|")[0].trim();
            oldData.Modules[cell] = moduleValue;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.getVersion4_11_02_00eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v4.11.02.00": this.version4_11_02_00.bind(this),
      "v4.11.03.21": this.version4_11_03_21.bind(this),
      "v5.00.01.04": this.version5_00_01_04.bind(this),
      "v5.03.00.00": this.version5_03_00_00.bind(this),
      "v5.05.00.00": this.version5_05_00_00.bind(this),
      "v5.05.01.00": this.version5_05_01_00.bind(this),
      "v5.06.02.00": this.version5_06_02_00.bind(this),
      "v5.08.00.00": this.version5_08_00_00.bind(this),
      "v5.08.04.00": this.version5_08_04_00.bind(this),
      "v5.09.00.00": this.version5_09_00_00.bind(this),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = Object.keys(this.convertVersionFunctions);
    var sortedThresholds = versionCompatibility.slice().sort(function (a, b) {
      return shared.compareVersions(b, a) === "newer" ? 1 : -1;
    });
    for (var key = 0; key < sortedThresholds.length; key++) {
      var threshold = sortedThresholds[key];
      var compareResult = shared.compareVersions(oldVersion, threshold);
      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }
    return null;
  },

};
