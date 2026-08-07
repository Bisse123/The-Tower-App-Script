const ePaths = {
  // #region Export Functions
  exportData: function (versionDifference) {
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
      var oldDataResult = getVersionFunction();
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
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: ePaths.importData");

      var newSpreadsheet = spreadsheets("Effective Paths newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(newSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(newSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(newSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "New spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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

      // Calculate column offsets for each range
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
      // Call the respective update functions
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
      console.log(`Error importing ePaths data: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing Effective Paths data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Update Functions

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
            // Search rows below "Total Value" in column j - 2 for custom names to update
            for (var nextRow = row + 1; nextRow < eHPData.length; nextRow++) {
              var customName = eHPData[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
          } else if (cell === "Perks") {
            var perksCol = shared.columnToLetter(columnOffset + column + 6);
            if (oldData.Perks && oldData.Perks.hasOwnProperty("Active")) {
              var perksAreActive = oldData.Perks["Active"];
              var perksCellAddress = `${perksCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${perksCellAddress}`,
                values: [[perksAreActive]],
              });
            }
            for (var nextRow = row + 2; nextRow < eHPData.length; nextRow++) {
              var perkName = eHPData[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (oldData.Perks && oldData.Perks.hasOwnProperty(perkName)) {
                var perkValue = oldData.Perks[perkName];
                var perkCellAddress = `${perksCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${perkCellAddress}`,
                  values: [[perkValue]],
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
          } else if (cell === "Presets") {
            if (oldData.hasOwnProperty("Presets")) {
              var presetValues = oldData.Presets;
              var presetCol = shared.columnToLetter(columnOffset + column + 5);
              for (var nextRow = row; nextRow < eHPData.length; nextRow++) {
                var presetName = eHPData[nextRow][column];
                if (!presetName) break;
                if (presetValues.hasOwnProperty(presetName)) {
                  var presetValue = presetValues[presetName];
                  var presetCellAddress = `${presetCol}${nextRow + 1}`;
                  batchUpdate.push({
                    range: `${sheetName}!${presetCellAddress}`,
                    values: [[presetValue]],
                  });
                }
              }
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
      console.log(`Error in ePaths.updateEHP: ${error.toString()}`);
      return {
        success: false,
        message: "Error in ePaths.updateEHP: " + error.message,
      };
    }
  },

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

      // if (oldData.hasOwnProperty("CLDamage")) {
      //   var cLDamageValue = oldData.CLDamage;
      //   batchUpdate.push({
      //     range: `${sheetName}!AM149`,
      //     values: [[cLDamageValue]],
      //   });
      // }

      for (var row = 0; row < eDamageData.length; row++) {
        for (var column = 0; column < eDamageData[row].length; column++) {
          var cell = eDamageData[row][column];
          if (cell === "Total Value") {
            // Search rows below "Total Value" in column j - 2 for custom names to update
            for (
              var nextRow = row + 1;
              nextRow < eDamageData.length;
              nextRow++
            ) {
              var customName = eDamageData[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
          } else if (cell === "Perks") {
            var perksCol = shared.columnToLetter(columnOffset + column + 6);
            if (oldData.Perks && oldData.Perks.hasOwnProperty("Active")) {
              var perksAreActive = oldData.Perks["Active"];
              var perksCellAddress = `${perksCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${perksCellAddress}`,
                values: [[perksAreActive]],
              });
            }
            for (
              var nextRow = row + 2;
              nextRow < eDamageData.length;
              nextRow++
            ) {
              var perkName = eDamageData[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (oldData.Perks && oldData.Perks.hasOwnProperty(perkName)) {
                var perkValue = oldData.Perks[perkName];
                var perkCellAddress = `${perksCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${perkCellAddress}`,
                  values: [[perkValue]],
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
          } else if (cell === "Presets") {
            if (oldData.hasOwnProperty("Presets")) {
              var presetValues = oldData.Presets;
              var presetCol = shared.columnToLetter(columnOffset + column + 5);
              for (var nextRow = row; nextRow < eDamageData.length; nextRow++) {
                var presetName = eDamageData[nextRow][column];
                if (!presetName) break;
                if (presetValues.hasOwnProperty(presetName)) {
                  var presetValue = presetValues[presetName];
                  var presetCellAddress = `${presetCol}${nextRow + 1}`;
                  batchUpdate.push({
                    range: `${sheetName}!${presetCellAddress}`,
                    values: [[presetValue]],
                  });
                }
              }
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
      console.log(`Error in ePaths.updateEDamage: ${error.toString()}`);
      return {
        success: false,
        message: "Error in ePaths.updateEDamage: " + error.message,
      };
    }
  },

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
          // if (cell === "Total Value") {
          //   // Search rows below "Total Value" in column j - 2 for custom names to update
          //   for (var nextRow = row + 1; nextRow < eEconData.length; nextRow++) {
          //     var customName = eEconData[nextRow][column - 2];
          //     if (!customName) break; // Stop when customName is empty
          //     if (oldData.Custom && oldData.Custom.hasOwnProperty(customName)) {
          //       var newValue = oldData.Custom[customName];
          //       var cellCol = shared.columnToLetter(columnOffset + (column - 1) + 1);
          //       var cellAddress = `${cellCol}${nextRow + 1}`;
          //       batchUpdate.push({
          //         range: `${sheetName}!${cellAddress}`,
          //         values: [[newValue]],
          //       });
          //     }
          //   }
          // } else
          if (cell === "Perks") {
            var perksCol = shared.columnToLetter(columnOffset + column + 6);
            if (oldData.Perks && oldData.Perks.hasOwnProperty("Active")) {
              var perksAreActive = oldData.Perks["Active"];
              var perksCellAddress = `${perksCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${perksCellAddress}`,
                values: [[perksAreActive]],
              });
            }
            for (var nextRow = row + 2; nextRow < eEconData.length; nextRow++) {
              var perkName = eEconData[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (oldData.Perks && oldData.Perks.hasOwnProperty(perkName)) {
                var perkValue = oldData.Perks[perkName];
                var perkCellAddress = `${perksCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${perkCellAddress}`,
                  values: [[perkValue]],
                });
              }
            }
          } else if (cell === "User Inputs") {
            for (var nextRow = row + 1; nextRow < eEconData.length; nextRow++) {
              var guessName = eEconData[nextRow][column];
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
      console.log(`Error in ePaths.updateEEcon: ${error.toString()}`);
      return {
        success: false,
        message: "Error in ePaths.updateEEcon: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version5_09_00_00: function () {
    try {
      console.log("Called: ePaths.version5_09_00_00");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version5_09_00_00: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version5_08_04_00: function () {
    try {
      console.log("Called: ePaths.version5_08_04_00");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version5_08_04_00: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version5_08_00_00: function () {
    try {
      console.log("Called: ePaths.version5_08_00_00");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version5_08_00_00: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version5_06_02_00: function () {
    try {
      console.log("Called: ePaths.version5_06_02_00");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version5_06_02_00: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version5_05_01_00: function () {
    try {
      console.log("Called: ePaths.version5_05_01_00");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version5_05_01_00: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version5_05_00_00: function () {
    try {
      console.log("Called: ePaths.version5_05_00_00");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version5_05_00_00: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version5_03_00_00: function () {
    try {
      console.log("Called: ePaths.version5_03_00_00");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version5_03_00_00: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version5_00_01_04: function () {
    try {
      console.log("Called: ePaths.version5_00_01_04");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version5_00_01_04: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version4_11_03_21: function () {
    try {
      console.log("Called: ePaths.version4_11_03_21");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version4_11_03_21: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  version4_11_02_00: function () {
    try {
      console.log("Called: ePaths.version4_11_02_00");
      var oldSpreadsheet = spreadsheets("Effective Paths oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eHP") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eDamage") ||
        !SheetsAPI.getSheetByName(oldSpreadsheet, "eEcon")
      ) {
        return {
          success: false,
          message:
            "Old spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
        };
      }

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
      console.log(`Error in ePaths.version4_11_02_00: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting Effective Paths data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get eHP
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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion5_09_00_00eHP: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_09_00_00eHP: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion5_05_01_00eHP: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_05_01_00eHP: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion5_03_00_00eHP: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_03_00_00eHP: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion4_11_03_21eHP: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4_11_03_21eHP: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion4_11_02_00eHP: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4_11_02_00eHP: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get eDamage
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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion5_09_00_00eDamage: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_09_00_00eDamage: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion5_06_02_00eDamage: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_06_02_00eDamage: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion5_05_00_00eDamage: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_05_00_00eDamage: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion4_11_03_21eDamage: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4_11_03_21eDamage: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion4_11_02_00eDamage: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4_11_02_00eDamage: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get eEcon
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
          // if (cell === "Total Value") {
          //   // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
          //   for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
          //     var customName = oldValues[nextRow][column - 2];
          //     if (!customName) break; // Stop when customName is empty
          //     if (customData.includes(customName)) {
          //       var customValue = oldValues[nextRow][column - 1];
          //       if (
          //         !String(customValue) ||
          //         String(customValue).startsWith("=")
          //       ) {
          //         continue;
          //       }
          //       if (!oldData.hasOwnProperty("Custom")) {
          //         oldData.Custom = {};
          //       }
          //       oldData.Custom[customName] = customValue;
          //     }
          //   }
          // } else
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
      console.log(`Error in getVersion5_09_00_00eEcon: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_09_00_00eEcon: " + error.message,
      };
    }
  },

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
          // if (cell === "Total Value") {
          //   // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
          //   for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
          //     var customName = oldValues[nextRow][column - 2];
          //     if (!customName) break; // Stop when customName is empty
          //     if (customData.includes(customName)) {
          //       var customValue = oldValues[nextRow][column - 1];
          //       if (
          //         !String(customValue) ||
          //         String(customValue).startsWith("=")
          //       ) {
          //         continue;
          //       }
          //       if (!oldData.hasOwnProperty("Custom")) {
          //         oldData.Custom = {};
          //       }
          //       oldData.Custom[customName] = customValue;
          //     }
          //   }
          // } else
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
      console.log(`Error in getVersion5_08_00_00eEcon: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_08_00_00eEcon: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion5_06_02_00eEcon: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_06_02_00eEcon: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion5_00_01_04eEcon: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion5_00_01_04eEcon: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion4_11_03_21eEcon: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4_11_03_21eEcon: " + error.message,
      };
    }
  },

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
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break; // Stop when customName is empty
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
      console.log(`Error in getVersion4_11_02_00eEcon: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4_11_02_00eEcon: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Version Functions Getter
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

  // #endregion
  // #region Compatibility Check
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
  // #endregion
};
