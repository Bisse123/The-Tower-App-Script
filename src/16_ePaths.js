const ePaths = {
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
          message: "New spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).",
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
        "IDS", eHPRange, eDamageRange, eEconRange,
        eHPLabRange, eRegenLabRange, eDamageLabRange, eEconLabRange, eDiscountLabRange];
      
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
      var eDiscountLabValues = batchGetResult[8].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(
        newSheetID,
        "IDS",
        "IDS Master's",
        idsData
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
          eDamageLabValues
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
          eDiscountLabValues
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
      shared.addIDUpdatesToBatch(batchUpdate, "Effective Paths", newSheetID, idsData, data.idMasterID);

      var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
      if (!updateResult) {
        return {
          success: false,
          message: "Failed to update new spreadsheet™ with Effective Paths data.",
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

  updateEHP: function (sheetName, oldData, eHPData, columnOffset, eHPLabData, eRegenLabData) {
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

      if (eHPLabData && eHPLabData[0] && eHPLabData[0][0] === "Running Time" && oldData.hasOwnProperty("eHPRunningTime")) {
        var eHPRunningTimeValue = oldData.eHPRunningTime;
        console.log("eHPRunningTimeValue:", eHPRunningTimeValue);
        batchUpdate.push({
          range: `${sheetName}!L5`,
          values: [[eHPRunningTimeValue]],
        });
      }
      if (eRegenLabData && eRegenLabData[0] && eRegenLabData[0][0] === "Running Time" && oldData.hasOwnProperty("eRegenRunningTime")) {
        var eRegenRunningTimeValue = oldData.eRegenRunningTime;
        console.log("eRegenRunningTimeValue:", eRegenRunningTimeValue);
        batchUpdate.push({
          range: `${sheetName}!AA5`,
          values: [[eRegenRunningTimeValue]],
        });
      }

      for (var i = 0; i < eHPData.length; i++) {
        for (var j = 0; j < eHPData[i].length; j++) {
          var cell = eHPData[i][j];
          if (cell === "Total Value") {
            // Search rows below "Total Value" in column j - 2 for custom names to update
            for (var k = i + 1; k < eHPData.length; k++) {
              var customName = eHPData[k][j - 2];
              if (!customName) break; // Stop when customName is empty
              if (oldData.Custom && oldData.Custom.hasOwnProperty(customName)) {
                var newValue = oldData.Custom[customName];
                var cellCol = shared.columnToLetter(columnOffset + (j - 1) + 1);
                var cellAddress = `${cellCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${cellAddress}`,
                  values: [[newValue]],
                });
              }
            }
          } else if (cell === "Perks") {
            if (oldData.Perks && oldData.Perks.hasOwnProperty("Active")) {
              var perksAreActive = oldData.Perks["Active"];
              var perksCol = shared.columnToLetter(columnOffset + j + 5);
              var perksCellAddress = `${perksCol}${i + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${perksCellAddress}`,
                values: [[perksAreActive]],
              });
            }
            for (var k = i + 2; k < eHPData.length; k++) {
              var perkName = eHPData[k][j];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (oldData.Perks && oldData.Perks.hasOwnProperty(perkName)) {
                var perkValue = oldData.Perks[perkName];
                var perkCol = shared.columnToLetter(columnOffset + j + 5);
                var perkCellAddress = `${perkCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${perkCellAddress}`,
                  values: [[perkValue]],
                });
              }
            }
          } else if (cell === "User Specific Guesses") {
            for (var k = i + 1; k < eHPData.length; k++) {
              var guessName = eHPData[k][j];
              if (!guessName) break;
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (oldData.UserGuess && oldData.UserGuess.hasOwnProperty(guessName)) {
                var guessValue = oldData.UserGuess[guessName];
                var guessCol = shared.columnToLetter(columnOffset + j + 5);
                var guessCellAddress = `${guessCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${guessCellAddress}`,
                  values: [[guessValue]],
                });
              }
            }
          } else if (cell === "Rows Calculated") {
            if (oldData.hasOwnProperty("rowsCalculated")) {
              var rowsCalculatedValue = oldData.rowsCalculated;
              var rowsCol = shared.columnToLetter(columnOffset + j + 1);
              var rowsCalculatedCellAddress = `${rowsCol}${i + 2}`;
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
      console.log(`Error in ePaths.updateEHP: ${error.toString()}`);
      return {
        success: false,
        message: "Error in ePaths.updateEHP: " + error.message,
      };
    }
  },

  updateEDamage: function (sheetName, oldData, eDamageData, columnOffset, eDamageLabData) {
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

      if (eDamageLabData && eDamageLabData[0] && eDamageLabData[0][0] === "Running Time" && oldData.hasOwnProperty("eDamageRunningTime")) {
        var eDamageRunningTimeValue = oldData.eDamageRunningTime;
        console.log("eDamageRunningTimeValue:", eDamageRunningTimeValue);
        batchUpdate.push({
          range: `${sheetName}!L5`,
          values: [[eDamageRunningTimeValue]],
        });
      }

      for (var i = 0; i < eDamageData.length; i++) {
        for (var j = 0; j < eDamageData[i].length; j++) {
          var cell = eDamageData[i][j];
          if (cell === "Total Value") {
            // Search rows below "Total Value" in column j - 2 for custom names to update
            for (var k = i + 1; k < eDamageData.length; k++) {
              var customName = eDamageData[k][j - 2];
              if (!customName) break; // Stop when customName is empty
              if (oldData.Custom && oldData.Custom.hasOwnProperty(customName)) {
                var newValue = oldData.Custom[customName];
                var cellCol = shared.columnToLetter(columnOffset + (j - 1) + 1);
                var cellAddress = `${cellCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${cellAddress}`,
                  values: [[newValue]],
                });
              }
            }
          } else if (cell === "Perks") {
            if (oldData.Perks && oldData.Perks.hasOwnProperty("Active")) {
              var perksAreActive = oldData.Perks["Active"];
              var perksCol = shared.columnToLetter(columnOffset + j + 5);
              var perksCellAddress = `${perksCol}${i + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${perksCellAddress}`,
                values: [[perksAreActive]],
              });
            }
            for (var k = i + 2; k < eDamageData.length; k++) {
              var perkName = eDamageData[k][j];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (oldData.Perks && oldData.Perks.hasOwnProperty(perkName)) {
                var perkValue = oldData.Perks[perkName];
                var perkCol = shared.columnToLetter(columnOffset + j + 5);
                var perkCellAddress = `${perkCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${perkCellAddress}`,
                  values: [[perkValue]],
                });
              }
            }
          } else if (cell === "User Specific Guesses") {
            for (var k = i + 1; k < eDamageData.length; k++) {
              var guessName = eDamageData[k][j];
              if (!guessName) break;
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (oldData.UserGuess && oldData.UserGuess.hasOwnProperty(guessName)) {
                var guessValue = oldData.UserGuess[guessName];
                var guessCol = shared.columnToLetter(columnOffset + j + 5);
                var guessCellAddress = `${guessCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${guessCellAddress}`,
                  values: [[guessValue]],
                });
              }
            }
          } else if (oldData.Modules && oldData.Modules.hasOwnProperty(cell)) {
            var moduleValue = oldData.Modules[cell];
            var moduleCol = shared.columnToLetter(columnOffset + j + 2);
            var moduleCellAddress = `${moduleCol}${i + 1}`;
            batchUpdate.push({
              range: `${sheetName}!${moduleCellAddress}`,
              values: [[moduleValue]],
            });
          } else if (cell === "Rows Calculated") {
            if (oldData.hasOwnProperty("rowsCalculated")) {
              var rowsCalculatedValue = oldData.rowsCalculated;
              var rowsCol = shared.columnToLetter(columnOffset + j + 1);
              var rowsCalculatedCellAddress = `${rowsCol}${i + 2}`;
              batchUpdate.push({
                range: `${sheetName}!${rowsCalculatedCellAddress}`,
                values: [[rowsCalculatedValue]],
              });
            }
          } else if (cell === "Total Lab Time") {
            if (oldData.hasOwnProperty("Preset")) {
              var presetValue = oldData.Preset;
              var presetCol = shared.columnToLetter(columnOffset + j + 3);
              var presetCellAddress = `${presetCol}${i + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${presetCellAddress}`,
                values: [[presetValue]],
              });
            }
          } else if (cell === "PS Beta Testing") {
            if (oldData.hasOwnProperty("PSBeta")) {
              var psBetaValue = oldData.PSBeta;
              var psBetaCol = shared.columnToLetter(columnOffset + j);
              var psBetaCellAddress = `${psBetaCol}${i}`;
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

  updateEEcon: function (sheetName, oldData, eEconData, columnOffset, eEconLabData, eDiscountLabData) {
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

      if (eEconLabData && eEconLabData[0] && eEconLabData[0][0] === "Running Time" && oldData.hasOwnProperty("eEconRunningTime")) {
        var eEconRunningTimeValue = oldData.eEconRunningTime;
        console.log("eEconRunningTimeValue:", eEconRunningTimeValue);
        batchUpdate.push({
          range: `${sheetName}!L5`,
          values: [[eEconRunningTimeValue]],
        });
      }
      if (eDiscountLabData && eDiscountLabData[0] && eDiscountLabData[0][0] === "Running Time" && oldData.hasOwnProperty("eDiscountRunningTime")) {
        var eDiscountRunningTime = oldData.eDiscountRunningTime;
        console.log("discountValue:", eDiscountRunningTime);
        batchUpdate.push({
          range: `${sheetName}!AG5`,
          values: [[eDiscountRunningTime]],
        });
      }

      for (var i = 0; i < eEconData.length; i++) {
        for (var j = 0; j < eEconData[i].length; j++) {
          var cell = eEconData[i][j];
          if (cell === "Total Value") {
            // Search rows below "Total Value" in column j - 2 for custom names to update
            for (var k = i + 1; k < eEconData.length; k++) {
              var customName = eEconData[k][j - 2];
              if (!customName) break; // Stop when customName is empty
              if (oldData.Custom && oldData.Custom.hasOwnProperty(customName)) {
                var newValue = oldData.Custom[customName];
                var cellCol = shared.columnToLetter(columnOffset + (j - 1) + 1);
                var cellAddress = `${cellCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${cellAddress}`,
                  values: [[newValue]],
                });
              }
            }
          } else if (cell === "Perks") {
            if (oldData.Perks && oldData.Perks.hasOwnProperty("Active")) {
              var perksAreActive = oldData.Perks["Active"];
              var perksCol = shared.columnToLetter(columnOffset + j + 5);
              var perksCellAddress = `${perksCol}${i + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${perksCellAddress}`,
                values: [[perksAreActive]],
              });
            }
            for (var k = i + 2; k < eEconData.length; k++) {
              var perkName = eEconData[k][j];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (oldData.Perks && oldData.Perks.hasOwnProperty(perkName)) {
                var perkValue = oldData.Perks[perkName];
                var perkCol = shared.columnToLetter(columnOffset + j + 5);
                var perkCellAddress = `${perkCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${perkCellAddress}`,
                  values: [[perkValue]],
                });
              }
            }
          } else if (cell === "User Specific Guesses") {
            for (var k = i + 1; k < eEconData.length; k++) {
              var guessName = eEconData[k][j];
              if (!guessName) break;
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (oldData.UserGuess && oldData.UserGuess.hasOwnProperty(guessName)) {
                var guessValue = oldData.UserGuess[guessName];
                var guessCol = shared.columnToLetter(columnOffset + j + 5);
                var guessCellAddress = `${guessCol}${k + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${guessCellAddress}`,
                  values: [[guessValue]],
                });
              }
            }
          } else if (oldData.Modules && oldData.Modules.hasOwnProperty(cell)) {
            var moduleValue = oldData.Modules[cell];
            var moduleCol = shared.columnToLetter(columnOffset + j + 2);
            var moduleCellAddress = `${moduleCol}${i + 1}`;
            batchUpdate.push({
              range: `${sheetName}!${moduleCellAddress}`,
              values: [[moduleValue]],
            });
          } else if (cell === "Rows Calculated") {
            if (oldData.hasOwnProperty("rowsCalculated")) {
              var rowsCalculatedValue = oldData.rowsCalculated;
              var rowsCol = shared.columnToLetter(columnOffset + j + 1);
              var rowsCalculatedCellAddress = `${rowsCol}${i + 2}`;
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

  version4110200: function () {
    console.log("Called: ePaths.version4110200");
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
      eHPRange, eDamageRange, eEconRange,
      eHPLabRange, eRegenLabRange, eDamageLabRange, eEconLabRange, eDiscountLabRange
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

    var eHPData = this.getVersion4110200eHP(eHPValues, eHPLabValues, eRegenLabValues);
    var eDamageData = this.getVersion4110200eDamage(eDamageValues, eDamageLabValues);
    var eEconData = this.getVersion4110200eEcon(eEconValues, eEconLabValues, eDiscountLabValues);

    return {
      success: true,
      eHP: eHPData,
      eDamage: eDamageData,
      eEcon: eEconData,
    };
  },

  getVersion4110200eHP: function (oldValues, oldeHPLabValues, oldeRegenLabValues) {
    try {
      console.log("Called: ePaths.getVersion4110200eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = { Custom: {}, Perks: {}, UserGuess: {}, Modules: {} };

      if (oldeHPLabValues && oldeHPLabValues[0] && oldeHPLabValues[0][0] === "Running Time") {
        var eHPRunningTime = oldeHPLabValues[1][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (oldeRegenLabValues && oldeRegenLabValues[0] && oldeRegenLabValues[0][0] === "Running Time") {
        var eRegenRunningTime = oldeRegenLabValues[1][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
        }
      }

      for (var i = 0; i < oldValues.length; i++) {
        for (var j = 0; j < oldValues[i].length; j++) {
          var cell = oldValues[i][j];
          if (cell === "Total Value") {
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var k = i + 1; k < oldValues.length; k++) {
              var customName = oldValues[k][j - 2];
              if (!customName) break; // Stop when customName is empty
              if (customData.includes(customName)) {
                var customValue = oldValues[k][j - 1];
                if (!String(customValue) || String(customValue).startsWith("=")) {
                  continue;
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[i][j + 4];
            if (String(perksAreActive) && !String(perksAreActive).startsWith("=")) {
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var k = i + 2; k < oldValues.length; k++) {
              var perkName = oldValues[k][j];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              oldData.Perks[perkName] = oldValues[k][j + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var k = i + 1; k < oldValues.length; k++) {
              var guessName = oldValues[k][j];
              if (!guessName) break;
              var guessValue = oldValues[k][j + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[i][j + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[i + 1][j];
            if (!String(rowsCalculated) || String(rowsCalculated).startsWith("=")) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[i + 1][j];
            console.log("Found running time:", runningTime);
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
      console.log(`Error in getVersion4110200eHP: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4110200eHP: " + error.message,
      };
    }
  },

  getVersion4110200eDamage: function (oldValues, oldeDamageLabValues) {
    try {
      console.log("Called: ePaths.getVersion4110200eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = { Custom: {}, Perks: {}, UserGuess: {}, Modules: {} };

      if (oldeDamageLabValues && oldeDamageLabValues[0] && oldeDamageLabValues[0][0] === "Running Time") {
        var eDamageRunningTime = oldeDamageLabValues[1][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
        }
      }

      for (var i = 0; i < oldValues.length; i++) {
        for (var j = 0; j < oldValues[i].length; j++) {
          var cell = oldValues[i][j];
          if (cell === "Total Value") {
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var k = i + 1; k < oldValues.length; k++) {
              var customName = oldValues[k][j - 2];
              if (!customName) break; // Stop when customName is empty
              if (customData.includes(customName)) {
                var customValue = oldValues[k][j - 1];
                if (!String(customValue) || String(customValue).startsWith("=")) {
                  continue;
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[i][j + 4];
            if (String(perksAreActive) && !String(perksAreActive).startsWith("=")) {
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var k = i + 2; k < oldValues.length; k++) {
              var perkName = oldValues[k][j];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              oldData.Perks[perkName] = oldValues[k][j + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var k = i + 1; k < oldValues.length; k++) {
              var guessName = oldValues[k][j];
              if (!guessName) break;
              var guessValue = oldValues[k][j + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[i][j + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[i + 1][j];
            if (!String(rowsCalculated) || String(rowsCalculated).startsWith("=")) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[i + 1][j];
            console.log("Found running time:", runningTime);
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Total Lab Time") {
            oldData.Preset = oldValues[i][j + 2];
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[i - 1][j];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      console.log(`Error in getVersion4110200eDamage: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4110200eDamage: " + error.message,
      };
    }
  },

  getVersion4110200eEcon: function (oldValues, oldeEconLabValues, oldeDiscountLabValues) {
    try {
      console.log("Called: ePaths.getVersion4110200eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = { Custom: {}, Perks: {}, UserGuess: {}, Modules: {} };

      if (oldeEconLabValues && oldeEconLabValues[0] && oldeEconLabValues[0][0] === "Running Time") {
        var eEconRunningTime = oldeEconLabValues[1][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (oldeDiscountLabValues && oldeDiscountLabValues[0] && oldeDiscountLabValues[0][0] === "Running Time") {
        var eDiscountRunningTime = oldeDiscountLabValues[1][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var i = 0; i < oldValues.length; i++) {
        for (var j = 0; j < oldValues[i].length; j++) {
          var cell = oldValues[i][j];
          if (cell === "Total Value") {
            // Search rows below "Total Value" in column j - 2 for custom names and j - 1 for values
            for (var k = i + 1; k < oldValues.length; k++) {
              var customName = oldValues[k][j - 2];
              if (!customName) break; // Stop when customName is empty
              if (customData.includes(customName)) {
                var customValue = oldValues[k][j - 1];
                if (!String(customValue) || String(customValue).startsWith("=")) {
                  continue;
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[i][j + 4];
            if (String(perksAreActive) && !String(perksAreActive).startsWith("=")) {
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var k = i + 2; k < oldValues.length; k++) {
              var perkName = oldValues[k][j];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              oldData.Perks[perkName] = oldValues[k][j + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var k = i + 1; k < oldValues.length; k++) {
              var guessName = oldValues[k][j];
              if (!guessName) break;
              var guessValue = oldValues[k][j + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[i][j + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[i + 1][j];
            if (!String(rowsCalculated) || String(rowsCalculated).startsWith("=")) {
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
      console.log(`Error in getVersion4110200eEcon: ${error.toString()}`);
      return {
        success: false,
        message: "Error in getVersion4110200eEcon: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v4.11.02.00": this.version4110200.bind(this),
    };
  },

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
};
