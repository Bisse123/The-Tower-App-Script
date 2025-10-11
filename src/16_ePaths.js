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
      var ranges = [eHPRange, eDamageRange, eEconRange, "IDS"];
      
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

      var eHPValues = batchGetResult[0].values;
      var eDamageValues = batchGetResult[1].values;
      var eEconValues = batchGetResult[2].values;
      var idsData = batchGetResult[3].values;

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
          eHPColumnOffset
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
          eDamageColumnOffset
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
          eEconColumnOffset
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

      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });

        var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        if (!updateResult) {
          return {
            success: false,
            message: "Failed to update new spreadsheet™ with Effective Paths data.",
          };
        }
        return {
          success: true,
          message: `Effective Paths data imported successfully`,
        };
      }
      return {
        success: true,
        message: "No Effective Paths data to update",
      };
    } catch (error) {
      console.log(`Error importing ePaths data: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing Effective Paths data: ${error.message}`,
      };
    }
  },

  updateEHP: function (sheetName, oldData, eHPData, columnOffset) {
    try {
      console.log("Called: ePaths.updateEHP");
      
      // Add debugging to check the structure of oldData
      console.log("oldData structure:", oldData);
      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }
      
      var batchUpdate = [];
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

  updateEDamage: function (sheetName, oldData, eDamageData, columnOffset) {
    try {
      console.log("Called: ePaths.updateEDamage");
      
      // Add debugging to check the structure of oldData
      console.log("oldData structure:", oldData);
      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }
      
      var batchUpdate = [];
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

  updateEEcon: function (sheetName, oldData, eEconData, columnOffset) {
    try {
      console.log("Called: ePaths.updateEEcon");
      
      // Add debugging to check the structure of oldData
      console.log("oldData structure:", oldData);
      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }
      
      var batchUpdate = [];
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
    var ranges = [eHPRange, eDamageRange, eEconRange];
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

    var eHPData = this.getVersion4110200eHP(eHPValues);
    var eDamageData = this.getVersion4110200eDamage(eDamageValues);
    var eEconData = this.getVersion4110200eEcon(eEconValues);

    return {
      success: true,
      eHP: eHPData,
      eDamage: eDamageData,
      eEcon: eEconData,
    };
  },

  getVersion4110200eHP: function (oldValues) {
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

  getVersion4110200eDamage: function (oldValues) {
    try {
      console.log("Called: ePaths.getVersion4110200eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = { Custom: {}, Perks: {}, UserGuess: {}, Modules: {} };

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

  getVersion4110200eEcon: function (oldValues) {
    try {
      console.log("Called: ePaths.getVersion4110200eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = { Custom: {}, Perks: {}, UserGuess: {}, Modules: {} };
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
