const workshop = {
  exportData: function (versionDifference) {
    try {
      console.log("Called: workshop.exportData");
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
        message: "Workshop export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting workshop data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      console.log("Called: workshop.importData");
      var newSpreadsheet = spreadsheets("Workshop newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }

      var requiredRanges = ["Master Sheet", "Desired Ratios", "IDS"];
      var batchResults = SheetsAPI.batchGetFormulas(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var desiredRatiosData = batchResults[1].values;
      var idsData = batchResults[2].values;

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

      // Only update workshop levels if key exists
      if (
        data.hasOwnProperty("oldWorkshopLevels") &&
        data.hasOwnProperty("oldWorkshopPlusLevels")
      ) {
        var oldWorkshopLevels = data.oldWorkshopLevels;
        var oldWorkshopPlusLevels = data.oldWorkshopPlusLevels;
        var workshopResult = this.updateWorkshopLevels(
          "Master Sheet",
          oldWorkshopLevels,
          oldWorkshopPlusLevels,
          masterSheetData
        );
        if (!workshopResult || !workshopResult.success) {
          console.log(
            `Error updating workshop levels: ${workshopResult.message}`
          );
          return workshopResult;
        }
        batchUpdate = batchUpdate.concat(workshopResult.batchUpdate || []);
      }

      // Only update workshop plus ratios if key exists
      if (data.hasOwnProperty("oldWorkshopPlusRatios")) {
        var oldWorkshopPlusRatios = data.oldWorkshopPlusRatios;
        var ratioResult = this.updateWorkshopPlusRatios(
          "Desired Ratios",
          oldWorkshopPlusRatios,
          desiredRatiosData
        );
        if (!ratioResult || !ratioResult.success) {
          console.log(
            `Error updating workshop plus ratios: ${ratioResult.message}`
          );
          return ratioResult;
        }
        batchUpdate = batchUpdate.concat(ratioResult.batchUpdate || []);
      }

      // Add import status update to batch if there were data updates
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(batchUpdate, "Workshop", newSheetID, idsData, data.idMasterID);

      // Apply all updates (including ID setting and import status)
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
        message: `Workshop import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: "Error importing workshop data: " + error.message,
      };
    }
  },

  updateWorkshopLevels: function (
    sheetName,
    oldWorkshopLevels,
    oldWorkshopPlusLevels,
    masterSheetData
  ) {
    try {
      console.log("Called: workshop.updateWorkshopLevels");
      if (!masterSheetData || masterSheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: `Not enough data in Master Sheet`,
        };
      }

      var headerRow = masterSheetData[0];

      var upgradeCol = 0;
      var enhancementCol = 0;

      for (var i = 0; i < headerRow.length; i++) {
        var cellValue = String(headerRow[i] || "").toLowerCase();
        if (cellValue.includes("workshop upgrade")) {
          upgradeCol = i + 1;
        } else if (cellValue.includes("workshop enhancement")) {
          enhancementCol = i + 1;
        }
        if (upgradeCol > 0 && enhancementCol > 0) {
          break;
        }
      }

      if (upgradeCol === 0) {
        console.log(`Workshop Upgrade column not found`);
        return {
          success: false,
          message: `Workshop Upgrade column not found`,
        };
      }
      if (enhancementCol === 0) {
        console.log(`Workshop Enhancement column not found`);
        return {
          success: false,
          message: `Workshop Enhancement column not found`,
        };
      }

      var workshopUnlocked = [];
      var workshopLevels = [];
      var workshopPlusLevels = [];
      var workshopLevelsStartCol = upgradeCol + 1;
      var workshopLevelsEndCol =
        workshopLevelsStartCol + oldWorkshopLevels.presetNames.length * 2 - 1;
      var workshopPlusLevelsStartCol = enhancementCol + 2;
      var workshopPlusLevelsEndCol =
        workshopPlusLevelsStartCol +
        oldWorkshopPlusLevels.presetNames.length -
        1;
      for (var i = 2; i < masterSheetData.length; i++) {
        var row = masterSheetData[i];

        var workshopName = row[upgradeCol - 1];
        if (workshopName && workshopName.startsWith("=")) {
          workshopName = workshopName
            .split(",")[1]
            .trim()
            .replace(/['"()]/g, "");
        }
        if (workshopName && oldWorkshopLevels.data[workshopName]) {
          var oldWorkshopRowData = oldWorkshopLevels.data[workshopName];
          workshopUnlocked.push([oldWorkshopRowData.unlocked || ""]);
          workshopLevels.push(oldWorkshopRowData.levels || []);
        }
        var enhancementName = row[enhancementCol - 1];
        if (enhancementName && enhancementName.startsWith("=")) {
          var parts = enhancementName.split(",");
          if (parts.length === 4) {
            enhancementName = parts[2].trim();
          } else {
            enhancementName = parts[parts.length - 1].trim();
          }
          enhancementName = enhancementName.replace(/['"()]/g, "");
        }
        if (enhancementName && oldWorkshopPlusLevels.data[enhancementName]) {
          var enhancementData = oldWorkshopPlusLevels.data[enhancementName];
          workshopPlusLevels.push(enhancementData || []);
        }
      }

      var batchUpdate = [];
      var startRow = 3;
      if (upgradeCol > 1 && workshopUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(upgradeCol - 1);
        var unlockedRange = `${sheetName}!${unlockedCol}${startRow}:${unlockedCol}${
          workshopUnlocked.length + startRow - 1
        }`;
        batchUpdate.push({
          range: unlockedRange,
          values: workshopUnlocked,
        });
      }

      if (upgradeCol > 0 && workshopLevels.length > 0) {
        var upgradeHeaders = [];
        oldWorkshopLevels.presetNames.forEach(function (presetName) {
          upgradeHeaders.push(presetName, "");
        });
        var levelsStartCol = shared.columnToLetter(workshopLevelsStartCol);
        var levelsEndCol = shared.columnToLetter(workshopLevelsEndCol);
        var levelsHeaderRange = `${sheetName}!${levelsStartCol}1:${levelsEndCol}1`;
        var levelsRange = `${sheetName}!${levelsStartCol}${startRow}:${levelsEndCol}${
          workshopLevels.length + startRow - 1
        }`;
        batchUpdate.push({
          range: levelsHeaderRange,
          values: [upgradeHeaders],
        });
        batchUpdate.push({
          range: levelsRange,
          values: workshopLevels,
        });
      }

      if (enhancementCol > 0 && workshopPlusLevels.length > 0) {
        var plusHeaders = [];
        oldWorkshopPlusLevels.presetNames.forEach(function (presetName) {
          plusHeaders.push(presetName);
        });
        var plusStartCol = shared.columnToLetter(workshopPlusLevelsStartCol);
        var plusEndCol = shared.columnToLetter(workshopPlusLevelsEndCol);
        var plusHeaderRange = `${sheetName}!${plusStartCol}1:${plusEndCol}1`;
        var plusRange = `${sheetName}!${plusStartCol}${startRow}:${plusEndCol}${
          workshopPlusLevels.length + startRow - 1
        }`;
        batchUpdate.push({
          range: plusHeaderRange,
          values: [plusHeaders],
        });
        batchUpdate.push({
          range: plusRange,
          values: workshopPlusLevels,
        });
      }

      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: `Workshop levels updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for workshop levels`,
      };
    } catch (error) {
      console.log(`Error in updateWorkshopLevels: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating workshop levels: ${error.message}`,
      };
    }
  },

  updateWorkshopPlusRatios: function (
    sheetName,
    oldWorkshopPlusRatios,
    desiredRatiosData
  ) {
    try {
      console.log("Called: workshop.updateWorkshopPlusRatios");
      if (!desiredRatiosData || desiredRatiosData.length < 2) {
        console.log(`Not enough data in Desired Ratios sheet`);
        return {
          success: false,
          message: `Not enough data in Desired Ratios sheet`,
        };
      }

      var headerRow = desiredRatiosData[0];
      var workshopEnhancementNameCol = headerRow.indexOf(
        "Workshop Enhancement"
      );
      if (workshopEnhancementNameCol === -1) {
        console.log(`Workshop Enhancement column not found`);
        return {
          success: false,
          message: `Workshop Enhancement column not found`,
        };
      }

      var ratiosToUpdate = [];
      for (var i = 1; i < desiredRatiosData.length; i++) {
        var row = desiredRatiosData[i];
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
        if (enhancementName && oldWorkshopPlusRatios[enhancementName]) {
          var ratioData = oldWorkshopPlusRatios[enhancementName];
          ratiosToUpdate.push([
            ratioData["Order"] || "",
            ratioData["Ratio"] || "",
          ]);
        } else {
          ratiosToUpdate.push(["", ""]);
        }
        if (enhancementName && enhancementName === "Order in between") {
          break;
        }
      }
      if (ratiosToUpdate.length > 0) {
        var startRow = 2;
        var startCol = shared.columnToLetter(workshopEnhancementNameCol + 2);
        var endCol = shared.columnToLetter(workshopEnhancementNameCol + 3);
        var ratiosRange = `${sheetName}!${startCol}${startRow}:${endCol}${
          startRow + ratiosToUpdate.length - 1
        }`;

        return {
          success: true,
          message: `Workshop Plus Ratios updated successfully`,
          batchUpdate: [
            {
              range: ratiosRange,
              values: ratiosToUpdate,
            },
          ],
        };
      }
      return {
        success: true,
        message: `No updates needed for Workshop Plus Ratios`,
      };
    } catch (error) {
      console.log(`Error in updateWorkshopPlusRatios: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating Workshop Plus Ratios: ${error.message}`,
      };
    }
  },

  version20: function () {
    try {
      console.log("Called: workshop.version20");
      var oldSpreadsheet = spreadsheets("Workshop oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old workshop spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet™ not found in old workshop spreadsheet™",
        };
      }

      var workshopLevelsRange = "EXPORT!B2:M";
      var workshopPlusLevelsRange = "EXPORT!P2:V";
      var valuesRanges = [workshopLevelsRange, workshopPlusLevelsRange];

      var updateWorkshopValuesBatchResult = SheetsAPI.batchGetValues(
        oldSheetID,
        valuesRanges
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
        formulasRanges
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

      // Process workshop levels
      var workshopLevelsResult = this.getVersion20WorkshopLevels(
        oldWorkshopLevelsValues
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      // Process workshop plus levels
      var workshopPlusLevelsResult = this.getVersion20WorkshopPlusLevels(
        oldWorkshopPlusLevelsValues
      );
      if (!workshopPlusLevelsResult || !workshopPlusLevelsResult.success) {
        return workshopPlusLevelsResult;
      }

      // Process workshop plus ratios
      var workshopPlusRatiosResult = this.getVersion20WorkshopPlusRatios(
        workshopPlusLevelsResult.oldWorkshopPlusLevels.presetNames,
        oldWorkshopPlusRatiosValues
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
      console.log("Error in version20: " + error.toString());
      return {
        success: false,
        message: "Error in version20: " + error.message,
      };
    }
  },

  version10: function () {
    try {
      console.log("Called: workshop.version10");
      var oldSpreadsheet = spreadsheets("Workshop oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old workshop spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet™ not found in old workshop spreadsheet™",
        };
      }

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

      // Process workshop levels
      var workshopLevelsResult = this.getVersion10WorkshopLevels(
        oldWorkshopLevelsValues
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      // Process workshop plus levels
      var workshopPlusLevelsResult = this.getVersion10WorkshopPlusLevels(
        oldWorkshopPlusLevelsValues
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
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10WorkshopLevels: function (oldWorkshopLevelsValues) {
    try {
      console.log("Called: workshop.getVersion10WorkshopLevels");
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
      console.log("Error in getVersion10WorkshopLevels: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10WorkshopLevels: " + error.message,
      };
    }
  },

  getVersion10WorkshopPlusLevels: function (oldWorkshopPlusLevelsValues) {
    try {
      console.log("Called: workshop.getVersion10WorkshopPlusLevels");
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
      console.log(
        "Error in getVersion10WorkshopPlusLevels: " + error.toString()
      );
      return {
        success: false,
        message: "Error in getVersion10WorkshopPlusLevels: " + error.message,
      };
    }
  },

  getVersion20WorkshopLevels: function (oldWorkshopLevelsValues) {
    try {
      console.log("Called: workshop.getVersion20WorkshopLevels");
      var oldWorkshopLevelsHeaders = oldWorkshopLevelsValues[0];
      var oldWorkshopLevelsPresetNames = [];
      oldWorkshopLevelsHeaders.forEach(function (name, index) {
        var presetName = name && name.trim() !== "" ? name.trim() : null;
        if (presetName && presetName.startsWith("Preset")) {
          oldWorkshopLevelsPresetNames.push(null);
        } else if (presetName && index > 1) {
          oldWorkshopLevelsPresetNames.push(presetName);
        }
      });
      var oldWorkshopLevels = {
        presetNames: oldWorkshopLevelsPresetNames,
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
          oldWorkshopLevels.data[row[1]] = {
            unlocked: row[0] || "",
            levels: [],
          };
          row.forEach(function (cell, index) {
            var presetName = oldWorkshopLevelsHeaders[index];
            if (
              presetName &&
              presetName.trim() !== "" &&
              (oldWorkshopLevelsPresetNames.includes(presetName) ||
                presetName.includes("Preset"))
            ) {
              if (
                presetName === "Tourney" &&
                row[index] === row[index - 2] &&
                (!row[index + 1] || row[index + 1] === row[index - 1])
              ) {
                oldWorkshopLevels.data[row[1]].levels.push(null, null);
              } else {
                oldWorkshopLevels.data[row[1]].levels.push(
                  row[index] || null,
                  row[index + 1] || null
                );
              }
            }
          });
        }
      });
      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: oldWorkshopLevels,
      };
    } catch (error) {
      console.log("Error in getVersion20WorkshopLevels: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion20WorkshopLevels: " + error.message,
      };
    }
  },

  getVersion20WorkshopPlusLevels: function (oldWorkshopPlusLevelsValues) {
    try {
      console.log("Called: workshop.getVersion20WorkshopPlusLevels");
      var oldWorkshopPlusLevelsHeaders = oldWorkshopPlusLevelsValues[0];
      var oldWorkshopPlusPresetNames = [];
      oldWorkshopPlusLevelsHeaders.forEach(function (name, index) {
        var presetName = name && name.trim() !== "" ? name.trim() : null;
        if (presetName && presetName.startsWith("Preset")) {
          oldWorkshopPlusPresetNames.push(null);
        } else if (presetName && index > 1) {
          oldWorkshopPlusPresetNames.push(presetName);
        }
      });
      var oldWorkshopPlusLevels = {
        presetNames: oldWorkshopPlusPresetNames,
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
          oldWorkshopPlusLevels.data[row[0]] = [];
          row.forEach(function (cell, index) {
            var presetName = oldWorkshopPlusLevelsHeaders[index];
            if (
              presetName &&
              presetName.trim() !== "" &&
              (oldWorkshopPlusPresetNames.includes(presetName) ||
                presetName.includes("Preset"))
            ) {
              if (presetName === "Tourney" && row[index] === row[index - 1]) {
                oldWorkshopPlusLevels.data[row[0]].push(null);
              } else {
                oldWorkshopPlusLevels.data[row[0]].push(row[index] || null);
              }
            }
          });
        }
      });
      return {
        success: true,
        message: "Workshop plus levels processed successfully",
        oldWorkshopPlusLevels: oldWorkshopPlusLevels,
      };
    } catch (error) {
      console.log(
        "Error in getVersion20WorkshopPlusLevels: " + error.toString()
      );
      return {
        success: false,
        message: "Error in getVersion20WorkshopPlusLevels: " + error.message,
      };
    }
  },

  getVersion20WorkshopPlusRatios: function (
    presetNames,
    oldWorkshopPlusRatiosValues
  ) {
    try {
      console.log("Called: workshop.getVersion20WorkshopPlusRatios");
      var oldWorkshopPlusRatios = {};
      var oldWorkshopPlusRatiosHeaders = oldWorkshopPlusRatiosValues[0];
      var workshopEnhancementNameCol = oldWorkshopPlusRatiosHeaders.indexOf(
        "Workshop Enhancement"
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
          if (
            enhancementName === "Workshop preset" &&
            !presetNames.includes(ratioValue) &&
            /^\d+$/.test(ratioValue)
          ) {
            ratioValue =
              presetNames[Number(ratioValue) - 1] || "Preset " + ratioValue;
          } else if (enhancementName === "Base ratio") {
            ratioValue = ratioValue.replace(/\+/g, "").trim();
          }
          oldWorkshopPlusRatios[enhancementName] = {
            Order: row[workshopEnhancementNameCol + 1] || "",
            Ratio: ratioValue || "",
          };
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
      console.log(
        "Error in getVersion20WorkshopPlusRatios: " + error.toString()
      );
      return {
        success: false,
        message: "Error in getVersion20WorkshopPlusRatios: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
      "v2.0": this.version20.bind(this),
    };
  },

  isCompatibleVersion: function (oldVersion) {
    console.log("Called: workshop.isCompatibleVersion");
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
