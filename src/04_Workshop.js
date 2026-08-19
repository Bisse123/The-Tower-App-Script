const workshop = {
  // #region Export Functions
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

  // #endregion
  // #region Import Functions
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

      var requiredFormulaRanges = ["Master Sheet"];
      var batchFormulaResults = SheetsAPI.batchGetFormulas(
        newSheetID,
        requiredFormulaRanges,
      );
      if (!batchFormulaResults || batchFormulaResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var requiredValueRanges = ["Desired Ratios", "IDS"];
      var batchValueResults = SheetsAPI.batchGetValues(
        newSheetID,
        requiredValueRanges,
      );
      if (!batchValueResults || batchValueResults.length < 2) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchFormulaResults[0].values;

      var desiredRatiosData = batchValueResults[0].values;
      var idsData = batchValueResults[1].values;

      var batchUpdate = [];

      // Only update workshop levels if key exists
      if (
        data.hasOwnProperty("oldWorkshopLevels") &&
        data.hasOwnProperty("oldWorkshopPlusLevels")
      ) {
        var oldWorkshopLevels = data.oldWorkshopLevels;
        var oldWorkshopPlusLevels = data.oldWorkshopPlusLevels;
        var hasPresets = data.hasOwnProperty("hasPresets") ? data.hasPresets : true;
        var workshopResult = this.updateWorkshopLevels(
          "Master Sheet",
          oldWorkshopLevels,
          oldWorkshopPlusLevels,
          hasPresets,
          masterSheetData,
        );
        if (!workshopResult || !workshopResult.success) {
          console.log(
            `Error updating workshop levels: ${workshopResult.message}`,
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
          desiredRatiosData,
        );
        if (!ratioResult || !ratioResult.success) {
          console.log(
            `Error updating workshop plus ratios: ${ratioResult.message}`,
          );
          return ratioResult;
        }
        batchUpdate = batchUpdate.concat(ratioResult.batchUpdate || []);
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Workshop",
        newSheetID,
        idsData,
        data.idMasterID,
      );

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

  // #endregion
  // #region Update Functions
  updateWorkshopLevels: function (
    sheetName,
    oldWorkshopLevels,
    oldWorkshopPlusLevels,
    hasPresets,
    masterSheetData,
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
        if (workshopName && oldWorkshopLevels.data.hasOwnProperty(workshopName)) {
          var oldWorkshopRowData = oldWorkshopLevels.data[workshopName];
          workshopUnlocked.push([oldWorkshopRowData.unlocked || ""]);
          workshopLevels.push(oldWorkshopRowData.levels || []);
        } else {
          workshopUnlocked.push([""]);
          workshopLevels.push([]);
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
        if (enhancementName && oldWorkshopPlusLevels.data.hasOwnProperty(enhancementName)) {
          var enhancementData = oldWorkshopPlusLevels.data[enhancementName];
          workshopPlusLevels.push(enhancementData);
        } else {
          workshopPlusLevels.push([]);
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
        var levelsRange = `${sheetName}!${levelsStartCol}${startRow}:${levelsEndCol}${
          workshopLevels.length + startRow - 1
        }`;
        batchUpdate.push({
          range: levelsRange,
          values: workshopLevels,
        });
        if (hasPresets) {
          var levelsHeaderRange = `${sheetName}!${levelsStartCol}1:${levelsEndCol}1`;
          batchUpdate.push({
            range: levelsHeaderRange,
            values: [upgradeHeaders],
          });
        }
      }

      if (enhancementCol > 0 && workshopPlusLevels.length > 0) {
        var plusHeaders = [];
        oldWorkshopPlusLevels.presetNames.forEach(function (presetName) {
          plusHeaders.push(presetName);
        });
        var plusStartCol = shared.columnToLetter(workshopPlusLevelsStartCol);
        var plusEndCol = shared.columnToLetter(workshopPlusLevelsEndCol);
        var plusRange = `${sheetName}!${plusStartCol}${startRow}:${plusEndCol}${
          workshopPlusLevels.length + startRow - 1
        }`;
        batchUpdate.push({
          range: plusRange,
          values: workshopPlusLevels,
        });
        if (hasPresets) {
          var plusHeaderRange = `${sheetName}!${plusStartCol}1:${plusEndCol}1`;
          batchUpdate.push({
            range: plusHeaderRange,
            values: [plusHeaders],
          });
        }
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
    desiredRatiosData,
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
        "Workshop Enhancement",
      );

      if (workshopEnhancementNameCol === -1) {
        console.log(`Workshop Enhancement column not found`);
        return {
          success: false,
          message: `Workshop Enhancement column not found`,
        };
      }

      var ratiosToUpdate = [];
      var batchUpdate = [];
      for (var i = 1; i < desiredRatiosData.length; i++) {
        var row = desiredRatiosData[i];
        var workshopPresetIndex = row.indexOf("Workshop preset");
        if (workshopPresetIndex !== -1) {
          var presetValue = oldWorkshopPlusRatios["Workshop preset"] || null;
          var presetCol = shared.columnToLetter(workshopPresetIndex + 2);
          var presetRange = `${sheetName}!${presetCol}${i + 1}`;
          batchUpdate.push({
            range: presetRange,
            values: [[presetValue]],
          });
        }
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
          ratiosToUpdate.push(ratioData);
        } else {
          ratiosToUpdate.push(["", "", "", "", "", "", "", "", "", ""]);
        }
        if (enhancementName && enhancementName === "Order in between") {
          break;
        }
      }
      if (ratiosToUpdate.length > 0) {
        var startRow = 2;
        var startCol = shared.columnToLetter(workshopEnhancementNameCol + 2);
        var endCol = shared.columnToLetter(workshopEnhancementNameCol + 11);
        var ratiosRange = `${sheetName}!${startCol}${startRow}:${endCol}${
          startRow + ratiosToUpdate.length - 1
        }`;
        batchUpdate.push({
          range: ratiosRange,
          values: ratiosToUpdate,
        });
      }

      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: `Workshop Plus Ratios updated successfully`,
          batchUpdate: batchUpdate,
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

  // #endregion
  // #region Convert Versions
  version2_2_8: function () {
    try {
      console.log("Called: workshop.version2_2_8");
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

      // Process workshop levels
      var workshopLevelsResult = this.getVersion2_0WorkshopLevels(
        oldWorkshopLevelsValues,
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      // Process workshop plus levels
      var workshopPlusLevelsResult = this.getVersion2_0WorkshopPlusLevels(
        oldWorkshopPlusLevelsValues,
      );
      if (!workshopPlusLevelsResult || !workshopPlusLevelsResult.success) {
        return workshopPlusLevelsResult;
      }

      // Process workshop plus ratios
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
      console.log("Error in version2_2_8: " + error.toString());
      return {
        success: false,
        message: "Error in version2_2_8: " + error.message,
      };
    }
  },

  version2_1: function () {
    try {
      console.log("Called: workshop.version2_1");
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

      // Process workshop levels
      var workshopLevelsResult = this.getVersion2_0WorkshopLevels(
        oldWorkshopLevelsValues,
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      // Process workshop plus levels
      var workshopPlusLevelsResult = this.getVersion2_0WorkshopPlusLevels(
        oldWorkshopPlusLevelsValues,
      );
      if (!workshopPlusLevelsResult || !workshopPlusLevelsResult.success) {
        return workshopPlusLevelsResult;
      }

      // Process workshop plus ratios
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
      console.log("Error in version2_1: " + error.toString());
      return {
        success: false,
        message: "Error in version2_1: " + error.message,
      };
    }
  },

  version2_0: function () {
    try {
      console.log("Called: workshop.version2_0");
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

      // Process workshop levels
      var workshopLevelsResult = this.getVersion2_0WorkshopLevels(
        oldWorkshopLevelsValues,
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      // Process workshop plus levels
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
      console.log("Error in version2_0: " + error.toString());
      return {
        success: false,
        message: "Error in version2_0: " + error.message,
      };
    }
  },

  version1_0: function () {
    try {
      console.log("Called: workshop.version1_0");
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
      var workshopLevelsResult = this.getVersion1_0WorkshopLevels(
        oldWorkshopLevelsValues,
      );
      if (!workshopLevelsResult || !workshopLevelsResult.success) {
        return workshopLevelsResult;
      }

      // Process workshop plus levels
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
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Workshop Levels
  getVersion2_0WorkshopLevels: function (oldWorkshopLevelsValues) {
    try {
      console.log("Called: workshop.getVersion2_0WorkshopLevels");
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

      var presetOrder = shared.resolvePresetOrder(
        presetColumns.map(function (column) {
          return column.presetName;
        }),
        shared.templatePresetNames,
      );
      var orderedColumns = presetOrder.indices.map(function (sourceIndex) {
        return presetColumns[sourceIndex];
      });
      var defaultColIndex = orderedColumns.length
        ? orderedColumns[0].colIndex
        : -1;

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
            if (
              index !== defaultColIndex &&
              row[index] === row[defaultColIndex] &&
              (!row[index + 1] || row[index + 1] === row[defaultColIndex + 1])
            ) {
              levels.push(null, null);
            } else {
              levels.push(row[index] || null, row[index + 1] || null);
            }
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
      console.log("Error in getVersion2_0WorkshopLevels: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion2_0WorkshopLevels: " + error.message,
      };
    }
  },

  getVersion1_0WorkshopLevels: function (oldWorkshopLevelsValues) {
    try {
      console.log("Called: workshop.getVersion1_0WorkshopLevels");
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
      console.log("Error in getVersion1_0WorkshopLevels: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0WorkshopLevels: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Workshop Plus Levels
  getVersion2_0WorkshopPlusLevels: function (oldWorkshopPlusLevelsValues) {
    try {
      console.log("Called: workshop.getVersion2_0WorkshopPlusLevels");
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

      var presetOrder = shared.resolvePresetOrder(
        presetColumns.map(function (column) {
          return column.presetName;
        }),
        shared.templatePresetNames,
      );
      var orderedColumns = presetOrder.indices.map(function (sourceIndex) {
        return presetColumns[sourceIndex];
      });
      var defaultColIndex = orderedColumns.length
        ? orderedColumns[0].colIndex
        : -1;

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
            if (index !== defaultColIndex && row[index] === row[defaultColIndex]) {
              levels.push(null);
            } else {
              levels.push(row[index] || null);
            }
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
      console.log(
        "Error in getVersion2_0WorkshopPlusLevels: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion2_0WorkshopPlusLevels: " + error.message,
      };
    }
  },

  getVersion1_0WorkshopPlusLevels: function (oldWorkshopPlusLevelsValues) {
    try {
      console.log("Called: workshop.getVersion1_0WorkshopPlusLevels");
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
        "Error in getVersion1_0WorkshopPlusLevels: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion1_0WorkshopPlusLevels: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Workshop Plus Ratios
  getVersion2_2_8WorkshopPlusRatios: function (
    presetNames,
    oldWorkshopPlusRatiosValues,
  ) {
    try {
      console.log("Called: workshop.getVersion2_2_8WorkshopPlusRatios");
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
      console.log(
        "Error in getVersion2_2_8WorkshopPlusRatios: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion2_2_8WorkshopPlusRatios: " + error.message,
      };
    }
  },

  getVersion2_1WorkshopPlusRatios: function (
    presetNames,
    oldWorkshopPlusRatiosValues,
  ) {
    try {
      console.log("Called: workshop.getVersion2_1WorkshopPlusRatios");
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
      console.log(
        "Error in getVersion2_1WorkshopPlusRatios: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion2_1WorkshopPlusRatios: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseWorkshopData: function (data) {
    const attackUpgradeNamesByIndex = {
      0: "Damage",
      1: "Attack Speed",
      2: "Critical Chance",
      3: "Critical Factor",
      4: "Range",
      5: "Damage / Meter",
      6: "Multishot Chance",
      7: "Multishot Targets",
      8: "Rapid Fire Chance",
      9: "Rapid Fire Duration",
      10: "Bounce Shot Chance",
      11: "Bounce Shot Targets",
      12: "Bounce Shot Range",
      13: "Super Critical Chance",
      14: "Super Critical Mult",
      15: "Rend Armor Chance",
      16: "Rend Armor Mult",
    };
    const defenseUpgradeNamesByIndex = {
      0: "Health",
      1: "Health Regen",
      2: "Defense %",
      3: "Defense Absolute",
      4: "Thorn Damage",
      5: "Lifesteal",
      6: "Knockback Chance",
      7: "Knockback Force",
      8: "Orb Speed",
      9: "Orbs",
      10: "Shockwave Size",
      11: "Shockwave Frequency",
      12: "Land Mine Chance",
      13: "Land Mine Damage",
      14: "Land Mine Radius",
      15: "Death Defy",
      16: "Wall Health",
      17: "Wall Rebuild",
    };
    const utilityUpgradeNamesByIndex = {
      0: "Cash Bonus",
      1: "Cash / Wave",
      2: "Coin / Kill Bonus",
      3: "Coin / Wave",
      4: "Free Attack Upgrade",
      5: "Free Defense Upgrade",
      6: "Free Utility Upgrade",
      7: "Interest / Wave",
      8: "Recovery Amount",
      9: "Max Amount",
      10: "Package Chance",
      11: "Enemy Attack Level Skip",
      12: "Enemy Health Level Skip",
    };
    const attackEnhancementNamesByIndex = {
      0: "Damage +",
      1: "Rend Armor Mult +",
      2: "Critical Factor +",
      3: "Damage / Meter +",
      4: "Super Crit Multi +",
      5: "Attack Speed +",
    };
    const defenseEnhancementNamesByIndex = {
      0: "Health +",
      1: "Health Regen +",
      2: "Defense Absolute +",
      3: "Land Mine Damage +",
      4: "Wall Health +",
      5: "Orb Size +",
    };
    const utilityEnhancementNamesByIndex = {
      0: "Cash Bonus +",
      1: "Coin Bonus +",
      2: "Cells / Kill Bonus +",
      3: "Free Upgrades +",
      4: "Recovery Package +",
      5: "Enemy Level Skips +",
    };
    const attackUpgradeUnlockedNamesByIndex = {
      0: "Range",
      1: "Multishot Chance",
      2: "Rapid Fire Chance",
      3: "Bounce Shot Chance",
      4: "Super Critical Chance",
      5: "Rend Armor Chance",
    };
    const defenseUpgradeUnlockedNamesByIndex = {
      0: "Defense %",
      1: "Thorn Damage",
      2: "Lifesteal",
      3: "Knockback Chance",
      4: "Orb Speed",
      5: "Shockwave Size",
      6: "Land Mine Chance",
      7: "Death Defy",
      8: "Wall Health",
    };
    const utilityUpgradeUnlockedNamesByIndex = {
      0: "Cash Bonus",
      1: "Coin / Kill Bonus",
      2: "Free Attack Upgrade",
      3: "Interest / Wave",
      4: "Recovery Amount",
      5: "Enemy Attack Level Skip",
    };

    function namesByIndexToArray(namesByIndex) {
      var array = [];
      Object.keys(namesByIndex).forEach(function (index) {
        array[Number(index)] = namesByIndex[index];
      });
      return array;
    }

    const attackUpgradeIndices = namesByIndexToArray(attackUpgradeNamesByIndex);
    const defenseUpgradeIndices = namesByIndexToArray(defenseUpgradeNamesByIndex);
    const utilityUpgradeIndices = namesByIndexToArray(utilityUpgradeNamesByIndex);
    const attackEnhancementIndices = namesByIndexToArray(attackEnhancementNamesByIndex);
    const defenseEnhancementIndices = namesByIndexToArray(defenseEnhancementNamesByIndex);
    const utilityEnhancementIndices = namesByIndexToArray(utilityEnhancementNamesByIndex);
    const attackUpgradeUnlockedIndices = namesByIndexToArray(attackUpgradeUnlockedNamesByIndex);
    const defenseUpgradeUnlockedIndices = namesByIndexToArray(defenseUpgradeUnlockedNamesByIndex);
    const utilityUpgradeUnlockedIndices = namesByIndexToArray(utilityUpgradeUnlockedNamesByIndex);

    const presetOrder = shared.resolvePresetOrder(
      data.presetNames || [],
      shared.templatePresetNames,
    );
    var presetNames = presetOrder.order;
    const presetIndices = presetOrder.indices;

    const attackUpgradeData = data.upgradeAttackLevels || [];
    const defenseUpgradeData = data.upgradeDefenseLevels || [];
    const utilityUpgradeData = data.upgradeUtilityLevels || [];

    const attackEnhancementData = data.enhancementAttackLevels || [];
    const defenseEnhancementData = data.enhancementDefenseLevels || [];
    const utilityEnhancementData = data.enhancementUtilityLevels || [];

    const attackUpgradeMax = attackUpgradeData.length;
    const defenseUpgradeMax = defenseUpgradeData.length;
    const utilityUpgradeMax = utilityUpgradeData.length;
    const attackEnhancementMax = attackEnhancementData.length;
    const defenseEnhancementMax = defenseEnhancementData.length;
    const utilityEnhancementMax = utilityEnhancementData.length;

    var attackPresetUpgradeData = data.presetUpgradeAttackLevels || [];
    var defensePresetUpgradeData = data.presetUpgradeDefenseLevels || [];
    var utilityPresetUpgradeData = data.presetUpgradeUtilityLevels || [];

    var attackPresetEnhancementData = data.presetEnhancementAttackLevels || [];
    var defensePresetEnhancementData = data.presetEnhancementDefenseLevels || [];
    var utilityPresetEnhancementData = data.presetEnhancementUtilityLevels || [];

    const attackUpgradeUnlocked = data.upgradeAttackUnlocked || [];
    const defenseUpgradeUnlocked = data.upgradeDefenseUnlocked || [];
    const utilityUpgradeUnlocked = data.upgradeUtilityUnlocked || [];

    const attackEnhancementUnlocked = data.enhancementAttackUnlocked || [];
    const defenseEnhancementUnlocked = data.enhancementDefenseUnlocked || [];
    const utilityEnhancementUnlocked = data.enhancementUtilityUnlocked || [];

    var attackPresetUpgradeUnlocked = data.presetUpgradeAttackUnlocked || [];
    var defensePresetUpgradeUnlocked = data.presetUpgradeDefenseUnlocked || [];
    var utilityPresetUpgradeUnlocked = data.presetUpgradeUtilityUnlocked || [];

    const hasPresets =
      attackPresetUpgradeUnlocked.some((unlocked) => unlocked) ||
      defensePresetUpgradeUnlocked.some((unlocked) => unlocked) ||
      utilityPresetUpgradeUnlocked.some((unlocked) => unlocked);
    
    if (!hasPresets) {
      attackPresetUpgradeData = attackUpgradeData;
      defensePresetUpgradeData = defenseUpgradeData;
      utilityPresetUpgradeData = utilityUpgradeData;
      attackPresetEnhancementData = attackEnhancementData;
      defensePresetEnhancementData = defenseEnhancementData;
      utilityPresetEnhancementData = utilityEnhancementData;
      attackPresetUpgradeUnlocked = attackUpgradeUnlocked;
      defensePresetUpgradeUnlocked = defenseUpgradeUnlocked;
      utilityPresetUpgradeUnlocked = utilityUpgradeUnlocked;
    }

    var oldWorkshopLevels = {
      presetNames: presetNames,
      data: {},
    };
    var oldWorkshopPlusLevels = {
      presetNames: presetNames,
      data: {},
    };

    function populateUpgradeLevels(upgradeIndices, upgradeData, presetUpgradeIndex, upgradeUnlockIndices, upgradeUnlockedData) {
      upgradeIndices.forEach((upgradeName, upgradeIndex) => {
        if (!oldWorkshopLevels.data[upgradeName]) {
          oldWorkshopLevels.data[upgradeName] = {
            unlocked: null,
            levels: [],
          };
        }
        const unlockedIndex = upgradeUnlockIndices.findIndex((unlocked) => unlocked === upgradeName);
        if (unlockedIndex !== -1 && !oldWorkshopLevels.data[upgradeName].unlocked) {
          oldWorkshopLevels.data[upgradeName].unlocked = upgradeUnlockedData[presetUpgradeIndex + unlockedIndex];
        }
        const levelValue = upgradeData[presetUpgradeIndex + upgradeIndex];
        oldWorkshopLevels.data[upgradeName].levels.push(levelValue, null);
      })
    }

    function populateEnhancementLevels(enhancementIndices, enhancementData, presetEnhancementIndex) {
      enhancementIndices.forEach((enhancementName, enhancementIndex) => {
        if (!oldWorkshopPlusLevels.data[enhancementName]) {
          oldWorkshopPlusLevels.data[enhancementName] = [];
        }
        const levelValue = enhancementData[presetEnhancementIndex + enhancementIndex];
        oldWorkshopPlusLevels.data[enhancementName].push(levelValue);
      })
    }

    presetIndices.forEach((presetIndex) => {
      const attackUpgradeIndex = presetIndex * attackUpgradeMax;
      const defenseUpgradeIndex = presetIndex * defenseUpgradeMax;
      const utilityUpgradeIndex = presetIndex * utilityUpgradeMax;

      const attackEnhancementIndex = presetIndex * attackEnhancementMax;
      const defenseEnhancementIndex = presetIndex * defenseEnhancementMax;
      const utilityEnhancementIndex = presetIndex * utilityEnhancementMax;

      populateUpgradeLevels(attackUpgradeIndices, attackPresetUpgradeData, attackUpgradeIndex, attackUpgradeUnlockedIndices, attackPresetUpgradeUnlocked);
      populateUpgradeLevels(defenseUpgradeIndices, defensePresetUpgradeData, defenseUpgradeIndex, defenseUpgradeUnlockedIndices, defensePresetUpgradeUnlocked);
      populateUpgradeLevels(utilityUpgradeIndices, utilityPresetUpgradeData, utilityUpgradeIndex, utilityUpgradeUnlockedIndices, utilityPresetUpgradeUnlocked);

      populateEnhancementLevels(attackEnhancementIndices, attackPresetEnhancementData, attackEnhancementIndex);
      populateEnhancementLevels(defenseEnhancementIndices, defensePresetEnhancementData, defenseEnhancementIndex);
      populateEnhancementLevels(utilityEnhancementIndices, utilityPresetEnhancementData, utilityEnhancementIndex);
    })

    return {
      oldWorkshopLevels: oldWorkshopLevels,
      oldWorkshopPlusLevels: oldWorkshopPlusLevels,
      hasPresets: hasPresets,
      upgradeIndices: {
        Attack: attackUpgradeIndices,
        Defense: defenseUpgradeIndices,
        Utility: utilityUpgradeIndices,
      },
      enhancementIndices: {
        Attack: attackEnhancementIndices,
        Defense: defenseEnhancementIndices,
        Utility: utilityEnhancementIndices,
      },
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v2.0": this.version2_0.bind(this),
      "v2.1": this.version2_1.bind(this),
      "v2.2.8": this.version2_2_8.bind(this),
    };
  },

  // #endregion
  // #region Compatibility Check
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

  // #endregion
};
