const workshop = {
  exportData: function (versionDifference) {
    try {
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
        data: {
          oldWorkshopLevels: oldDataResult.oldWorkshopLevels || [],
          oldWorkshopPlusLevels: oldDataResult.oldWorkshopPlusLevels || []
        }
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
      // Use sheet type-based naming for parallel execution support
      var newSpreadsheet = spreadsheets("Workshop newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }

      var oldWorkshopLevels = data.oldWorkshopLevels || [];
      var oldWorkshopPlusLevels = data.oldWorkshopPlusLevels || [];

      var requiredRanges = ["Master Sheet", "IDS"];
      var batchResults = SheetsAPI.batchGetFormulas(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var idsData = batchResults[1].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(newSheetID, "IDS", "IDS Master's", idsData);
      if (!newSheetInfo || !newSheetInfo.importStatus || !newSheetInfo.importStatus.range) {
        console.log(`Could not find import status range in IDS sheet`);
        return {
          success: false,
          message: "Could not find import status range in IDS sheet",
        };
      }

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

      var batchUpdate = workshopResult.batchUpdate || [];

      // Add import status update to batch
      batchUpdate.push({
        range: newSheetInfo.importStatus.range,
        values: [["✅"]],
      });

      var updateResult = SheetsAPI.batchUpdateValues(
        newSheetID,
        batchUpdate
      );
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
      if (!masterSheetData || masterSheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: `Not enough data in Master Sheet`,
        };
      }

      var headerRow = masterSheetData[0];

      var upgradeCol = headerRow.indexOf("Workshop Upgrade") + 1;
      var enhancementCol = headerRow.indexOf("Workshop Enhancement") + 1;

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

      for (var i = 1; i < masterSheetData.length; i++) {
        var row = masterSheetData[i];
        
        var workshopName = row[upgradeCol - 1];
        if (workshopName && workshopName.startsWith("=")) {
          workshopName = workshopName.split(",")[1].trim().replace(/['"()]/g, '');
        }
        if (workshopName && oldWorkshopLevels[workshopName]) {
          var data = oldWorkshopLevels[workshopName];
          workshopUnlocked.push([data[0]]);
          workshopLevels.push([data[1], data[2]]);
        }

        var enhancementName = row[enhancementCol - 1];
        if (enhancementName && enhancementName.startsWith("=")) {
          var parts = enhancementName.split(",");
          if (parts.length === 4) {
            enhancementName = parts[2].trim();
          } else {
            enhancementName = parts[parts.length - 1].trim();
          }
          enhancementName = enhancementName.replace(/['"()]/g, '');
        }
        if (enhancementName && oldWorkshopPlusLevels[enhancementName]) {
          var level = oldWorkshopPlusLevels[enhancementName];
          workshopPlusLevels.push([level]);
        }
      }

      var batchUpdate = [];

      if (upgradeCol > 1 && workshopUnlocked.length) {
        var unlockedCol = shared.columnToLetter(upgradeCol - 1);
        var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
          workshopUnlocked.length + 1
        }`;
        batchUpdate.push({
          range: unlockedRange,
          values: workshopUnlocked,
        });
      }

      if (upgradeCol > 0 && workshopLevels.length) {
        var levelsStartCol = shared.columnToLetter(upgradeCol + 1);
        var levelsEndCol = shared.columnToLetter(upgradeCol + 2);
        var levelsRange = `${sheetName}!${levelsStartCol}2:${levelsEndCol}${
          workshopLevels.length + 1
        }`;
        batchUpdate.push({
          range: levelsRange,
          values: workshopLevels,
        });
      }

      if (enhancementCol > 0 && workshopPlusLevels.length) {
        var plusCol = shared.columnToLetter(enhancementCol + 2);
        var plusRange = `${sheetName}!${plusCol}2:${plusCol}${
          workshopPlusLevels.length + 1
        }`;
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

  version10: function () {
    try {
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
        workshopPlusLevelsRange
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

      return this.getVersion10Values(oldWorkshopLevelsValues, oldWorkshopPlusLevelsValues);
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10Values: function (oldWorkshopLevelsValues, oldWorkshopPlusLevelsValues) {
    try {
      var oldWorkshopLevels = {};
      oldWorkshopLevelsValues.forEach(function(row) {
        var hasData = row.some(function(cell) {
          return cell !== null && cell !== undefined && String(cell || "").trim() !== "";
        });
        if (hasData && row[1]) {
          oldWorkshopLevels[row[1]] = [row[0] || "", row[2] || "", row[3] || ""];
        }
      });

      var oldWorkshopPlusLevels = {};
      oldWorkshopPlusLevelsValues.forEach(function(row) {
        var hasData = row.some(function(cell) {
          return cell !== null && cell !== undefined && String(cell || "").trim() !== "";
        });
        if (hasData && row[0]) {
          oldWorkshopPlusLevels[row[0]] = row[2] || "";
        }
      });

      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: oldWorkshopLevels,
        oldWorkshopPlusLevels: oldWorkshopPlusLevels,
      };
    } catch (error) {
      console.log("Error in getVersion10Values: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10Values: " + error.message,
      };
    }
  },

  version19: function () {
    try {
      var oldSpreadsheet = spreadsheets("Workshop oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      
      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old workshop spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet™ not found in old workshop spreadsheet™",
        };
      }

      // Update ranges for presets
      var workshopLevelsRange = "EXPORT!B3:F";
      var workshopPlusLevelsRange = "EXPORT!H3:K";

      var updateWorkshopBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        workshopLevelsRange,
        workshopPlusLevelsRange
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

      return this.getVersion19Values(oldWorkshopLevelsValues, oldWorkshopPlusLevelsValues);
    } catch (error) {
      console.log("Error in version19: " + error.toString());
      return {
        success: false,
        message: "Error in version19: " + error.message,
      };
    }
  },

  getVersion19Values: function (oldWorkshopLevelsValues, oldWorkshopPlusLevelsValues) {
    try {
      var oldWorkshopLevels = {};
      oldWorkshopLevelsValues.forEach(function(row) {
        var hasData = row.some(function(cell) {
          return cell !== null && cell !== undefined && String(cell || "").trim() !== "";
        });
        if (hasData && row[1]) {
          // Update data for presets
          oldWorkshopLevels[row[1]] = [row[0] || "", row[2] || "", row[3] || ""];
        }
      });

      var oldWorkshopPlusLevels = {};
      oldWorkshopPlusLevelsValues.forEach(function(row) {
        var hasData = row.some(function(cell) {
          return cell !== null && cell !== undefined && String(cell || "").trim() !== "";
        });
        if (hasData && row[0]) {
          // Update data for presets
          oldWorkshopPlusLevels[row[0]] = row[2] || "";
        }
      });

      return {
        success: true,
        message: "Workshop levels processed successfully",
        oldWorkshopLevels: oldWorkshopLevels,
        oldWorkshopPlusLevels: oldWorkshopPlusLevels,
      };
    } catch (error) {
      console.log("Error in getVersion19Values: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion19Values: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
      // "v1.9": this.version19.bind(this),
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
