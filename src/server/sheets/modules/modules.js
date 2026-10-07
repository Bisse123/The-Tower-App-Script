const modules = {

  /**
   * Reads Modules data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: modules.exportData");
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
        message: "Modules export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("modules.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Modules data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: modules.importData");

      var requiredRanges = ["Inventory", "Presets", "Planner v2", "Tracker", "IDS"];
      var dvtIndex = requiredRanges.length;
      var dvtNamedRanges = {
        "Main Efficiency": "DVT_Mod_Assist_Bonus_Level",
        "Substat Efficiency": "DVT_Mod_Assist_Substat_Level",
      };

      Object.keys(dvtNamedRanges).forEach(function (item) {
        requiredRanges.push(dvtNamedRanges[item]);
      });

      var batchUpdate = [];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log("Error getting modules sheet data");
        return {
          success: false,
          message: "Error getting modules sheet data",
        };
      }

      var newModuleInventoryValues = batchResults[0].values;
      var newModulePresetsValues = batchResults[1].values;
      var newModulesPlannerValues = batchResults[2].values;
      var newModulesTrackerValues = batchResults[3].values;
      var idsData = batchResults[4].values;

      var dvtNamedRangesData = {};
      Object.keys(dvtNamedRanges).forEach(function (item) {
        dvtNamedRangesData[item] = batchResults[dvtIndex]
          ? batchResults[dvtIndex].values
          : [];
        dvtIndex++;
      });

      var batchUpdate = [];

      if (data.hasOwnProperty("oldModulesInventory")) {
        var oldModulesInventory = data.oldModulesInventory;
        var inventoryResult = modulesWriter.updateModulesInventory(
          "Inventory",
          oldModulesInventory,
          newModuleInventoryValues,
        );
        if (!inventoryResult || !inventoryResult.success) {
          return {
            success: false,
            message: inventoryResult.message,
          };
        }
        batchUpdate = batchUpdate.concat(inventoryResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldModulesPresets")) {
        var oldModulesPresets = data.oldModulesPresets;
        var presetsResult = modulesWriter.updateModulesPresets(
          "Presets",
          oldModulesPresets,
          newModulePresetsValues,
          dvtNamedRangesData,
        );
        if (!presetsResult || !presetsResult.success) {
          return {
            success: false,
            message: presetsResult.message,
          };
        }
        batchUpdate = batchUpdate.concat(presetsResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldModulesPlanner")) {
        var oldModulesPlanner = data.oldModulesPlanner;
        var PlannerResult = modulesWriter.updateModulesInventory(
          "Planner v2",
          oldModulesPlanner,
          newModulesPlannerValues,
        );
        if (!PlannerResult || !PlannerResult.success) {
          return {
            success: false,
            message: PlannerResult.message,
          };
        }
        batchUpdate = batchUpdate.concat(PlannerResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldModulesTracker")) {
        var oldModulesTracker = data.oldModulesTracker;
        var trackerResult = modulesWriter.updateModulesTracker(
          "Tracker",
          oldModulesTracker,
          newModulesTrackerValues,
        );
        if (!trackerResult || !trackerResult.success) {
          return {
            success: false,
            message: trackerResult.message,
          };
        }
        batchUpdate = batchUpdate.concat(trackerResult.batchUpdate || []);
      }

      labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Modules",
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
        message:
          batchUpdate.length > 1
            ? `Modules data imported successfully`
            : "No modules data to update, but ID setting completed",
      };
    } catch (error) {
      var errorReport = errors.report("modules.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v4.0": modulesReader.version4_0.bind(modulesReader),
      "v4.7": modulesReader.version4_7.bind(modulesReader),
      "v5.0": modulesReader.version5_0.bind(modulesReader),
      "v5.2.1": modulesReader.version5_2_1.bind(modulesReader),
      "v6.4.3": modulesReader.version6_4_3.bind(modulesReader),
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
      return versionUtils.compareVersions(b, a) === "newer" ? 1 : -1;
    });

    for (var i = 0; i < sortedThresholds.length; i++) {
      var threshold = sortedThresholds[i];
      var compareResult = versionUtils.compareVersions(oldVersion, threshold);

      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }

    return null;
  },
};
