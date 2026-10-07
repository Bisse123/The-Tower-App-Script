const modulesReader = {

  /**
   * Reads Modules data from a v6.4.3 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version6_4_3: function (oldSheetID) {
    try {
      console.log("Called: modulesReader.version6_4_3");

      var ranges = ["Inventory", "Presets", "Planner v2", "Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length === 0) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesPlannerValues = batchResult[2].values;
      var oldModulesTrackerValues = batchResult[3].values;

      var formulaRanges = ["Tracker"];
      var formulaBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulaRanges,
      );
      if (!formulaBatchResult || formulaBatchResult.length === 0) {
        console.log("Could not read module formulas from old spreadsheet");
        return {
          success: false,
          message: "Could not read module formulas from old spreadsheet",
        };
      }
      var oldModulesTrackerFormulas = formulaBatchResult[0].values;

      var inventoryData = modulesInventoryReader.getVersion5_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = modulesPresetsReader.getVersion5_0ModulesPresets(
        oldModulesPresetsValues,
      );
      var plannerData = modulesPlannerReader.getVersion6_4_3ModulesPlanner(
        oldModulesPlannerValues,
        inventoryData.oldModulesInventory,
      );
      var trackerData = modulesTrackerReader.getVersion6_4_3ModulesTracker(
        oldModulesTrackerValues,
        oldModulesTrackerFormulas,
      );

      var success =
        inventoryData.success &&
        presetsData.success &&
        plannerData.success &&
        trackerData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesPlanner: plannerData.oldModulesPlanner || {},
        oldModulesTracker: trackerData.oldModulesTracker || {},
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version6_4_3", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Modules data from a v5.2.1 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_2_1: function (oldSheetID) {
    try {
      console.log("Called: modulesReader.version5_2_1");

      var ranges = ["Inventory", "Presets", "Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length === 0) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesTrackerValues = batchResult[2].values;

      var formulaRanges = ["Tracker"];
      var formulaBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulaRanges,
      );
      if (!formulaBatchResult || formulaBatchResult.length === 0) {
        console.log("Could not read module formulas from old spreadsheet");
        return {
          success: false,
          message: "Could not read module formulas from old spreadsheet",
        };
      }
      var oldModulesTrackerFormulas = formulaBatchResult[0].values;

      var inventoryData = modulesInventoryReader.getVersion5_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = modulesPresetsReader.getVersion5_0ModulesPresets(
        oldModulesPresetsValues,
      );
      var trackerData = modulesTrackerReader.getVersion4_7ModulesTracker(
        oldModulesTrackerValues,
        oldModulesTrackerFormulas,
      );

      var success =
        inventoryData.success && presetsData.success && trackerData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesTracker: trackerData.oldModulesTracker || {},
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version5_2_1", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Modules data from a v5.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version5_0: function (oldSheetID) {
    try {
      console.log("Called: modulesReader.version5_0");

      var ranges = ["Modules Inventory", "Modules Presets", "Modules Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length === 0) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesTrackerValues = batchResult[2].values;

      var formulaRanges = ["Tracker"];
      var formulaBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulaRanges,
      );
      if (!formulaBatchResult || formulaBatchResult.length === 0) {
        console.log("Could not read module formulas from old spreadsheet");
        return {
          success: false,
          message: "Could not read module formulas from old spreadsheet",
        };
      }
      var oldModulesTrackerFormulas = formulaBatchResult[0].formulas;

      var inventoryData = modulesInventoryReader.getVersion5_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = modulesPresetsReader.getVersion5_0ModulesPresets(
        oldModulesPresetsValues,
      );
      var trackerData = modulesTrackerReader.getVersion4_7ModulesTracker(
        oldModulesTrackerValues,
        oldModulesTrackerFormulas,
      );

      var success =
        inventoryData.success && presetsData.success && trackerData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesTracker: trackerData.oldModulesTracker || {},
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version5_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Modules data from a v4.7 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_7: function (oldSheetID) {
    try {
      console.log("Called: modulesReader.version4_7");

      var ranges = ["Modules Inventory", "Modules Presets", "Modules Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length < 2) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesTrackerValues = batchResult[2].values;

      var formulaRanges = ["Tracker"];
      var formulaBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulaRanges,
      );
      if (!formulaBatchResult || formulaBatchResult.length === 0) {
        console.log("Could not read module formulas from old spreadsheet");
        return {
          success: false,
          message: "Could not read module formulas from old spreadsheet",
        };
      }
      var oldModulesTrackerFormulas = formulaBatchResult[0].formulas;

      var inventoryData = modulesInventoryReader.getVersion4_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = modulesPresetsReader.getVersion4_0ModulesPresets(
        oldModulesPresetsValues,
      );
      var trackerData = modulesTrackerReader.getVersion4_7ModulesTracker(
        oldModulesTrackerValues,
        oldModulesTrackerFormulas,
      );

      var success =
        inventoryData.success && presetsData.success && trackerData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesTracker: trackerData.oldModulesTracker || {},
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version4_7", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Modules data from a v4.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_0: function (oldSheetID) {
    try {
      console.log("Called: modulesReader.version4_0");

      var ranges = ["Modules Inventory", "Modules Presets"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length < 2) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;

      var inventoryData = modulesInventoryReader.getVersion4_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = modulesPresetsReader.getVersion4_0ModulesPresets(
        oldModulesPresetsValues,
      );

      var success = inventoryData.success && presetsData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version4_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },
};
