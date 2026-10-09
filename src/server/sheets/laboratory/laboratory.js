const lab = {

  /**
   * Reads Laboratory data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: lab.exportData");
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
        message: "Laboratory export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {

      var errorReport = errors.report("lab.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
        oldDataResult: oldDataResult,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Laboratory data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: lab.importData");
      var newSpreadsheet = spreadsheets("Laboratory newSpreadsheet", newSheetID);
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }

      var requiredRanges = ["Master Sheet", "IDS"];
      var labPlannerSheetName = "";
      var batchUpdate = [];

      var labPlannerSheet = SheetsAPI.getSheetBySubstring(
        newSpreadsheet,
        "Lab Planner",
      );
      if (labPlannerSheet) {
        labPlannerSheetName = labPlannerSheet.title;
        requiredRanges.push(labPlannerSheetName);
      }

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
      var labPlannerData = batchResults[2] ? batchResults[2].values : null;

      if (data.hasOwnProperty("oldLabLevels")) {
        var oldLabLevels = data.oldLabLevels;
        var labResult = labWriter.updateLabLevels(
          "Master Sheet",
          oldLabLevels,
          masterSheetData,
        );
        if (!labResult || !labResult.success) {
          console.log(`Error updating lab levels: ${labResult.message}`);
          return labResult;
        }
        batchUpdate = batchUpdate.concat(labResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldLabPlanner")) {
        var oldLabPlanner = data.oldLabPlanner;
        var labPlannerResult = labWriter.updateLabPlanner(
          labPlannerSheetName,
          oldLabPlanner,
          labPlannerData,
        );
        if (!labPlannerResult || !labPlannerResult.success) {
          console.log(
            `Error updating lab planner: ${labPlannerResult.message}`,
          );
          return labPlannerResult;
        }
        batchUpdate = batchUpdate.concat(labPlannerResult.batchUpdate || []);
      }

      batchUpdate = labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Laboratory",
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
        message: `Laboratory import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("lab.importData", error, {
        newSheetID: newSheetID,
        data: data,
        requiredRanges: requiredRanges,
        batchUpdate: batchUpdate,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": labReader.version1_0.bind(labReader),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
  isCompatibleVersion: function (oldVersion) {
    console.log("Called: lab.isCompatibleVersion");
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
