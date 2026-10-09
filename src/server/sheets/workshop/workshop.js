const workshop = {

  /**
   * Reads Workshop data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
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

      var oldDataResult = getVersionFunction(oldSheetID);
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
      var errorReport = errors.report("workshop.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Workshop data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: workshop.importData");

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

      if (
        data.hasOwnProperty("oldWorkshopLevels") &&
        data.hasOwnProperty("oldWorkshopPlusLevels")
      ) {
        var oldWorkshopLevels = data.oldWorkshopLevels;
        var oldWorkshopPlusLevels = data.oldWorkshopPlusLevels;
        var hasPresets = data.hasOwnProperty("hasPresets") ? data.hasPresets : true;
        var workshopResult = workshopWriter.updateWorkshopLevels(
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

      if (data.hasOwnProperty("oldWorkshopPlusRatios")) {
        var oldWorkshopPlusRatios = data.oldWorkshopPlusRatios;
        var ratioResult = workshopWriter.updateWorkshopPlusRatios(
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

      labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Workshop",
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
        message: `Workshop import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("workshop.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": workshopReader.version1_0.bind(workshopReader),
      "v2.0": workshopReader.version2_0.bind(workshopReader),
      "v2.1": workshopReader.version2_1.bind(workshopReader),
      "v2.2.8": workshopReader.version2_2_8.bind(workshopReader),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
  isCompatibleVersion: function (oldVersion) {
    console.log("Called: workshop.isCompatibleVersion");
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
