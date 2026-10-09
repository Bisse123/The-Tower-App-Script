const master = {

  /**
   * Reads IDS_Master data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: master.exportData");
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
        message: "IDS Master export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("master.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported IDS_Master data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: master.importData");

      var batchUpdate = [];
      var failedUpdates = [];

      var requiredRanges = ["IDS", "Presets Presets"];
      var idsData = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!idsData || !idsData[0] || !idsData[0].values) {
        console.log(`Could not read IDS data from new spreadsheet`);
        return {
          success: false,
          message: "Could not read IDS data from new spreadsheet",
        };
      }

      var idsValues = idsData[0].values;
      var presetsValues = idsData[1].values;

      if (data.hasOwnProperty("oldIdsData")) {
        try {
          var idsUpdateResult = masterWriter.updateIDSData(data.oldIdsData, idsValues);
          if (
            idsUpdateResult.success &&
            idsUpdateResult.batchUpdate.length > 0
          ) {
            batchUpdate = batchUpdate.concat(idsUpdateResult.batchUpdate);
          } else if (!idsUpdateResult.success) {
            failedUpdates.push({
              sheetType: "IDS",
              message: idsUpdateResult.message,
            });
          }
        } catch (error) {
          var errorReport = errors.report("master.importData", error, {
            note: `Error updating IDS data`,
            data: data,
            newSheetID: newSheetID,
          });
          failedUpdates.push({
            sheetType: "IDS",
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

      if (data.hasOwnProperty("oldPresetsData")) {
        try {
          var presetsUpdateResult = masterWriter.updatePresetsData(
            "Presets Presets",
            data.oldPresetsData,
            presetsValues,
          );
          if (
            presetsUpdateResult.success &&
            presetsUpdateResult.batchUpdate.length > 0
          ) {
            batchUpdate = batchUpdate.concat(presetsUpdateResult.batchUpdate);
          } else if (!presetsUpdateResult.success) {
            failedUpdates.push({
              sheetType: "Presets Presets",
              message: presetsUpdateResult.message,
            });
          }
        } catch (error) {
          var errorReport = errors.report("master.importData", error, {
            note: `Error updating Presets data`,
            data: data,
            newSheetID: newSheetID,
          });
          failedUpdates.push({
            sheetType: "Presets Presets",
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

      var thisSheetInfo = labelUtils.findSheetTypeID(
        newSheetID,
        "IDS",
        "This Sheet ID",
        idsValues,
      );
      if (thisSheetInfo && thisSheetInfo.cell && thisSheetInfo.cell.range) {
        batchUpdate.push({
          range: thisSheetInfo.cell.range,
          values: [[newSheetID]],
        });
      }

      if (batchUpdate.length > 0) {
        var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        if (!updateResult) {
          console.log(`Failed to update IDS Master data`);
          return {
            success: false,
            message: "Failed to update IDS Master data",
          };
        }
      }

      var successMessage = "IDS Master data imported successfully";
      if (failedUpdates.length > 0) {
        successMessage += ` (${failedUpdates.length} sections failed)`;
      }

      return {
        success: true,
        message: successMessage,
        failedUpdates: failedUpdates,
      };
    } catch (error) {
      var errorReport = errors.report("master.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v2.0": masterReader.version2_0.bind(masterReader),
      "v4.0": masterReader.version4_0.bind(masterReader),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
  isCompatibleVersion: function (oldVersion) {
    console.log("Called: master.isCompatibleVersion");
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
