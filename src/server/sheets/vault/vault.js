const vault = {

  /**
   * Reads Vault data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: vault.exportData");
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
        message: "Vault export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("vault.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Vault data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: vault.importData");

      var requiredRanges = ["IDS", "Master Sheet"];

      var batchResults = SheetsAPI.batchGetValues(
        newSheetID,
        requiredRanges
      );
      if (!batchResults || batchResults.length === 0) {
        console.log("Error getting vault sheet data");
        return {
          success: false,
          message: "Error getting vault sheet data",
        };
      }

      var idsData = batchResults[0].values;
      var masterSheetData = batchResults[1].values;

      var batchUpdate = [];
      console.log("data", JSON.stringify(data, null, 2));
      if (data.hasOwnProperty("oldVault")) {
        var updateResult = vaultWriter.updateVault(
          "Master Sheet",
          data.oldVault,
          masterSheetData
        );
        if (!updateResult || !updateResult.success) {
          console.log(`Error updating vault: ${updateResult.message}`);
          return updateResult;
        }
        batchUpdate = batchUpdate.concat(updateResult.batchUpdate || []);
      }

      labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Vault",
        newSheetID,
        idsData,
        data.idMasterID
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
        message: `Vault import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("vault.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": vaultReader.version1_0.bind(vaultReader),
      "v3.1": vaultReader.version3_1.bind(vaultReader),
      "v4.0": vaultReader.version4_0.bind(vaultReader),
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
