const guardians = {

  /**
   * Reads Guardians data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: guardians.exportData");
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
        message: "Guardians export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("guardians.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Guardians data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: guardians.importData");

      var requiredRanges = ["Master Sheet", "IDS"];
      var dvtIndex = requiredRanges.length;
      var dvtNamedRanges = {
        Attack: {
          Percentage: "DVT_GAR_UG_AT_PER",
          Cooldown: "DVT_GAR_UG_AT_COO",
          Targets: "DVT_GAR_UG_AT_TAR",
        },
        Ally: {
          "Recovery Amount": "DVT_GAR_UG_AL_REC",
          "Max Recovery": "DVT_GAR_UG_AL_MAX",
          Cooldown: "DVT_GAR_UG_AL_COO",
        },
        Bounty: {
          Multiplier: "DVT_GAR_UG_BO_MUL",
          Cooldown: "DVT_GAR_UG_BO_COO",
          Targets: "DVT_GAR_UG_BO_TAR",
        },
        Fetch: {
          Cooldown: "DVT_GAR_UG_FE_COO",
          "Find Chance": "DVT_GAR_UG_FE_FIN",
          "Double Find Chance": "DVT_GAR_UG_FE_DOU",
        },
        Summon: {
          Cooldown: "DVT_GAR_UG_SU_COO",
          Duration: "DVT_GAR_UG_SU_DUR",
          "Cash Bonus": "DVT_GAR_UG_SU_CAS",
        },
        Scout: {
          Cooldown: "DVT_GAR_UG_SC_COO",
          "Range Bonus": "DVT_GAR_UG_SC_RAN",
          Duration: "DVT_GAR_UG_SC_DUR",
        },
      };

      Object.keys(dvtNamedRanges).forEach(function (guardian) {
        Object.keys(dvtNamedRanges[guardian]).forEach(function (prop) {
          requiredRanges.push(dvtNamedRanges[guardian][prop]);
        });
      });

      var batchResult = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log("Error getting guardians sheet data");
        return {
          success: false,
          message: "Error getting guardians sheet data",
        };
      }

      var masterSheetData = batchResult[0].values;
      var idsData = batchResult[1].values;

      var dvtNamedRangesData = {};
      Object.keys(dvtNamedRanges).forEach(function (guardian) {
        dvtNamedRangesData[guardian] = {};
        Object.keys(dvtNamedRanges[guardian]).forEach(function (prop) {
          if (batchResult[dvtIndex]) {
            dvtNamedRangesData[guardian][prop] = batchResult[dvtIndex].values;
          } else {
            dvtNamedRangesData[guardian][prop] = [];
          }
          dvtIndex++;
        });
      });

      var batchUpdate = [];

      if (data.hasOwnProperty("oldGuardians")) {
        var oldGuardians = data.oldGuardians;
        var guardiansResult = guardiansWriter.updateGuardianLevels(
          "Master Sheet",
          oldGuardians,
          masterSheetData,
          dvtNamedRangesData
        );
        if (!guardiansResult || !guardiansResult.success) {
          console.log(`Error updating guardians: ${guardiansResult.message}`);
          return guardiansResult;
        }
        batchUpdate = batchUpdate.concat(guardiansResult.batchUpdate || []);
      }

      labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Guardians",
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
        message: `Guardians import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("guardians.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": guardiansReader.version1_0.bind(guardiansReader),
      "v2.1": guardiansReader.version2_1.bind(guardiansReader),
      "v2.2": guardiansReader.version2_2.bind(guardiansReader),
      "v3.1": guardiansReader.version3_1.bind(guardiansReader),
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
