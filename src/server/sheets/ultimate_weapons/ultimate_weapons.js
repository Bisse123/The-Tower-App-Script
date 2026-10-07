const ultimate = {

  /**
   * Reads Ultimate_Weapons data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: ultimate.exportData");
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
        message: "Ultimate weapons export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("ultimate.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Ultimate_Weapons data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: ultimate.importData");

      var requiredRanges = ["Master Sheet", "UW Cost Calculator v3", "IDS"];
      var dvtIndex = requiredRanges.length;
      var dvtNamedRanges = {
        "Chain Lightning": {
          Damage: "DVT_UW_UG_CL_DMG",
          Quantity: "DVT_UW_UG_CL_QNT",
          Chance: "DVT_UW_UG_CL_CH",
          Smite: "DVT_UW_UG_CL_SM",
        },
        "Smart Missiles": {
          Damage: "DVT_UW_UG_SM_DMG",
          Quantity: "DVT_UW_UG_SM_QNT",
          Cooldown: "DVT_UW_UG_SM_CD",
          "Cover Fire": "DVT_UW_UG_SM_CF",
        },
        "Death Wave": {
          Damage: "DVT_UW_UG_DW_DMG",
          Quantity: "DVT_UW_UG_DW_QNT",
          Cooldown: "DVT_UW_UG_DW_CD",
          "Kill Wall": "DVT_UW_UG_DW_KW",
        },
        "Chrono Field": {
          Duration: "DVT_UW_UG_CF_DU",
          "Speed Reduction": "DVT_UW_UG_CF_SP",
          Cooldown: "DVT_UW_UG_CF_CD",
          "Chrono Loop": "DVT_UW_UG_CF_CL",
        },
        "Inner Land Mines": {
          Damage: "DVT_UW_UG_ILM_DMG",
          Quantity: "DVT_UW_UG_ILM_QNT",
          Cooldown: "DVT_UW_UG_ILM_CD",
          "Charged Mines": "DVT_UW_UG_ILM_CM",
        },
        "Golden Tower": {
          Multiplier: "DVT_UW_UG_GT_M",
          Duration: "DVT_UW_UG_GT_DU",
          Cooldown: "DVT_UW_UG_GT_CD",
          "Golden Combo": "DVT_UW_UG_GT_GC",
        },
        "Poison Swamp": {
          Damage: "DVT_UW_UG_PS_DMG",
          Duration: "DVT_UW_UG_PS_DU",
          Cooldown: "DVT_UW_UG_PS_CH",
          "Death Creep": "DVT_UW_UG_PS_DC",
        },
        "Black Hole": {
          Size: "DVT_UW_UG_BH_SZ",
          Duration: "DVT_UW_UG_BH_DU",
          Cooldown: "DVT_UW_UG_BH_CD",
          Consume: "DVT_UW_UG_BH_C",
        },
        Spotlight: {
          Multiplier: "DVT_UW_UG_SL_MU",
          Angle: "DVT_UW_UG_SL_AN",
          Quantity: "DVT_UW_UG_SL_QNT",
          "Light Range": "DVT_UW_UG_SL_LR",
        },
      };
      Object.keys(dvtNamedRanges).forEach(function (weapon) {
        Object.keys(dvtNamedRanges[weapon]).forEach(function (level) {
          requiredRanges.push(dvtNamedRanges[weapon][level]);
        });
      });

      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var ultimateCostCalculatorData = batchResults[1].values;
      var idsData = batchResults[2].values;

      var dvtNamedRangesData = {};
      Object.keys(dvtNamedRanges).forEach(function (weapon) {
        dvtNamedRangesData[weapon] = {};
        Object.keys(dvtNamedRanges[weapon]).forEach(function (level) {
          if (batchResults[dvtIndex]) {
            dvtNamedRangesData[weapon][level] = batchResults[dvtIndex].values;
          } else {
            dvtNamedRangesData[weapon][level] = [];
          }
          dvtIndex++;
        });
      });

      var batchUpdate = [];

      if (data.hasOwnProperty("oldUltimate")) {
        var oldUltimate = data.oldUltimate;
        var ultimateResult = ultimateWriter.updateUltimateLevels(
          "Master Sheet",
          oldUltimate,
          masterSheetData,
          dvtNamedRangesData,
        );
        if (!ultimateResult || !ultimateResult.success) {
          console.log(
            `Error updating ultimate weapon levels: ${ultimateResult.message}`,
          );
          return ultimateResult;
        }
        batchUpdate = batchUpdate.concat(ultimateResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldUltimateCostCalculator")) {
        var oldUltimateCostCalculator = data.oldUltimateCostCalculator;
        var ultimateCostCalculatorResult = ultimateWriter.updateUltimateCostCalculator(
          "UW Cost Calculator v3",
          oldUltimateCostCalculator,
          ultimateCostCalculatorData,
        );
        if (
          !ultimateCostCalculatorResult ||
          !ultimateCostCalculatorResult.success
        ) {
          console.log(
            `Error updating ultimate cost calculator: ${ultimateCostCalculatorResult.message}`,
          );
          return ultimateCostCalculatorResult;
        }
        batchUpdate = batchUpdate.concat(
          ultimateCostCalculatorResult.batchUpdate || [],
        );
      }

      labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Ultimate Weapon",
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
        message: `Ultimate Weapons import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("ultimate.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": ultimateReader.version1_0.bind(ultimateReader),
      "v2.0": ultimateReader.version2_0.bind(ultimateReader),
      "v3.1.1": ultimateReader.version3_1_1.bind(ultimateReader),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
  isCompatibleVersion: function (oldVersion) {
    console.log("Called: ultimate.isCompatibleVersion");
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
