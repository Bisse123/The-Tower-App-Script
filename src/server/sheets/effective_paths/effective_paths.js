const ePaths = {

  /**
   * Reads EPaths data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: ePaths.exportData");
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
        message: "Effective Paths export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported EPaths data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: ePaths.importData");

      var eHPRange = "eHP!AJ1:AY50";
      var eDamageRange = "eDamage!AI1:AY100";
      var eEconRange = "eEcon!AK1:AZ65";

      var eHPLabRange = "eHP!L3:L5";
      var eRegenLabRange = "eHP!AH3:AH5";
      var eDamageLabRange = "eDamage!L3:L5";
      var eEconLabRange = "eEcon!O3:O5";
      var eEconStoneMult = "eEcon!X3:X5";
      var eDiscountLabRange = "eEcon!AH3:AH5";

      var eHPLabColumn = "L";
      var eRegenLabColumn = "AH";
      var eDamageLabColumn = "L";
      var eEconLabColumn = "O";
      var eEconStoneMultColumn = "X";
      var eDiscountLabColumn = "AH";

      var ranges = [
        "IDS",
        eHPRange,
        eDamageRange,
        eEconRange,
        eHPLabRange,
        eRegenLabRange,
        eDamageLabRange,
        eEconLabRange,
        eEconStoneMult,
        eDiscountLabRange,
      ];

      var eHPColumnOffset = sheetRefs.getColumnOffsetFromRange(eHPRange);
      var eDamageColumnOffset = sheetRefs.getColumnOffsetFromRange(eDamageRange);
      var eEconColumnOffset = sheetRefs.getColumnOffsetFromRange(eEconRange);

      var batchGetResult = SheetsAPI.batchGetFormulas(newSheetID, ranges);
      if (!batchGetResult || batchGetResult.length === 0) {
        return {
          success: false,
          message: "Failed to fetch data from new spreadsheet™.",
        };
      }

      var idsData = batchGetResult[0].values;
      var eHPValues = batchGetResult[1].values;
      var eDamageValues = batchGetResult[2].values;
      var eEconValues = batchGetResult[3].values;
      var eHPLabValues = batchGetResult[4].values;
      var eRegenLabValues = batchGetResult[5].values;
      var eDamageLabValues = batchGetResult[6].values;
      var eEconLabValues = batchGetResult[7].values;
      var eEconStoneMultValues = batchGetResult[8].values;
      var eDiscountLabValues = batchGetResult[9].values;

      var batchUpdate = [];

      if (data.hasOwnProperty("eHP")) {
        var eHPResult = ePathsEHP.updateEHP(
          "eHP",
          data.eHP.oldData,
          eHPValues,
          eHPColumnOffset,
          eHPLabValues,
          eRegenLabValues,
          eHPLabColumn,
          eRegenLabColumn,
        );
        if (!eHPResult || !eHPResult.success) {
          console.log(`eHP update failed: ${eHPResult.message}`);
          return {
            success: false,
            message: `eHP update failed: ${eHPResult.message}`,
          };
        }
        batchUpdate = batchUpdate.concat(eHPResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("eDamage")) {
        var eDamageResult = ePathsEDamage.updateEDamage(
          "eDamage",
          data.eDamage.oldData,
          eDamageValues,
          eDamageColumnOffset,
          eDamageLabValues,
          eDamageLabColumn,
        );
        if (!eDamageResult || !eDamageResult.success) {
          console.log(`eDamage update failed: ${eDamageResult.message}`);
          return {
            success: false,
            message: `eDamage update failed: ${eDamageResult.message}`,
          };
        }
        batchUpdate = batchUpdate.concat(eDamageResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("eEcon")) {
        var eEconResult = ePathsEEcon.updateEEcon(
          "eEcon",
          data.eEcon.oldData,
          eEconValues,
          eEconColumnOffset,
          eEconLabValues,
          eEconStoneMultValues,
          eDiscountLabValues,
          eEconLabColumn,
          eEconStoneMultColumn,
          eDiscountLabColumn,
        );
        if (!eEconResult || !eEconResult.success) {
          console.log(`eEcon update failed: ${eEconResult.message}`);
          return {
            success: false,
            message: `eEcon update failed: ${eEconResult.message}`,
          };
        }
        batchUpdate = batchUpdate.concat(eEconResult.batchUpdate || []);
      }

      labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Effective Paths",
        newSheetID,
        idsData,
        data.idMasterID,
      );

      var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
      if (!updateResult) {
        return {
          success: false,
          message:
            "Failed to update new spreadsheet™ with Effective Paths data.",
        };
      }
      return {
        success: true,
        message: `Effective Paths import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths.importData", error, {
        note: `Error importing ePaths data`,
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v4.11.02.00": ePathsReader.version4_11_02_00.bind(ePathsReader),
      "v4.11.03.21": ePathsReader.version4_11_03_21.bind(ePathsReader),
      "v5.00.01.04": ePathsReader.version5_00_01_04.bind(ePathsReader),
      "v5.03.00.00": ePathsReader.version5_03_00_00.bind(ePathsReader),
      "v5.05.00.00": ePathsReader.version5_05_00_00.bind(ePathsReader),
      "v5.05.01.00": ePathsReader.version5_05_01_00.bind(ePathsReader),
      "v5.06.02.00": ePathsReader.version5_06_02_00.bind(ePathsReader),
      "v5.08.00.00": ePathsReader.version5_08_00_00.bind(ePathsReader),
      "v5.08.04.00": ePathsReader.version5_08_04_00.bind(ePathsReader),
      "v5.09.00.00": ePathsReader.version5_09_00_00.bind(ePathsReader),
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
    for (var key = 0; key < sortedThresholds.length; key++) {
      var threshold = sortedThresholds[key];
      var compareResult = versionUtils.compareVersions(oldVersion, threshold);
      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }
    return null;
  },
};
