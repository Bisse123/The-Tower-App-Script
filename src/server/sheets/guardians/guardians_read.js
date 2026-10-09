const guardiansReader = {

  /**
   * Reads Guardians data from a v3.1 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_1: function (oldSheetID) {
    try {
      console.log("Called: guardiansReader.version3_1");

      var guardianLevelsRange = "EXPORT!B4:O";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion3_1Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version3_1", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Guardians data from a v2.2 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_2: function (oldSheetID) {
    try {
      console.log("Called: guardiansReader.version2_2");

      var guardianLevelsRange = "EXPORT!B5:F";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion2_2Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version2_2", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Guardians data from a v2.1 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_1: function (oldSheetID) {
    try {
      console.log("Called: guardiansReader.version2_1");

      var guardianLevelsRange = "EXPORT!B5:F";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion2_1Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version2_1", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Guardians data from a v1.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: guardiansReader.version1_0");

      var guardianLevelsRange = "EXPORT!B5:F";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion1_0Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.version1_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Guardians from a v3.1 sheet's values.
   * @param {Array<Array<*>>} oldGuardianLevelsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion3_1Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardiansReader.getVersion3_1Guardians");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch", "Summon", "Scout"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      if (!oldGuardianLevels || oldGuardianLevels.length === 0) {
        return {
          success: false,
          message: "No guardian levels data found",
        };
      }

      var oldGuardiansHeaderRow = oldGuardianLevels[0] || [];

      var firstPresetIndex = 4;

      if (oldGuardiansHeaderRow.length <= firstPresetIndex) {
        console.log(`Could not find the preset header row in guardian data`);
        return {
          success: false,
          message: "Could not find the preset header row in guardian data",
        };
      }

      var oldGuardiansPresetNames = [];
      var presetColumnMapping = [];

      for (
        var colIdx = firstPresetIndex;
        colIdx < oldGuardiansHeaderRow.length;
        colIdx++
      ) {
        var presetName = String(oldGuardiansHeaderRow[colIdx] || "").trim();
        if (!presetName) {
          continue;
        }

        oldGuardiansPresetNames.push(presetName);
        presetColumnMapping.push({
          presetName: presetName,
          equippedColIndex: colIdx,
          levelColIndex: colIdx + 1,
        });
      }

      var oldGuardians = {
        presetNames: presetUtils.resolvePresetOrder(
          oldGuardiansPresetNames,
          presetUtils.templatePresetNames,
        ).order,
        data: {},
      };

      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianRowData = oldGuardianLevels[row];
        var guardianName = String(guardianRowData[0] || "").trim();
        if (!guardianName || !targetGuardians.includes(guardianName)) {
          continue;
        }

        var unlocked;
        if (guardianName === "Attack" || guardianName === "Ally") {
          unlocked = null;
        } else {
          unlocked = oldGuardianLevels[row + 2]
            ? oldGuardianLevels[row + 2][0]
            : null;
        }

        var guardian = {
          unlocked: unlocked,
          presets: {},
        };

        presetColumnMapping.forEach(function (presetMap) {
          guardian.presets[presetMap.presetName] = {
            props: {},
            equipped: guardianRowData[presetMap.equippedColIndex],
          };
        });

        for (var nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
          var nextRowData = oldGuardianLevels[nextRow];
          if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
            row = nextRow - 1;
            break;
          }
          var key = String(nextRowData[2] || "").trim();
          if (!key) {
            continue;
          }

          presetColumnMapping.forEach(function (presetMap) {
            var levelValue = nextRowData[presetMap.levelColIndex];

            guardian.presets[presetMap.presetName].props[key] = levelValue;
          });
        }

        oldGuardians.data[guardianName] = guardian;
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion3_1Guardians", error, {
        oldGuardianLevelsData: oldGuardianLevelsData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Guardians from a v2.2 sheet's values.
   * @param {Array<Array<*>>} oldGuardianLevelsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_2Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardiansReader.getVersion2_2Guardians");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch", "Summon", "Scout"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {
        presetNames: ["Farming"],
        data: {},
      };
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            presets: {
              Farming: {
                props: {},
              },
            },
          };

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              guardian.presets.Farming.props[key] = value;
            }
          }
          oldGuardians.data[guardianName] = guardian;
        }
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      var errorReport = errors.report("guardian.getVersion2_2Guardians", error, {
        oldGuardianLevelsData: oldGuardianLevelsData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Guardians from a v2.1 sheet's values.
   * @param {Array<Array<*>>} oldGuardianLevelsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_1Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardiansReader.getVersion2_1Guardians");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {
        presetNames: ["Farming"],
        data: {},
      };
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            presets: {
              Farming: {
                props: {},
              },
            },
          };

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              guardian.presets.Farming.props[key] = value;
            }
          }
          oldGuardians.data[guardianName] = guardian;
        }
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      var errorReport = errors.report("guardian.getVersion2_1Guardians", error, {
        oldGuardianLevelsData: oldGuardianLevelsData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Guardians from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldGuardianLevelsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardiansReader.getVersion1_0Guardians");
      var targetGuardians = ["Attack", "Ally", "Steal", "Fetch"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {
        presetNames: ["Farming"],
        data: {},
      };
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            presets: {
              Farming: {
                props: {},
              },
            },
          };

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              value = (value - 1).toString().padStart(2, "0");
              guardian.presets.Farming.props[key] = value;
            }
          }
          guardianName = guardianName === "Steal" ? "Bounty" : guardianName;
          oldGuardians.data[guardianName] = guardian;
        }
      }
      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      var errorReport = errors.report("guardian.getVersion1_0Guardians", error, {
        oldGuardianLevelsData: oldGuardianLevelsData,
      });
      return errors.fail(errorReport);
    }
  },
};
