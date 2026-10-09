const playerStuffReader = {

  /**
   * Reads Player_&_Stuff data from a v4.2 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_2: function (oldSheetID) {
    try {
      console.log("Called: playerStuffReader.version4_2");

      var tierRange = "EXPORT!B3:H";
      var statsRange = "EXPORT!J3:K";
      var ranges = [tierRange, statsRange, "Perk Preset"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);
      if (!batchResult || batchResult.length === 0) {
        console.log(`Could not read old player & stuff data`);
        return {
          success: false,
          message: `Could not read old player & stuff data`,
        };
      }
      var oldPlayerStuffTierValues = batchResult[0].values;
      var oldPlayerStuffStatsValues = batchResult[1].values;
      var oldPerksPresetValues = batchResult[2].values;
      var tierDataResult = this.getVersion4_0PlayerStuffTiers(
        oldPlayerStuffTierValues,
      );
      var statsDataResult = this.getVersion3_2PlayerStuffStats(
        oldPlayerStuffStatsValues,
      );
      var perksPresetResult =
        this.getVersion4_2PlayerStuffPerks(oldPerksPresetValues);
      success =
        tierDataResult.success &&
        statsDataResult.success &&
        perksPresetResult.success;
      return {
        success: success,
        message: success
          ? "Player & Stuff processed successfully"
          : "Error processing Player & Stuff data",
        oldPlayerStuffTierData: tierDataResult.oldPlayerStuffTierData,
        oldPlayerStuffStatsData: statsDataResult.oldPlayerStuffStatsData,
        oldPerksPreset: perksPresetResult.oldPerksPreset,
        shouldRemoveUsedPerks: perksPresetResult.shouldRemoveUsedPerks,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.version4_2", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Player_&_Stuff data from a v4.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_0: function (oldSheetID) {
    try {
      console.log("Called: playerStuffReader.version4_0");

      var tierRange = "EXPORT!B3:H";
      var statsRange = "EXPORT!J3:K";
      var ranges = [tierRange, statsRange];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);
      if (!batchResult || batchResult.length === 0) {
        console.log(`Could not read old player & stuff data`);
        return {
          success: false,
          message: `Could not read old player & stuff data`,
        };
      }
      var oldPlayerStuffTierValues = batchResult[0].values;
      var oldPlayerStuffStatsValues = batchResult[1].values;
      var tierDataResult = this.getVersion4_0PlayerStuffTiers(
        oldPlayerStuffTierValues,
      );
      var statsDataResult = this.getVersion3_2PlayerStuffStats(
        oldPlayerStuffStatsValues,
      );
      success = tierDataResult.success && statsDataResult.success;
      return {
        success: success,
        message: success
          ? "Player & Stuff processed successfully"
          : "Error processing Player & Stuff data",
        oldPlayerStuffTierData: tierDataResult.oldPlayerStuffTierData,
        oldPlayerStuffStatsData: statsDataResult.oldPlayerStuffStatsData,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.version4_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Player_&_Stuff data from a v3.2 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_2: function (oldSheetID) {
    try {
      console.log("Called: playerStuffReader.version3_2");

      var tierRange = "EXPORT!B3:D";
      var statsRange = "EXPORT!F3:G";
      var ranges = [tierRange, statsRange];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);
      if (!batchResult || batchResult.length === 0) {
        console.log(`Could not read old player & stuff data`);
        return {
          success: false,
          message: `Could not read old player & stuff data`,
        };
      }
      var oldPlayerStuffTierValues = batchResult[0].values;
      var oldPlayerStuffStatsValues = batchResult[1].values;
      var tierDataResult = this.getVersion2_0PlayerStuffTiers(
        oldPlayerStuffTierValues,
      );
      var statsDataResult = this.getVersion3_2PlayerStuffStats(
        oldPlayerStuffStatsValues,
      );
      success = tierDataResult.success && statsDataResult.success;
      return {
        success: success,
        message: success
          ? "Player & Stuff processed successfully"
          : "Error processing Player & Stuff data",
        oldPlayerStuffTierData: tierDataResult.oldPlayerStuffTierData,
        oldPlayerStuffStatsData: statsDataResult.oldPlayerStuffStatsData,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.version3_2", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Player_&_Stuff data from a v2.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_0: function (oldSheetID) {
    try {
      console.log("Called: playerStuffReader.version2_0");

      var tierRange = "EXPORT!B16:D";
      var statsRange = "EXPORT!B2:C12";
      var ranges = [tierRange, statsRange];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);
      if (!batchResult || batchResult.length === 0) {
        console.log(`Could not read old player & stuff data`);
        return {
          success: false,
          message: `Could not read old player & stuff data`,
        };
      }
      var oldPlayerStuffTierValues = batchResult[0].values;
      var oldPlayerStuffStatsValues = batchResult[1].values;
      var tierDataResult = this.getVersion2_0PlayerStuffTiers(
        oldPlayerStuffTierValues,
      );
      var statsDataResult = this.getVersion2_0PlayerStuffStats(
        oldPlayerStuffStatsValues,
      );
      success = tierDataResult.success && statsDataResult.success;
      return {
        success: success,
        message: success
          ? "Player & Stuff processed successfully"
          : "Error processing Player & Stuff data",
        oldPlayerStuffTierData: tierDataResult.oldPlayerStuffTierData,
        oldPlayerStuffStatsData: statsDataResult.oldPlayerStuffStatsData,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.version2_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts PlayerStuffTiers from a v4.0 sheet's values.
   * @param {Array<Array<*>>} oldPlayerStuffTierValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0PlayerStuffTiers: function (oldPlayerStuffTierValues) {
    try {
      console.log("Called: playerStuffReader.getVersion4_0PlayerStuffTiers");

      if (!oldPlayerStuffTierValues || oldPlayerStuffTierValues.length === 0) {
        console.log(`No data found in old player & stuff tier data`);
        return {
          success: false,
          message: "No data found in old player & stuff tier data",
        };
      }

      var stripCommas = function (v) {
        if (v === null || v === undefined || v === "") return v;
        var s = String(v).replace(/,/g, "");
        return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : s;
      };
      var oldPlayerStuffTierData = {};
      for (var row = 0; row < oldPlayerStuffTierValues.length; row++) {
        var rowData = oldPlayerStuffTierValues[row];
        var tier = rowData[0] || "";
        var wave = stripCommas(rowData[1] || "");
        var dissAttack = stripCommas(rowData[2] || "");
        var dissDefense = stripCommas(rowData[3] || "");
        var dissUtility = stripCommas(rowData[4] || "");
        var dissUltimate = stripCommas(rowData[5] || "");
        var premium = rowData[6] || "";
        if (tier) {
          oldPlayerStuffTierData[tier] = {
            wave: wave,
            diss: {
              attack: dissAttack,
              defense: dissDefense,
              utility: dissUtility,
              ultimate: dissUltimate,
            },
            premium: premium,
          };
        }
      }

      return {
        success: true,
        message: "Player & Stuff tier processed successfully",
        oldPlayerStuffTierData: oldPlayerStuffTierData,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.getVersion4_0PlayerStuffTiers", error, {
        oldPlayerStuffTierValues: oldPlayerStuffTierValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts PlayerStuffTiers from a v2.0 sheet's values.
   * @param {Array<Array<*>>} oldPlayerStuffTierValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_0PlayerStuffTiers: function (oldPlayerStuffTierValues) {
    try {
      console.log("Called: playerStuffReader.getVersion2_0PlayerStuffTiers");

      if (!oldPlayerStuffTierValues || oldPlayerStuffTierValues.length === 0) {
        console.log(`No data found in old player & stuff tier data`);
        return {
          success: false,
          message: "No data found in old player & stuff tier data",
        };
      }

      var stripCommas = function (v) {
        if (v === null || v === undefined || v === "") return v;
        var s = String(v).replace(/,/g, "");
        return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : s;
      };
      var oldPlayerStuffTierData = {};
      for (var row = 0; row < oldPlayerStuffTierValues.length; row++) {
        var rowData = oldPlayerStuffTierValues[row];
        var tier = rowData[0] || "";
        var wave = stripCommas(rowData[1] || "");
        var premium = rowData[2] || "";
        if (tier) {
          oldPlayerStuffTierData[tier] = {
            wave: wave,
            premium: premium,
          };
        }
      }

      return {
        success: true,
        message: "Player & Stuff tier processed successfully",
        oldPlayerStuffTierData: oldPlayerStuffTierData,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.getVersion2_0PlayerStuffTiers", error, {
        oldPlayerStuffTierValues: oldPlayerStuffTierValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts PlayerStuffStats from a v3.2 sheet's values.
   * @param {Array<Array<*>>} oldPlayerStuffStatsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion3_2PlayerStuffStats: function (oldPlayerStuffStatsValues) {
    try {
      console.log("Called: playerStuffReader.getVersion3_2PlayerStuffStats");

      if (
        !oldPlayerStuffStatsValues ||
        oldPlayerStuffStatsValues.length === 0
      ) {
        console.log(`No data found in old player & stuff stat data`);
        return {
          success: false,
          message: "No data found in old player & stuff stat data",
        };
      }
      var oldPlayerStuffStatsData = {};
      var header = "Stat";
      oldPlayerStuffStatsData[header] = {};
      for (var row = 0; row < oldPlayerStuffStatsValues.length; row++) {
        var rowData = oldPlayerStuffStatsValues[row];
        var name = rowData[0] || "";
        var value = rowData[1] || "";
        if (name === "Premium Packs" || name === "Premium Perk") {
          header = "Premium Packs";
          oldPlayerStuffStatsData[header] = {};
          continue;
        }
        if (name === "Coin Multiplier") {
          break;
        }
        if (name) {
          oldPlayerStuffStatsData[header][name] = value;
        }
      }
      return {
        success: true,
        message: "Player & Stuff stats processed successfully",
        oldPlayerStuffStatsData: oldPlayerStuffStatsData,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.getVersion3_2PlayerStuffStats", error, {
        oldPlayerStuffStatsValues: oldPlayerStuffStatsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts PlayerStuffStats from a v2.0 sheet's values.
   * @param {Array<Array<*>>} oldPlayerStuffStatsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_0PlayerStuffStats: function (oldPlayerStuffStatsValues) {
    try {
      console.log("Called: playerStuffReader.getVersion2_0PlayerStuffStats");

      if (
        !oldPlayerStuffStatsValues ||
        oldPlayerStuffStatsValues.length === 0
      ) {
        console.log(`No data found in old player & stuff stat data`);
        return {
          success: false,
          message: "No data found in old player & stuff stat data",
        };
      }
      var oldPlayerStuffStatsData = {};
      var header = "Stat";
      oldPlayerStuffStatsData[header] = {};
      for (var row = 0; row < oldPlayerStuffStatsValues.length; row++) {
        var rowData = oldPlayerStuffStatsValues[row];
        var name = rowData[0] || "";
        var value = rowData[1] || "";
        if (name === "Premium Perk") {
          header = "Premium Packs";
          oldPlayerStuffStatsData[header] = {};
          continue;
        }
        if (name === "Coin Multiplier") {
          break;
        }
        if (name) {
          oldPlayerStuffStatsData[header][name] = value;
        }
      }
      return {
        success: true,
        message: "Player & Stuff stats processed successfully",
        oldPlayerStuffStatsData: oldPlayerStuffStatsData,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.getVersion2_0PlayerStuffStats", error, {
        oldPlayerStuffStatsValues: oldPlayerStuffStatsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts PlayerStuffPerks from a v4.2 sheet's values.
   * @param {Array<Array<*>>} oldPlayerStuffPerksValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_2PlayerStuffPerks: function (oldPlayerStuffPerksValues) {
    try {
      console.log("Called: playerStuffReader.getVersion4_2PlayerStuffPerks");
      if (
        !oldPlayerStuffPerksValues ||
        oldPlayerStuffPerksValues.length === 0
      ) {
        console.log(`No data found in old player & stuff perks data`);
        return {
          success: false,
          message: "No data found in old player & stuff perks data",
        };
      }
      var shouldRemoveUsedPerks;
      var oldPerksPreset = {};

      var headerRowIndex = -1;
      for (
        var rowIndex = 0;
        rowIndex < oldPlayerStuffPerksValues.length;
        rowIndex++
      ) {
        var colIndex = oldPlayerStuffPerksValues[rowIndex].indexOf(
          "Remove used perks from the pool",
        );
        if (colIndex !== -1) {
          shouldRemoveUsedPerks =
            oldPlayerStuffPerksValues[rowIndex][colIndex - 1] === "TRUE" ||
            oldPlayerStuffPerksValues[rowIndex][colIndex - 1] === "true" ||
            oldPlayerStuffPerksValues[rowIndex][colIndex - 1] === true;
          headerRowIndex = rowIndex + 2;
          break;
        }
      }

      if (headerRowIndex === -1 || !oldPlayerStuffPerksValues[headerRowIndex]) {
        console.log(`Could not find the preset header row in the perks data`);
        return {
          success: false,
          message: "Could not find the preset header row in the perks data",
        };
      }

      var row = oldPlayerStuffPerksValues[headerRowIndex];
      var oldPerkPresetNameIdxs = row
        .map(function (cell, idx) {
          return String(cell || "").trim() !== "" ? idx : -1;
        })
        .filter(function (idx) {
          return idx !== -1;
        });

      var presetOrder = presetUtils.resolvePresetOrder(
        oldPerkPresetNameIdxs.map(function (colIdx) {
          return row[colIdx];
        }),
        presetUtils.templatePresetNames,
      );
      var orderBySourceIndex = {};
      presetOrder.indices.forEach(function (sourceIndex, slot) {
        orderBySourceIndex[sourceIndex] = slot + 1;
      });

      oldPerkPresetNameIdxs.forEach(function (colIdx, sourceIndex) {
        var presetName = row[colIdx];
        if (!oldPerksPreset.hasOwnProperty(presetName)) {
          oldPerksPreset[presetName] = {
            perks: [],
            order: orderBySourceIndex[sourceIndex],
          };
        }
        var bannedAmount =
          oldPlayerStuffPerksValues[headerRowIndex + 1][colIdx + 2] || 0;
        oldPerksPreset[presetName].bannedAmount = bannedAmount;
        for (
          var rowIdx = headerRowIndex + 3;
          rowIdx < oldPlayerStuffPerksValues.length;
          rowIdx++
        ) {
          var perkValue = oldPlayerStuffPerksValues[rowIdx][colIdx + 1];
          var perkNumber = oldPlayerStuffPerksValues[rowIdx][colIdx];
          if (!perkNumber || String(perkNumber).trim() === "") {
            break;
          }
          oldPerksPreset[presetName].perks.push(perkValue || null);
        }
        var lastPerk =
          oldPerksPreset[presetName].perks[
            oldPerksPreset[presetName].perks.length - 1
          ];
        if (
          lastPerk &&
          String(lastPerk).trim() === "Unlock a random Ultimate Weapon"
        ) {
          oldPerksPreset[presetName].perks.pop();
        }
      });

      return {
        success: true,
        message: "Player & Stuff perks processed successfully",
        oldPerksPreset: oldPerksPreset,
        shouldRemoveUsedPerks: shouldRemoveUsedPerks,
      };
    } catch (error) {
      var errorReport = errors.report("ranges.getVersion4_2PlayerStuffPerks", error, {
        oldPlayerStuffPerksValues: oldPlayerStuffPerksValues,
      });
      return errors.fail(errorReport);
    }
  },
};
