const playerStuff = {

  /**
   * Reads Player_&_Stuff data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: playerStuff.exportData");
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
        message: "Player & Stuff export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("playerStuff.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Player_&_Stuff data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: playerStuff.importData");

      var requiredRanges = ["Master Sheet", "Perk Preset", "IDS"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var perksSheetData = batchResults[1].values;
      var idsData = batchResults[2].values;

      var batchUpdate = [];

      if (
        data.hasOwnProperty("oldPlayerStuffTierData") &&
        data.hasOwnProperty("oldPlayerStuffStatsData")
      ) {
        var oldPlayerStuffTierData = data.oldPlayerStuffTierData;
        var oldPlayerStuffStatsData = data.oldPlayerStuffStatsData;
        var playerStuffResult = this.updatePlayerStuffData(
          "Master Sheet",
          oldPlayerStuffTierData,
          oldPlayerStuffStatsData,
          masterSheetData,
        );
        if (!playerStuffResult || !playerStuffResult.success) {
          console.log(
            `Error updating player data: ${playerStuffResult.message}`,
          );
          return playerStuffResult;
        }
        batchUpdate = batchUpdate.concat(playerStuffResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldPerksPreset")) {
        var oldPerksPreset = data.oldPerksPreset;
        var shouldRemoveUsedPerks = data.hasOwnProperty("shouldRemoveUsedPerks")
          ? data.shouldRemoveUsedPerks
          : true;

        var playerPerksResult = this.updatePlayerPerksPreset(
          "Perk Preset",
          oldPerksPreset,
          shouldRemoveUsedPerks,
          perksSheetData,
        );
        if (!playerPerksResult || !playerPerksResult.success) {
          console.log(
            `Error updating player perks data: ${playerPerksResult.message}`,
          );
          return playerPerksResult;
        }
        batchUpdate = batchUpdate.concat(playerPerksResult.batchUpdate || []);
      }

      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Player & Stuff",
        newSheetID,
        idsData,
        data.idMasterID,
      );

      var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
      if (!updateResult) {
        console.log(`Error applying batch updates to new spreadsheet`);
        return {
          success: false,
          message: "Error applying batch updates to new spreadsheet",
        };
      }

      return {
        success: true,
        message: `Player & Stuff import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("playerStuff.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes PlayerStuffData into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldPlayerTierData
   * @param {Object} oldPlayerStatsData
   * @param {Object} masterSheetData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updatePlayerStuffData: function (
    sheetName,
    oldPlayerTierData,
    oldPlayerStatsData,
    masterSheetData,
  ) {
    try {
      console.log("Called: playerStuff.updatePlayerStuffData");
      if (!masterSheetData || masterSheetData.length === 0) {
        console.log(`Master sheet data is empty or not found`);
        return {
          success: false,
          message: "Master sheet data is empty or not found",
        };
      }
      var headerRow = masterSheetData[0] || [];
      var statCol = headerRow.indexOf("Stat");
      var tierCol = headerRow.indexOf("Tier");
      var dissCol = headerRow.indexOf("Dissonant Runs");
      var passCol = headerRow.indexOf("Pass");

      if (
        statCol === -1 ||
        tierCol === -1 ||
        passCol === -1 ||
        dissCol === -1
      ) {
        console.log(
          `Stat, Tier, Pass, or Dissonant Runs column not found in master sheet`,
        );
        return {
          success: false,
          message:
            "Stat, Tier, Pass, or Dissonant Runs column not found in master sheet",
        };
      }
      var header = headerRow[statCol] || "";
      var perkRow = -1;
      var values = {
        Stat: [],
        Tier: [],
        Pass: [],
        "Premium Packs": [],
      };
      var dissHeaders = ["attack", "defense", "utility", "ultimate"];
      var dissValuesByName = {};
      var dissColsByName = {};

      var firstRow = -1;
      for (var row = 0; row < masterSheetData.length; row++) {
        var dissHeaderRow = masterSheetData[row] || [];
        var dissIndexesByName = {};
        var allFound = true;
        for (var i = 0; i < dissHeaders.length; i++) {
          var dissHeaderName = dissHeaders[i];
          var dissIndex = dissHeaderRow.findIndex(function (cell) {
            return String(cell || "")
              .toLowerCase()
              .includes(dissHeaderName);
          });
          if (dissIndex === -1) {
            allFound = false;
            break;
          }
          dissIndexesByName[dissHeaderName] = dissIndex;
        }
        if (allFound) {
          firstRow = row + 2;
          for (var j = 0; j < dissHeaders.length; j++) {
            var dissName = dissHeaders[j];
            dissColsByName[dissName] = dissIndexesByName[dissName] + 1;
          }
          break;
        }
      }
      if (firstRow === -1) {
        console.log(
          `Could not find Dissonant Runs subheader row in master sheet`,
        );
        return {
          success: false,
          message:
            "Could not find Dissonant Runs subheader row in master sheet",
        };
      }

      for (var row = firstRow - 1; row < masterSheetData.length; row++) {
        var rowData = masterSheetData[row];
        var statName = rowData[statCol] || "";
        var tierValue = rowData[tierCol] || "";
        if (!tierValue) {
          break;
        }

        if (oldPlayerTierData && oldPlayerTierData[tierValue]) {
          var wave = oldPlayerTierData[tierValue].wave || null;
          var premium = oldPlayerTierData[tierValue].premium || null;
          values.Tier.push([wave]);
          values.Pass.push([premium]);
          if (oldPlayerTierData[tierValue].diss) {
            for (var i = 0; i < dissHeaders.length; i++) {
              var dissHeaderName = dissHeaders[i];
              var dissValue =
                oldPlayerTierData[tierValue].diss[dissHeaderName] || null;
              if (!dissValuesByName[dissHeaderName]) {
                dissValuesByName[dissHeaderName] = [];
              }
              dissValuesByName[dissHeaderName].push([dissValue]);
            }
          }
        }

        if (!statName) {
          continue;
        }

        if (statName === "Premium Packs") {
          header = "Premium Packs";
          perkRow = row + 2;
          continue;
        }

        if (
          oldPlayerStatsData[header] &&
          oldPlayerStatsData[header][statName]
        ) {
          var value = oldPlayerStatsData[header][statName] || null;
          values[header].push([value]);
        } else {
          values[header].push([null]);
        }
      }

      var statColLetter = shared.columnToLetter(statCol + 2);
      var tierColLetter = shared.columnToLetter(tierCol + 2);
      var passColLetter = shared.columnToLetter(passCol + 1);
      var batchUpdate = [];
      var ranges = {
        Stat: `${sheetName}!${statColLetter}${firstRow}:${statColLetter}${
          firstRow + values.Stat.length - 1
        }`,
        Tier: `${sheetName}!${tierColLetter}${firstRow}:${tierColLetter}${
          firstRow + values.Tier.length - 1
        }`,
        Pass: `${sheetName}!${passColLetter}${firstRow}:${passColLetter}${
          firstRow + values.Pass.length - 1
        }`,
        "Premium Packs": `${sheetName}!${statColLetter}${perkRow}:${statColLetter}${
          perkRow + values["Premium Packs"].length - 1
        }`,
      };
      for (var key in values) {
        if (values[key].length > 0) {
          batchUpdate.push({
            range: ranges[key],
            values: values[key],
          });
        }
      }
      for (var i = 0; i < dissHeaders.length; i++) {
        var dissHeaderName = dissHeaders[i];
        var dissValues = dissValuesByName[dissHeaderName];
        if (dissValues && dissValues.length > 0) {
          var dissColLetter = shared.columnToLetter(
            dissColsByName[dissHeaderName],
          );
          batchUpdate.push({
            range: `${sheetName}!${dissColLetter}${firstRow}:${dissColLetter}${
              firstRow + dissValues.length - 1
            }`,
            values: dissValues,
          });
        }
      }
      if (batchUpdate.length === 0) {
        console.log(`No data to update in player & stuff data`);
        return {
          success: false,
          message: "No data to update in player & stuff data",
        };
      }
      return {
        success: true,
        message: "Player & Stuff data updated successfully",
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      var errorReport = errors.report("playerStuff.updatePlayerStuffData", error, {
        sheetName: sheetName,
        oldPlayerTierData: oldPlayerTierData,
        oldPlayerStatsData: oldPlayerStatsData,
        masterSheetData: masterSheetData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes PlayerPerksPreset into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldPlayerPerksData
   * @param {*} shouldRemoveUsedPerks
   * @param {Object} perksSheetData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updatePlayerPerksPreset: function (
    sheetName,
    oldPlayerPerksData,
    shouldRemoveUsedPerks,
    perksSheetData,
  ) {
    try {
      console.log("Called: playerStuff.updatePlayerPerksPreset");
      if (!perksSheetData) {
        console.log(`Error getting perks sheet data`);
        return {
          success: false,
          message: "Error getting perks sheet data",
        };
      }
      if (perksSheetData.length < 3) {
        console.log(`Perks sheet has no data or only header row`);
        return {
          success: false,
          message: "Perks sheet has no data or only header row",
        };
      }

      var headerRowIndex = -1;
      var headerColIndices = [];
      var batchUpdate = [];

      for (var row = 0; row < perksSheetData.length; row++) {
        var removeUsedPerksIndex = perksSheetData[row].indexOf(
          "Remove used perks from the pool",
        );
        if (removeUsedPerksIndex === -1) {
          continue;
        }
        batchUpdate.push({
          range:
            sheetName +
            "!" +
            shared.columnToLetter(removeUsedPerksIndex) +
            (row + 1),
          values: [[shouldRemoveUsedPerks]],
        });
        headerRowIndex = row + 2;
        break;
      }

      if (headerRowIndex === -1 || headerRowIndex >= perksSheetData.length) {
        console.log(`Could not find "Remove used perks from the pool"`);
        return {
          success: false,
          message: "Could not find Remove used perks from the pool",
        };
      }

      var presetHeaderRow = perksSheetData[headerRowIndex];
      for (
        var col = 0;
        col < presetHeaderRow.length && headerColIndices.length < 5;
        col++
      ) {
        if (String(presetHeaderRow[col] || "").trim() !== "") {
          headerColIndices.push(col);
        }
      }

      if (headerColIndices.length < 5) {
        console.log(
          `Expected 5 preset columns but found ${headerColIndices.length}`,
        );
        return {
          success: false,
          message: `Expected 5 preset columns but found ${headerColIndices.length}`,
        };
      }

      Object.keys(oldPlayerPerksData).forEach(function (presetName) {
        var presetData = oldPlayerPerksData[presetName];

        if (!presetData || !presetData.order) return;

        var orderIndex = presetData.order - 1;
        if (orderIndex < 0 || orderIndex >= headerColIndices.length) {
          return;
        }
        var colIndex = headerColIndices[orderIndex];

        var headerCell =
          shared.columnToLetter(colIndex + 1) + (headerRowIndex + 1);
        batchUpdate.push({
          range: sheetName + "!" + headerCell,
          values: [[presetName]],
        });

        var perksStartRow = headerRowIndex + 3;

        if (presetData.bannedAmount && presetData.bannedAmount > 0) {
          var bannedAmountCell =
            shared.columnToLetter(colIndex + 3) + (headerRowIndex + 2);
          batchUpdate.push({
            range: sheetName + "!" + bannedAmountCell,
            values: [[presetData.bannedAmount]],
          });
        }

        if (presetData.perks && presetData.perks.length > 0) {
          var perksData = presetData.perks.map(function (perk) {
            return [perk];
          });

          var startCell =
            shared.columnToLetter(colIndex + 2) + (perksStartRow + 1);
          var endCell =
            shared.columnToLetter(colIndex + 2) +
            (perksStartRow + perksData.length);

          batchUpdate.push({
            range: sheetName + "!" + startCell + ":" + endCell,
            values: perksData,
          });
        }
      });

      if (batchUpdate.length !== 0) {
        return {
          success: true,
          message: "Player perks data updated successfully",
          batchUpdate: batchUpdate,
        };
      }

      return {
        success: true,
        message: "No updates needed for player perks data",
      };
    } catch (error) {
      var errorReport = errors.report("ranges.updatePlayerPerksPreset", error, {
        sheetName: sheetName,
        oldPlayerPerksData: oldPlayerPerksData,
        shouldRemoveUsedPerks: shouldRemoveUsedPerks,
        perksSheetData: perksSheetData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Player_&_Stuff data from a v4.2 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_2: function (oldSheetID) {
    try {
      console.log("Called: playerStuff.version4_2");

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
      console.log("Called: playerStuff.version4_0");

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
      console.log("Called: playerStuff.version3_2");

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
      console.log("Called: playerStuff.version2_0");

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
      console.log("Called: playerStuff.getVersion4_0PlayerStuffTiers");

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
      console.log("Called: playerStuff.getVersion2_0PlayerStuffTiers");

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
      console.log("Called: playerStuff.getVersion3_2PlayerStuffStats");

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
      console.log("Called: playerStuff.getVersion2_0PlayerStuffStats");

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
      console.log("Called: playerStuff.getVersion4_2PlayerStuffPerks");
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

      var presetOrder = shared.resolvePresetOrder(
        oldPerkPresetNameIdxs.map(function (colIdx) {
          return row[colIdx];
        }),
        shared.templatePresetNames,
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

  /**
   * Parses Player_&_Stuff data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parsePlayerStuffData: function (data) {
    try {
      const tourneyNames = {
        0: "Copper",
        1: "Silver",
        2: "Gold",
        3: "Platinum",
        4: "Champions",
        5: "Legends",
      };

      const tourneyName = data.tourneyID ? tourneyNames[data.tourneyID] : null;
      const highestWavePerTier = data.highestWavePerTier || [];
      const premiumPass = data.premiumPass || [];
      const atkDissonance = data.atkDissonance || [];
      const hpDissonance = data.hpDissonance || [];
      const coinDissonance = data.coinDissonance || [];
      const uwDissonance = data.uwDissonance || [];

      const lifetimeCoins = data.totalCoinsEarned || 0;
      const lifetimeStones = data.totalStonesEarned + data.totalStonesBought || 0;
      const lifetimeGems = data.totalGemsEarned + data.totalGemsBought || 0;
      const lifetimeKeys = data.totalKeysEarned || 0;

      const battleHistory = data.battleHistory || [];

      /**
       * Formats a lifetime total with a magnitude suffix.
       * @param {*} value
       * @returns {string|null} Null when there is no usable number.
       */
      function formatLifeTime(value) {
        if (value === null || value === undefined || value === "") {
          return null;
        }
        if (typeof value !== "number" || !isFinite(value)) {
          return value;
        }

        /**
         * The magnitude suffix for a thousands group: K, M, B, then aa, ab, …
         * @param {number} group
         * @returns {string}
         */
        function getSuffix(group) {
          var named = ["", "K", "M", "B", "T", "q", "Q", "s", "S", "O", "N", "D"];
          if (group < named.length) {
            return named[group];
          }
          var n = group - named.length;
          var first = Math.floor(n / 26);
          var second = n % 26;
          return (
            String.fromCharCode(97 + first) + String.fromCharCode(97 + second)
          );
        }

        var negative = value < 0;
        var num = Math.abs(value);

        if (num < 1000) {
          return (negative ? "-" : "") + String(Math.round(num * 100) / 100);
        }

        var group = Math.floor(Math.log10(num) / 3);
        var mantissa = num / Math.pow(1000, group);

        if (mantissa >= 1000) {
          mantissa /= 1000;
          group += 1;
        }
        var rounded = Math.round(mantissa * 100) / 100;
        if (rounded >= 1000) {
          rounded /= 1000;
          group += 1;
        }

        return (negative ? "-" : "") + rounded.toFixed(2) + getSuffix(group);
      }

      var allBattlesCoinPerHour = battleHistory
        .map(function (battle) {
          if (battle && battle.coinsEarned && battle.realTime) {
            var hours = battle.realTime / 3600;
            if (hours > 0) {
              return battle.coinsEarned / hours;
            }
          }
          return null;
        })
        .filter(function (cph) {
          return cph !== null;
        })
        .sort(function (a, b) {
          return b - a;
        });

      var numBattles = 3;
      const coinPerHour =
        allBattlesCoinPerHour.length > 0
          ? allBattlesCoinPerHour.slice(0, numBattles).reduce(function (
              sum,
              cph,
            ) {
              return sum + cph;
            }, 0) / Math.min(numBattles, allBattlesCoinPerHour.length)
          : null;

      var oldPlayerStuffTierData = {};
      var oldPlayerStuffStatsData = {
        Stat: {
          "Player ID": data.playerID,
          "Farming Tier": "Tier " + data.currentTier,
          "Tourney League": tourneyName,
          "Lifetime Coins": formatLifeTime(lifetimeCoins),
          Stones: formatLifeTime(lifetimeStones),
          Gems: formatLifeTime(lifetimeGems),
          Keys: formatLifeTime(lifetimeKeys),
          "Coin / Hour": formatLifeTime(coinPerHour),
        },
        "Premium Packs": {
          "Disable Ads": data.addPack,
          "Starter Pack": data.starterPack,
          "Epic Pack": data.epicPack,
        },
      };

      var nextPremium = 0;
      for (var tier = 0; tier < highestWavePerTier.length; tier++) {

        var wave = highestWavePerTier[tier];
        if (wave <= 0) {
          continue;
        }
        var premium = null;
        if (tier % 3 === 1) {
          premium = premiumPass[nextPremium] || null;
          nextPremium++;
        }

        oldPlayerStuffTierData["Tier " + tier] = {
          wave: wave,
          diss: {
            attack: atkDissonance[tier] || 0,
            defense: hpDissonance[tier] || 0,
            utility: coinDissonance[tier] || 0,
            ultimate: uwDissonance[tier] || 0,
          },
          premium: premium,
        };
      }

      return {
        success: true,
        oldPlayerStuffTierData: oldPlayerStuffTierData,
        oldPlayerStuffStatsData: oldPlayerStuffStatsData,
        statOrder: Object.keys(oldPlayerStuffStatsData.Stat),
        premiumOrder: Object.keys(oldPlayerStuffStatsData["Premium Packs"]),
      };
    } catch (error) {
      var errorReport = errors.report("playerStuff.parsePlayerStuffData", error, {
        data: data,
        oldPlayerStuffTierData: oldPlayerStuffTierData,
        oldPlayerStuffStatsData: oldPlayerStuffStatsData,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v2.0": this.version2_0.bind(this),
      "v3.2": this.version3_2.bind(this),
      "v4.0": this.version4_0.bind(this),
      "v4.2": this.version4_2.bind(this),
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
      return shared.compareVersions(b, a) === "newer" ? 1 : -1;
    });

    for (var i = 0; i < sortedThresholds.length; i++) {
      var threshold = sortedThresholds[i];
      var compareResult = shared.compareVersions(oldVersion, threshold);

      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }

    return null;
  },

};
