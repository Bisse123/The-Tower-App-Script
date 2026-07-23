const playerStuff = {
  // #region Export Functions
  exportData: function (versionDifference) {
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

      var oldDataResult = getVersionFunction();
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
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting player & stuff data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: playerStuff.importData");
      var newSpreadsheet = spreadsheets("Player & Stuff newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }

      // Batch get required data for update function only
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

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(
        newSheetID,
        "IDS",
        "IDS Master's",
        idsData,
      );
      if (
        !newSheetInfo ||
        !newSheetInfo.importStatus ||
        !newSheetInfo.importStatus.range
      ) {
        console.log(`Could not find import status range in IDS sheet`);
        return {
          success: false,
          message: "Could not find import status range in IDS sheet",
        };
      }

      var batchUpdate = [];

      // Only update player & stuff data if key exists
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

      // Add import status update to batch if there were data updates
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }

      // Always add ID updates
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
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing player & stuff data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Update Functions
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
      console.log(`Error in updatePlayerStuffData: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating player & stuff data: ${error.message}`,
      };
    }
  },

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
        if (
          perksSheetData[row].indexOf("Remove used perks from the pool") !== -1
        ) {
          var removeUsedPerksCol = shared.columnToLetter(
            perksSheetData[row].indexOf("Remove used perks from the pool"),
          );
          batchUpdate.push({
            range: sheetName + "!" + removeUsedPerksCol + (row + 1),
            values: [[shouldRemoveUsedPerks]],
          });
          continue;
        }
        var nonEmptyCells = perksSheetData[row].filter(function (cell) {
          return (
            cell !== null && cell !== undefined && String(cell).trim() !== ""
          );
        });

        if (nonEmptyCells.indexOf("Farming") !== -1) {
          headerRowIndex = row;

          nonEmptyCells.forEach(function (header) {
            var colIndex = perksSheetData[row].indexOf(header);
            if (colIndex !== -1) {
              headerColIndices.push(colIndex);
            }
          });
          break;
        }
      }

      if (headerRowIndex === -1) {
        console.log(`Could not find header row with "Farming"`);
        return {
          success: false,
          message: "Could not find header row with Farming",
        };
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
        if (orderIndex >= 0 && orderIndex < headerColIndices.length) {
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
      console.log(`Error in updatePlayerPerksPreset: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating player perks data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version4_2: function () {
    try {
      console.log("Called: playerStuff.version4_2");
      var oldSpreadsheet = spreadsheets("Player & Stuff oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }
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
      console.log("Error in version4_2: " + error.toString());
      return {
        success: false,
        message: "Error in version4_2: " + error.message,
      };
    }
  },

  version4_0: function () {
    try {
      console.log("Called: playerStuff.version4_0");
      var oldSpreadsheet = spreadsheets("Player & Stuff oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }
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
      console.log("Error in version4_0: " + error.toString());
      return {
        success: false,
        message: "Error in version4_0: " + error.message,
      };
    }
  },

  version3_2: function () {
    try {
      console.log("Called: playerStuff.version3_2");
      var oldSpreadsheet = spreadsheets("Player & Stuff oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }
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
      console.log("Error in version3_2: " + error.toString());
      return {
        success: false,
        message: "Error in version3_2: " + error.message,
      };
    }
  },

  version2_0: function () {
    try {
      console.log("Called: playerStuff.version2_0");
      var oldSpreadsheet = spreadsheets("Player & Stuff oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }
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
      console.log("Error in version2_0: " + error.toString());
      return {
        success: false,
        message: "Error in version2_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get PlayerStuff Tiers
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
      // Wave / dissonant cells may be display-formatted (e.g. "5,000"); strip
      // the thousands separators so values match the raw save file numbers.
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
      console.log(
        "Error in getVersion4_0PlayerStuffTiers: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion4_0PlayerStuffTiers: " + error.message,
      };
    }
  },

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
      // Strip thousands separators (e.g. "5,000") so wave values compare cleanly.
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
      console.log(
        "Error in getVersion2_0PlayerStuffTiers: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion2_0PlayerStuffTiers: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get PlayerStuff Stats
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
      console.log(
        "Error in getVersion3_2PlayerStuffStats: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion3_2PlayerStuffStats: " + error.message,
      };
    }
  },

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
      console.log(
        "Error in getVersion2_0PlayerStuffStats: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion2_0PlayerStuffStats: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get PlayerStuff Perks
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
      for (
        var rowIndex = 0;
        rowIndex < oldPlayerStuffPerksValues.length;
        rowIndex++
      ) {
        var row = oldPlayerStuffPerksValues[rowIndex];
        var colIndex = row.indexOf("Remove used perks from the pool");
        if (colIndex !== -1) {
          shouldRemoveUsedPerks =
            row[colIndex - 1] === "TRUE" ||
            row[colIndex - 1] === "true" ||
            row[colIndex - 1] === true;
        }

        var oldPerkPresetNameIdxs = row
          .map(function (cell, idx) {
            return String(cell || "").trim() !== "" ? idx : -1;
          })
          .filter(function (idx) {
            return idx !== -1;
          });

        if (row.indexOf("Farming") !== -1) {
          oldPerkPresetNameIdxs.forEach(function (colIdx, orderIndex) {
            var presetName = row[colIdx];
            if (!oldPerksPreset.hasOwnProperty(presetName)) {
              oldPerksPreset[presetName] = {
                perks: [],
                order: orderIndex + 1,
              };
            }
            var bannedAmount =
              oldPlayerStuffPerksValues[rowIndex + 1][colIdx + 2] || 0;
            oldPerksPreset[presetName].bannedAmount = bannedAmount;
            for (
              var rowIdx = rowIndex + 3;
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
          break;
        }
      }
      return {
        success: true,
        message: "Player & Stuff perks processed successfully",
        oldPerksPreset: oldPerksPreset,
        shouldRemoveUsedPerks: shouldRemoveUsedPerks,
      };
    } catch (error) {
      console.log(
        "Error in getVersion4_2PlayerStuffPerks: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion4_2PlayerStuffPerks: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parsePlayerStuffData: function (data) {
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

    function formatLifeTime(value) {
      if (value === null || value === undefined || value === "") {
        return null;
      }
      if (typeof value !== "number" || !isFinite(value)) {
        return value;
      }

      // Suffix for a given power-of-1000 group.
      // 0 -> "", 1 -> K, 2 -> M, ... 11 -> D, then aa, ab, ac, ...
      function getSuffix(group) {
        var named = [
          "",
          "K",
          "M",
          "B",
          "T",
          "q",
          "Q",
          "s",
          "S",
          "O",
          "N",
          "D",
        ];
        if (group < named.length) {
          return named[group];
        }
        var n = group - named.length; // 0-based index into aa, ab, ac, ...
        var first = Math.floor(n / 26);
        var second = n % 26;
        return (
          String.fromCharCode(97 + first) + String.fromCharCode(97 + second)
        );
      }

      var negative = value < 0;
      var num = Math.abs(value);

      // Below 1000 there is no suffix.
      if (num < 1000) {
        return (negative ? "-" : "") + String(Math.round(num * 100) / 100);
      }

      var group = Math.floor(Math.log10(num) / 3);
      var mantissa = num / Math.pow(1000, group);

      // Guard against floating point / rounding pushing the mantissa to 1000+.
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
    
    var allBattlesCoinPerHour = battleHistory.map(function (battle) {
      if (battle && battle.coinsEarned && battle.realTime) {
        var hours = battle.realTime / 3600;
        if (hours > 0) {
          return battle.coinsEarned / hours;
        }
      }
      return null;
    }).filter(function (cph) {
      return cph !== null;
    }).sort(function (a, b) {
      return b - a;
    });

    var numBattles = 3;
    const coinPerHour = allBattlesCoinPerHour.length > 0 ? allBattlesCoinPerHour.slice(0, numBattles).reduce(function (sum, cph) {
      return sum + cph;
    }, 0) / Math.min(numBattles, allBattlesCoinPerHour.length) : null;
    
    var oldPlayerStuffTierData = {};
    var oldPlayerStuffStatsData = {
      Stat: {
        "Player ID": data.playerID,
        "Farming Tier": "Tier " + data.currentTier,
        "Tourney League": tourneyName,
        "Lifetime Coins": formatLifeTime(lifetimeCoins),
        "Stones": formatLifeTime(lifetimeStones),
        "Gems": formatLifeTime(lifetimeGems),
        "Keys": formatLifeTime(lifetimeKeys),
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
      var wave = Math.min(4500, highestWavePerTier[tier]);
      if (wave <= 0) {
        continue;
      }
      var premium = null;
      if (tier % 3 === 1) {
        premium = premiumPass[nextPremium] || null;
        nextPremium++;
      }

      oldPlayerStuffTierData["Tier " + (tier)] = {
        wave: wave,
        diss: {
          attack: Math.min(5000, atkDissonance[tier] || 0),
          defense: Math.min(5000, hpDissonance[tier] || 0),
          utility: Math.min(5000, coinDissonance[tier] || 0),
          ultimate: Math.min(5000, uwDissonance[tier] || 0),
        },
        premium: premium,
      };
    }

    return {
      oldPlayerStuffTierData: oldPlayerStuffTierData,
      oldPlayerStuffStatsData: oldPlayerStuffStatsData,
      statOrder: Object.keys(oldPlayerStuffStatsData.Stat),
      premiumOrder: Object.keys(oldPlayerStuffStatsData["Premium Packs"]),
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v2.0": this.version2_0.bind(this),
      "v3.2": this.version3_2.bind(this),
      "v4.0": this.version4_0.bind(this),
      "v4.2": this.version4_2.bind(this),
    };
  },

  // #endregion
  // #region Compatibility Check
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
  
  // #endregion
};
