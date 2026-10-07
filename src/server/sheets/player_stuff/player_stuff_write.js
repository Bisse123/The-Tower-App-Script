const playerStuffWriter = {

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
      console.log("Called: playerStuffWriter.updatePlayerStuffData");
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

      var statColLetter = sheetRefs.columnToLetter(statCol + 2);
      var tierColLetter = sheetRefs.columnToLetter(tierCol + 2);
      var passColLetter = sheetRefs.columnToLetter(passCol + 1);
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
          var dissColLetter = sheetRefs.columnToLetter(
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
      var errorReport = errors.report("playerStuffWriter.updatePlayerStuffData", error, {
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
      console.log("Called: playerStuffWriter.updatePlayerPerksPreset");
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
            sheetRefs.columnToLetter(removeUsedPerksIndex) +
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
          sheetRefs.columnToLetter(colIndex + 1) + (headerRowIndex + 1);
        batchUpdate.push({
          range: sheetName + "!" + headerCell,
          values: [[presetName]],
        });

        var perksStartRow = headerRowIndex + 3;

        if (presetData.bannedAmount && presetData.bannedAmount > 0) {
          var bannedAmountCell =
            sheetRefs.columnToLetter(colIndex + 3) + (headerRowIndex + 2);
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
            sheetRefs.columnToLetter(colIndex + 2) + (perksStartRow + 1);
          var endCell =
            sheetRefs.columnToLetter(colIndex + 2) +
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
};
