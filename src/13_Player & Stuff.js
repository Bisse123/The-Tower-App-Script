const playerStuff = {
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
      var requiredRanges = ["Master Sheet", "IDS"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var idsData = batchResults[1].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(
        newSheetID,
        "IDS",
        "IDS Master's",
        idsData
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
      if (data.hasOwnProperty("oldPlayerStuffTierData") && data.hasOwnProperty("oldPlayerStuffStatsData")) {
        var oldPlayerStuffTierData = data.oldPlayerStuffTierData;
        var oldPlayerStuffStatsData = data.oldPlayerStuffStatsData;
        var playerResult = this.updatePlayerStuffData(
          "Master Sheet",
          oldPlayerStuffTierData,
          oldPlayerStuffStatsData,
          masterSheetData
        );
        if (!playerResult || !playerResult.success) {
          console.log(`Error updating player data: ${playerResult.message}`);
          return playerResult;
        }
        batchUpdate = batchUpdate.concat(playerStuffResult.batchUpdate || []);
      }

      // Add import status update to batch if there were data updates
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(batchUpdate, "Player & Stuff", newSheetID, idsData, data.idMasterID);

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

  updatePlayerStuffData: function (sheetName, oldPlayerTierData, oldPlayerStatsData, masterSheetData) {
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

      if (statCol === -1 || tierCol === -1) {
        console.log(`Stat or Tier column not found in master sheet`);
        return {
          success: false,
          message: "Stat or Tier column not found in master sheet",
        };
      }
      var header = headerRow[statCol] || "";
      var perkRow = -1;
      var values = {
        Stat: [],
        Tier: [],
        "Premium Packs": [],
      };
      for (var row = 1; row < masterSheetData.length; row++) {
        var rowData = masterSheetData[row];
        var statName = rowData[statCol] || "";
        var tierValue = rowData[tierCol] || "";
        if (!tierValue) {
          break;
        }

        if (oldPlayerTierData && oldPlayerTierData[tierValue]) {
          var wave = oldPlayerTierData[tierValue].wave || "";
          var premium = oldPlayerTierData[tierValue].premium || "";
          values["Tier"].push([wave, premium]);
        }

        if (!statName) {
          continue;
        }

        if (statName === "Premium Packs") {
          header = "Premium Packs";
          perkRow = row + 2;
        }

        if (oldPlayerStatsData[header] && oldPlayerStatsData[header][statName]) {
          var value = oldPlayerStatsData[header][statName].value || "";
          values[header].push([value]);
        }
      }

      var statColLetter = shared.columnToLetter(statCol + 2);
      var tierStartColLetter = shared.columnToLetter(tierCol + 2);
      var tierEndColLetter = shared.columnToLetter(tierCol + 3);
      var batchUpdate = [];
      var ranges = {
        Stat: `${sheetName}!${statColLetter}2:${statColLetter}${
          2 + values.Stat.length - 1
        }`,
        Tier: `${sheetName}!${tierStartColLetter}2:${tierEndColLetter}${
          2 + values.Tier.length - 1
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

  version32: function () {
    try {
      console.log("Called: playerStuff.version32");
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
      if (!batchResult || batchResult.length ===0) {
        console.log(`Could not read old player & stuff data`);
        return {
          success: false,
          message: `Could not read old player & stuff data`,
        };
      }
      var oldPlayerStuffTierValues = batchResult[0].values;
      var oldPlayerStuffStatsValues = batchResult[1].values;
      var tierDataResult = this.getVersion20PlayerStuffTiers(oldPlayerStuffTierValues);
      var statsDataResult = this.getVersion32PlayerStuffStats(oldPlayerStuffStatsValues);
      success = tierDataResult.success && statsDataResult.success;
      return {
        success: success,
        message: success ? "Player & Stuff processed successfully" : "Error processing Player & Stuff data",
        oldPlayerStuffTierData: tierDataResult.oldPlayerStuffTierData,
        oldPlayerStuffStatsData: statsDataResult.oldPlayerStuffStatsData,
      };

    } catch (error) {
      console.log("Error in version32: " + error.toString());
      return {
        success: false,
        message: "Error in version32: " + error.message,
      };
    }
  },

  version20: function () {
    try {
      console.log("Called: playerStuff.version20");
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
      if (!batchResult || batchResult.length ===0) {
        console.log(`Could not read old player & stuff data`);
        return {
          success: false,
          message: `Could not read old player & stuff data`,
        };
      }
      var oldPlayerStuffTierValues = batchResult[0].values;
      var oldPlayerStuffStatsValues = batchResult[1].values;
      var tierDataResult = this.getVersion20PlayerStuffTiers(oldPlayerStuffTierValues);
      var statsDataResult = this.getVersion20PlayerStuffStats(oldPlayerStuffStatsValues);
      success = tierDataResult.success && statsDataResult.success;
      return {
        success: success,
        message: success ? "Player & Stuff processed successfully" : "Error processing Player & Stuff data",
        oldPlayerStuffTierData: tierDataResult.oldPlayerStuffTierData,
        oldPlayerStuffStatsData: statsDataResult.oldPlayerStuffStatsData,
      };

    } catch (error) {
      console.log("Error in version20: " + error.toString());
      return {
        success: false,
        message: "Error in version20: " + error.message,
      };
    }
  },

  getVersion20PlayerStuffTiers: function (oldPlayerStuffTierValues) {
    try {
      console.log("Called: playerStuff.getversion20PlayerStuffTiers");

      if (!oldPlayerStuffTierValues || oldPlayerStuffTierValues.length === 0) {
        console.log(`No data found in old player & stuff tier data`);
        return {
          success: false,
          message: "No data found in old player & stuff tier data",
        };
      }
      var oldPlayerStuffTierData = {};
      for (var row = 0; row < oldPlayerStuffTierValues.length; row++) {
        var rowData = oldPlayerStuffTierValues[row];
        var tier = rowData[0] || "";
        var wave = rowData[1] || "";
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
      console.log("Error in getversion20PlayerStuffTiers: " + error.toString());
      return {
        success: false,
        message: "Error in getversion20PlayerStuffTiers: " + error.message,
      };
    }
  },

  getVersion20PlayerStuffStats: function (oldPlayerStuffStatsValues) {
    try {
      console.log("Called: playerStuff.getversion20PlayerStuffStats");

      if (!oldPlayerStuffStatsValues || oldPlayerStuffStatsValues.length === 0) {
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
        if (name) {
          oldPlayerStuffStatsData[header][name] = {
            value: value,
          };
        }
      }
      return {
        success: true,
        message: "Player & Stuff stats processed successfully",
        oldPlayerStuffStatsData: oldPlayerStuffStatsData,
      };
    } catch (error) {
      console.log("Error in getversion20PlayerStuffStats: " + error.toString());
      return {
        success: false,
        message: "Error in getversion20PlayerStuffStats: " + error.message,
      };
    }
  },

  getVersion32PlayerStuffStats: function (oldPlayerStuffStatsValues) {
    try {
      console.log("Called: playerStuff.getVersion32PlayerStuffStats");

      if (!oldPlayerStuffStatsValues || oldPlayerStuffStatsValues.length === 0) {
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
        if (name) {
          oldPlayerStuffStatsData[header][name] = {
            value: value,
          };
        }
      }
      return {
        success: true,
        message: "Player & Stuff stats processed successfully",
        oldPlayerStuffStatsData: oldPlayerStuffStatsData,
      };
    } catch (error) {
      console.log("Error in getversion32PlayerStuffStats: " + error.toString());
      return {
        success: false,
        message: "Error in getversion32PlayerStuffStats: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v2.0": this.version20.bind(this),
      "v3.2": this.version32.bind(this),
    };
  },

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
