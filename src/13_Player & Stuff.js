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
      if (data.hasOwnProperty("oldPlayerStuffData")) {
        var oldPlayerStuffData = data.oldPlayerStuffData;
        var playerResult = this.updatePlayerStuffData(
          "Master Sheet",
          oldPlayerStuffData,
          masterSheetData
        );
        if (!playerResult || !playerResult.success) {
          console.log(`Error updating player data: ${playerResult.message}`);
          return playerResult;
        }
        batchUpdate = batchUpdate.concat(playerResult.batchUpdate || []);
      }

      // Add import status update to batch if any update was made
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });

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
      }
      return {
        success: true,
        message: "No player & stuff data to update",
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing player & stuff data: ${error.message}`,
      };
    }
  },

  updatePlayerStuffData: function (sheetName, oldPlayerData, masterSheetData) {
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
        "Premium Perk": [],
      };
      for (var row = 1; row < masterSheetData.length; row++) {
        var rowData = masterSheetData[row];
        var statName = rowData[statCol] || "";
        var tierValue = rowData[tierCol] || "";
        if (!tierValue) {
          break;
        }

        if (oldPlayerData["Tier"] && oldPlayerData["Tier"][tierValue]) {
          var wave = oldPlayerData["Tier"][tierValue].value || "";
          var premium = oldPlayerData["Tier"][tierValue].premium || "";
          values["Tier"].push([wave, premium]);
        }

        if (!statName) {
          continue;
        }

        if (statName === "Premium Perk") {
          header = "Premium Perk";
          perkRow = row + 2;
        }

        if (oldPlayerData[header] && oldPlayerData[header][statName]) {
          var value = oldPlayerData[header][statName].value || "";
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
        "Premium Perk": `${sheetName}!${statColLetter}${perkRow}:${statColLetter}${
          perkRow + values["Premium Perk"].length - 1
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

      // TODO: Define appropriate ranges for player & stuff data
      var playerStuffRange = "EXPORT!B2:D";
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, [
        playerStuffRange,
      ]);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log(`Could not read old player & stuff data`);
        return {
          success: false,
          message: `Could not read old player & stuff data`,
        };
      }
      var oldPlayerStuffValues = batchResult[0].values;

      var playerStuffData = this.getVersion20PlayerStuff(oldPlayerStuffValues);
      return playerStuffData;
    } catch (error) {
      console.log("Error in version20: " + error.toString());
      return {
        success: false,
        message: "Error in version20: " + error.message,
      };
    }
  },

  getVersion20PlayerStuff: function (oldPlayerStuffValues) {
    try {
      console.log("Called: playerStuff.getVersion20PlayerStuff");
      var headers = ["Stat", "Premium Perk", "Tier"];

      if (!oldPlayerStuffValues || oldPlayerStuffValues.length === 0) {
        console.log(`No data found in old player & stuff data`);
        return {
          success: false,
          message: "No data found in old player & stuff data",
        };
      }
      var header = headers[0];
      var oldPlayerStuffData = {};
      oldPlayerStuffData[header] = {};
      for (var row = 0; row < oldPlayerStuffValues.length; row++) {
        var rowData = oldPlayerStuffValues[row];
        var name = rowData[0] || "";
        var value = rowData[1] || "";
        var premium = rowData[2] || "";
        if (headers.includes(name)) {
          header = name;
          oldPlayerStuffData[header] = {};
          continue;
        }
        if (header === "Premium Perk" && name === "Coin Multiplier") {
          continue;
        }
        if (name) {
          oldPlayerStuffData[header][name] = {
            value: value,
            premium: premium,
          };
        }
      }

      return {
        success: true,
        message: "Player & Stuff processed successfully",
        oldPlayerStuffData: oldPlayerStuffData,
      };
    } catch (error) {
      console.log("Error in getVersion20PlayerStuff: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion20PlayerStuff: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v2.0": this.version20.bind(this),
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
