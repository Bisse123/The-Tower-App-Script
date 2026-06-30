const relics = {
  // #region Export Functions
  exportData: function (versionDifference) {
    try {
      console.log("Called: relics.exportData");
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
        message: "Relics export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting relics data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: relics.importData");
      var newSpreadsheet = spreadsheets("Relics newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      var requiredRanges = ["Relics", "IDS"];
      var newRelicsBatchResult = SheetsAPI.batchGetValues(
        newSheetID,
        requiredRanges
      );
      if (!newRelicsBatchResult || newRelicsBatchResult.length === 0) {
        console.log("Error getting relics sheet data");
        return {
          success: false,
          message: "Error getting relics sheet data",
        };
      }

      var newRelicsData = newRelicsBatchResult[0].values;
      var idsData = newRelicsBatchResult[1].values;

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

      // Only update relics if key exists
      if (data.hasOwnProperty("oldRelics")) {
        var oldRelics = data.oldRelics;
        var relicsResult = this.updateRelics(
          "Relics",
          oldRelics,
          newRelicsData
        );
        if (!relicsResult || !relicsResult.success) {
          console.log(`Error updating relics: ${relicsResult.message}`);
          return relicsResult;
        }
        batchUpdate = batchUpdate.concat(relicsResult.batchUpdate || []);
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
        "Relics",
        newSheetID,
        idsData,
        data.idMasterID
      );

      // Apply all updates (including ID setting and import status)
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
        message: `Relics import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing relics data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Update Functions
  updateRelics: function (sheetName, oldRelics, newRelicsData) {
    try {
      console.log("Called: relics.updateRelics");
      if (!newRelicsData || newRelicsData.length < 3) {
        console.log(`Not enough data in new Relics sheet`);
        return {
          success: false,
          message: `Not enough data in new Relics sheet`,
        };
      }

      var newRelicHeaderRow = null;
      var newRelicNameCol = null;
      var newRelicUnlockedCol = null;

      // Scan each row to find the header
      for (var row = 0; row < newRelicsData.length; row++) {
        var rowValues = newRelicsData[row];
        var relicNameIndex = rowValues.indexOf("Relic Name");
        var relicUnlockedIndex = rowValues.indexOf("Unlocked");
        if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
          newRelicHeaderRow = row + 1;
          newRelicNameCol = relicNameIndex + 1;
          newRelicUnlockedCol = relicUnlockedIndex + 1;
          break;
        }
      }

      if (!newRelicHeaderRow) {
        console.log(`Could not find header row in new Relics sheet`);
        return {
          success: false,
          message: `Could not find header row in new Relics sheet`,
        };
      }

      var startRow = newRelicHeaderRow + 1;

      // Build unlocked status array directly by iterating through new relics data
      var newRelicsUnlocked = [];
      newRelicsData.slice(startRow - 1).forEach(function (row) {
        var relicName = (row[newRelicNameCol - 1] || "").trim();
        if (String(relicName).trim() !== "") {
          if (oldRelics.includes(relicName)) {
            newRelicsUnlocked.push([true]);
          } else {
            newRelicsUnlocked.push([false]);
          }
        }
      });
      if (newRelicsUnlocked.length > 0) {
        var endRow = startRow + newRelicsUnlocked.length - 1;
        var unlockedRange = `${sheetName}!${shared.columnToLetter(
          newRelicUnlockedCol
        )}${startRow}:${shared.columnToLetter(newRelicUnlockedCol)}${endRow}`;

        var batchUpdate = [
          {
            range: unlockedRange,
            values: newRelicsUnlocked,
          },
        ];
        return {
          success: true,
          message: `Relics updated successfully: ${newRelicsUnlocked.length} relics processed`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for relics`,
      };
    } catch (error) {
      console.log("Error in updateRelics: " + error.toString());
      return {
        success: false,
        message: "Error updating relics: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version1_0: function () {
    try {
      console.log("Called: relics.version1_0");
      var oldSpreadsheet = spreadsheets("Relics oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Check if Relics sheet exists in old spreadsheet
      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "Relics")) {
        console.log("Relics sheet not found in old relic spreadsheet");
        return {
          success: false,
          message: `Relics sheet not found in old relic spreadsheet™`,
        };
      }

      var oldRelicsBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        "Relics",
      ]);
      if (
        !oldRelicsBatchResult ||
        oldRelicsBatchResult.length === 0 ||
        !oldRelicsBatchResult[0].values
      ) {
        console.log(`Could not read data from old Relics sheet`);
        return {
          success: false,
          message: `Could not read data from old Relics sheet`,
        };
      }
      var oldRelicsData = oldRelicsBatchResult[0].values;

      var relicsData = this.getVersion1_0Relics(oldRelicsData);
      return relicsData;
    } catch (error) {
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Relics
  getVersion1_0Relics: function (oldRelicsData) {
    try {
      console.log("Called: relics.getVersion1_0Relics");
      var oldRelicHeaderRow = -1;
      var relicNameIndex = -1;
      var relicUnlockedIndex = -1;

      // Scan each row to find the header
      for (var row = 0; row < oldRelicsData.length; row++) {
        var rowValues = oldRelicsData[row];
        relicNameIndex = rowValues.indexOf("Relic Name");
        relicUnlockedIndex = rowValues.indexOf("Unlocked");
        if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
          oldRelicHeaderRow = row + 1;
          break;
        }
      }

      if (oldRelicHeaderRow === -1) {
        console.log(`Could not find header row in old Relics sheet`);
        return {
          success: false,
          message: `Could not find header row in old Relics sheet`,
        };
      }

      var startRow = oldRelicHeaderRow + 1;

      var oldRelics = [];
      oldRelicsData.slice(startRow - 1).forEach(function (row) {
        var relicName = row[relicNameIndex].trim();
        if (relicName.includes("T:")) {
          relicName = relicName.replace(/T:\s+/g, "T: ");
        }
        var isUnlocked = row[relicUnlockedIndex];

        if (
          relicName &&
          (isUnlocked === true ||
            isUnlocked === "TRUE" ||
            isUnlocked === "true")
        ) {
          oldRelics.push(relicName);
        }
      });

      return {
        success: true,
        oldRelics: oldRelics,
      };
    } catch (error) {
      console.log("Error in getVersion1_0Relics: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0Relics: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseRelicsData: function (data, sheetID) {
    var relicsSheetData = null;
    if (sheetID) {
      var batchResult = SheetsAPI.batchGetValues(sheetID, ["Relics"]);
      if (batchResult && batchResult[0] && batchResult[0].values) {
        relicsSheetData = batchResult[0].values;
      }
    }

    if (!relicsSheetData) {
      console.log("Could not read Relics sheet to map relic indices");
      return {
        oldRelics: [],
      };
    }

    var headerRow = -1;
    var relicNameCol = -1;
    var ingameIdCol = -1;
    for (var headerScan = 0; headerScan < relicsSheetData.length; headerScan++) {
      var headerValues = relicsSheetData[headerScan];
      var nameIndex = headerValues.indexOf("Relic Name");
      var idIndex = headerValues.indexOf("Ingame_ID");
      if (nameIndex !== -1 && idIndex !== -1) {
        headerRow = headerScan;
        relicNameCol = nameIndex;
        ingameIdCol = idIndex;
        break;
      }
    }

    if (headerRow === -1) {
      console.log(
        "Could not find 'Relic Name' and 'Ingame_ID' headers in Relics sheet"
      );
      return {
        oldRelics: [],
      };
    }

    var relicNameByIndex = {};
    for (var r = headerRow + 1; r < relicsSheetData.length; r++) {
      var relicRow = relicsSheetData[r];
      var mappedRelicName = (relicRow[relicNameCol] || "").toString().trim();
      var ingameId = relicRow[ingameIdCol];
      if (
        mappedRelicName === "" ||
        ingameId === "" ||
        ingameId === null ||
        ingameId === undefined
      ) {
        continue;
      }
      var relicIndex = parseInt(ingameId, 10);
      if (!isNaN(relicIndex)) {
        relicNameByIndex[relicIndex] = mappedRelicName;
      }
    }

    const relicsData = data.relicsUnlocked || [];

    var oldRelics = [];
    relicsData.forEach(function (relicStatus, index) {
      if (relicStatus > 0) {
        var relicName = relicNameByIndex[index];
        if (relicName) {
          oldRelics.push(relicName);
        }
      }
    });

    return {
      oldRelics: oldRelics,
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
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
