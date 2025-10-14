const vault = {
  exportData: function (versionDifference) {
    try {
      console.log("Called: vault.exportData");
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
        message: "Vault export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting vault data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      console.log("Called: vault.importData");
      var newSpreadsheet = spreadsheets("Vault newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      var requiredRanges = ["Harmony", "Power", "IDS"];
      var newVaultBatchResult = SheetsAPI.batchGetValues(
        newSheetID,
        requiredRanges
      );
      if (!newVaultBatchResult || newVaultBatchResult.length < 2) {
        console.log("Error getting vault sheet data");
        return {
          success: false,
          message: "Error getting vault sheet data",
        };
      }

      var harmonyData = newVaultBatchResult[0].values;
      var powerData = newVaultBatchResult[1].values;
      var idsData = newVaultBatchResult[2].values;

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

      // Only update Harmony if key exists
      if (data.hasOwnProperty("oldVaultHarmony")) {
        var oldVaultHarmony = data.oldVaultHarmony;
        var harmonyResult = this.updateVault(
          "Harmony",
          harmonyData,
          oldVaultHarmony
        );
        if (!harmonyResult || !harmonyResult.success) {
          console.log(`Error updating Harmony vault: ${harmonyResult.message}`);
          return harmonyResult;
        }
        batchUpdate = batchUpdate.concat(harmonyResult.batchUpdate || []);
      }

      // Only update Power if key exists
      if (data.hasOwnProperty("oldVaultPower")) {
        var oldVaultPower = data.oldVaultPower;
        var powerResult = this.updateVault("Power", powerData, oldVaultPower);
        if (!powerResult || !powerResult.success) {
          console.log(`Error updating Power vault: ${powerResult.message}`);
          return powerResult;
        }
        batchUpdate = batchUpdate.concat(powerResult.batchUpdate || []);
      }

      // Add import status update to batch if there were data updates
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(batchUpdate, "Vault", newSheetID, idsData);

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
        message: `Vault import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing vault data: ${error.message}`,
      };
    }
  },

  updateVault: function (sheetName, newVaultData, oldVault) {
    try {
      console.log("Called: vault.updateVault");
      if (!newVaultData) {
        console.log(`Error getting sheet data - no pre-fetched data available`);
        return {
          success: false,
          message: `Error getting sheet data`,
        };
      }

      if (newVaultData.length < 2) {
        console.log(`Not enough data in sheet`);
        return {
          success: false,
          message: `Not enough data in sheet`,
        };
      }

      // Find the header row dynamically
      var newVaultHeaders = null;
      var headerRowIndex = -1;

      for (var i = 0; i < newVaultData.length; i++) {
        var row = newVaultData[i];

        var uIdx = row.indexOf("U");
        var valueIdx = row.indexOf("Value");
        var bonusTypeIdx = row.indexOf("Bonus Type");
        if (uIdx !== -1 && valueIdx !== -1 && bonusTypeIdx !== -1) {
          newVaultHeaders = row;
          headerRowIndex = i;
          break;
        }
      }

      if (!newVaultHeaders || headerRowIndex === -1) {
        console.log(`Could not find header pattern in sheet data`);
        return {
          success: false,
          message: `Could not find header pattern in sheet data`,
        };
      }

      var vaultData = newVaultData.slice(headerRowIndex + 1);

      // Find all occurrences of each column type and group them
      var uIndices = [];
      var valueIndices = [];
      var bonusTypeIndices = [];

      for (var col = 0; col < newVaultHeaders.length; col++) {
        if (newVaultHeaders[col] === "U") {
          uIndices.push(col);
        } else if (newVaultHeaders[col] === "Value") {
          valueIndices.push(col);
        } else if (newVaultHeaders[col] === "Bonus Type") {
          bonusTypeIndices.push(col);
        }
      }

      // Pair them together by position
      var columnGroups = [];
      for (var i = 0; i < uIndices.length; i++) {
        if (i < valueIndices.length && i < bonusTypeIndices.length) {
          columnGroups.push({
            uIdx: uIndices[i],
            valueIdx: valueIndices[i],
            bonusTypeIdx: bonusTypeIndices[i],
          });
        }
      }

      var newVault = {};
      var batchUpdate = [];

      for (var r = 0; r < vaultData.length; r++) {
        var row = vaultData[r];
        for (var g = 0; g < columnGroups.length; g++) {
          var group = columnGroups[g];
          var uIdx = group.uIdx;
          var valueIdx = group.valueIdx;
          var bonusTypeIdx = group.bonusTypeIdx;
          var u = row[uIdx];
          var value = row[valueIdx];
          var bonusType = row[bonusTypeIdx];
          var key = bonusType || value;
          if (oldVault.hasOwnProperty(key)) {
            u = oldVault[key].shift();
            if (oldVault[key].length === 0) {
              delete oldVault[key];
            }
          }
          if (!newVault.hasOwnProperty(uIdx)) {
            newVault[uIdx] = [];
          }
          newVault[uIdx].push([u]);
        }
      }

      Object.keys(newVault).forEach(function (colKey) {
        var colIdx = parseInt(colKey, 10);
        var colLetter = shared.columnToLetter(colIdx + 1);
        var values = newVault[colKey];
        var startRow = headerRowIndex + 2;
        var lastRow = startRow + values.length - 1;
        var range = `${sheetName}!${colLetter}${startRow}:${colLetter}${lastRow}`;
        batchUpdate.push({
          range: range,
          values: values,
        });
      });

      if (batchUpdate.length !== 0) {
        return {
          success: true,
          message: `Vault updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for vault`,
      };
    } catch (error) {
      console.log("Error in updateVault: " + error.toString());
      return {
        success: false,
        message: "Error in updateVault: " + error.message,
      };
    }
  },

  version10: function () {
    try {
      console.log("Called: vault.version10");
      var oldSpreadsheet = spreadsheets("Vault oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var oldVaultBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        "Harmony",
        "Power",
      ]);
      if (!oldVaultBatchResult || oldVaultBatchResult.length < 2) {
        console.log(`Error getting old vault sheet data`);
        return {
          success: false,
          message: `Error getting old vault sheet data`,
        };
      }

      var harmonyData =
        oldVaultBatchResult[0] && oldVaultBatchResult[0].values
          ? oldVaultBatchResult[0].values
          : [];
      var powerData =
        oldVaultBatchResult[1] && oldVaultBatchResult[1].values
          ? oldVaultBatchResult[1].values
          : [];

      var harmonyResult = this.getVersion10Vault(harmonyData);
      if (!harmonyResult || !harmonyResult.success) {
        console.log(
          `Error getting harmony vault data: ${harmonyResult.message}`
        );
        return harmonyResult;
      }

      var powerResult = this.getVersion10Vault(powerData);
      if (!powerResult || !powerResult.success) {
        console.log(`Error getting power vault data: ${powerResult.message}`);
        return powerResult;
      }

      return {
        success: true,
        message: "Vault processed successfully",
        oldVaultHarmony: harmonyResult.oldVault,
        oldVaultPower: powerResult.oldVault,
      };
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10Vault: function (oldSheetData) {
    try {
      console.log("Called: vault.getVersion10Vault");
      if (!oldSheetData || oldSheetData.length === 0) {
        console.log(`No sheet data provided for vault`);
        return { success: false, message: `No sheet data provided for vault` };
      }

      var oldVaultHeaders = null;
      var headerRowIndex = -1;

      for (var i = 0; i < oldSheetData.length; i++) {
        var row = oldSheetData[i];

        var uIdx = row.indexOf("U");
        var valueIdx = row.indexOf("Value");
        var bonusTypeIdx = row.indexOf("Bonus Type");
        if (uIdx !== -1 && valueIdx !== -1 && bonusTypeIdx !== -1) {
          oldVaultHeaders = row;
          headerRowIndex = i;
          break;
        }
      }

      if (!oldVaultHeaders || headerRowIndex === -1) {
        console.log(`Could not find header pattern in sheet data`);
        return {
          success: false,
          message: `Could not find header pattern in sheet data`,
        };
      }

      var oldVaultData = oldSheetData.slice(headerRowIndex + 1);

      var uIndices = [];
      var valueIndices = [];
      var bonusTypeIndices = [];

      for (var col = 0; col < oldVaultHeaders.length; col++) {
        if (oldVaultHeaders[col] === "U") {
          uIndices.push(col);
        } else if (oldVaultHeaders[col] === "Value") {
          valueIndices.push(col);
        } else if (oldVaultHeaders[col] === "Bonus Type") {
          bonusTypeIndices.push(col);
        }
      }

      var columnGroups = [];
      for (var i = 0; i < uIndices.length; i++) {
        if (i < valueIndices.length && i < bonusTypeIndices.length) {
          columnGroups.push({
            uIdx: uIndices[i],
            valueIdx: valueIndices[i],
            bonusTypeIdx: bonusTypeIndices[i],
          });
        }
      }

      var oldVault = {};
      for (var r = 0; r < oldVaultData.length; r++) {
        var row = oldVaultData[r];
        for (var g = 0; g < columnGroups.length; g++) {
          var group = columnGroups[g];
          var u = row[group.uIdx];
          var value = row[group.valueIdx];
          var bonusType = row[group.bonusTypeIdx];
          var key = bonusType || value;
          if (key && isNaN(key)) {
            if (!oldVault.hasOwnProperty(key)) {
              oldVault[key] = [];
            }
            oldVault[key].push(u);
          }
        }
      }
      return { success: true, oldVault: oldVault };
    } catch (error) {
      console.log("Error in getVersion10Vault: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10Vault: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
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
