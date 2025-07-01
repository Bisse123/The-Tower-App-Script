const vault = {
  convertVersionFunctions: {},

  importData: function (sheetType, newVaultSpreadsheetId) {
    function importVaultData(sheetType, newVaultSpreadsheetId) {
      var idType = sheetType + " ID";

      // Get new vault version using SheetsAPI
      var newVaultVersion = SheetsAPI.getValue(newVaultSpreadsheetId, "STATS!A1");
      if (!newVaultVersion) {
      console.log("Error getting new vault version");
      return {
        success: false,
        message: "Error getting new vault version"
      };
      }

      // Get ID Master spreadsheet info
      var idMasterSpreadsheetInfo = shared.findSheetTypeID(
        newVaultSpreadsheetId,
        "IDS"
      );
      var idMasterSpreadsheetId = shared.extractSheetId(
        idMasterSpreadsheetInfo.id
      );
      if (!idMasterSpreadsheetId) {
      console.log("Could not find ID Master spreadsheet");
      return {
        success: false,
        message: "Could not find ID Master spreadsheet"
      };
      }

      var oldVaultSpreadsheetInfo = shared.findSheetTypeID(
        idMasterSpreadsheetId,
        "IDS",
        idType
      );
      var oldVaultSpreadsheetId = shared.extractSheetId(
        oldVaultSpreadsheetInfo.id
      );
      if (!oldVaultSpreadsheetId) {
      console.log("Could not find old vault spreadsheet");
      return {
        success: false,
        message: "Could not find old vault spreadsheet"
      };
      }

      // Get old vault version using SheetsAPI
      var oldVaultVersion = SheetsAPI.getValue(oldVaultSpreadsheetId, "STATS!A1");
      if (!oldVaultVersion) {
      console.log("Error getting old vault version");
      return {
        success: false,
        message: "Error getting old vault version"
      };
      }

      var versionCheck = shared.compareVersions(
        oldVaultVersion,
        newVaultVersion
      );
      if (versionCheck === 0) {
        console.log("Same Version");

        // Get old vault data using Sheets API
        var vaultHarmonyHeaderPattern = ["U", "Value", "Bonus Type"];
        var oldVaultHarmony = getOldVault(
          oldVaultSpreadsheetId,
          "Harmony",
          vaultHarmonyHeaderPattern,
          1
        );

        var vaultPowerHeaderPattern = ["U", "", "Value", "Bonus Type"];
        var oldVaultPower = getOldVault(
          oldVaultSpreadsheetId,
          "Power",
          vaultPowerHeaderPattern,
          1
        );

        var result = updateVault(
          newVaultSpreadsheetId,
          "Harmony",
          oldVaultHarmony,
          vaultHarmonyHeaderPattern,
          1
        );
        if (!result || !result.success) {
          console.log("Error updating Harmony vault: " + result.message);
          return result;
        }

        var result = updateVault(
          newVaultSpreadsheetId,
          "Power",
          oldVaultPower,
          vaultPowerHeaderPattern,
          1
        );
        if (!result || !result.success) {
          console.log("Error updating Power vault: " + result.message);
          return result;
        }
        console.log("Vault data imported successfully");
        return {
          success: true,
          message: "Vault data imported successfully"
        };
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
    }

    function updateVault(
      spreadsheetId,
      sheetName,
      oldVault,
      vaultPattern,
      skipRows
    ) {
      var tierUnlock = ["Tier x2 Unlock", "Tier x3 Unlock"];

      // Get sheet data using SheetsAPI
      var sheetData = SheetsAPI.getDataRange(spreadsheetId, sheetName);
      if (!sheetData) {
        console.log("Error getting sheet data");
        return {
          success: false,
          message: "Error getting sheet data"
        };
      }

      if (sheetData.length < skipRows + 2) return; // Not enough data
      if (sheetData.length < skipRows + 2) {
        return {
          success: false,
          message: "Not enough data in sheet"
        };
      }

      var newVaultHeaders = sheetData[skipRows];
      var newVaultData = sheetData.slice(skipRows + 1);

      var newHeaderIndices = findHeaderIndices(newVaultHeaders, vaultPattern);

      // Iterate through data (excluding headers)
      var newVault = {};
      var updates = []; // Store updates to batch them

      for (var r = 0; r < newVaultData.length; r++) {
        var row = newVaultData[r];
        for (var t = 0; t < newHeaderIndices.length; t++) {
          var uIdx = newHeaderIndices[t][vaultPattern.indexOf("U")];
          var valueIdx = newHeaderIndices[t][vaultPattern.indexOf("Value")];
          var bonusTypeIdx =
            newHeaderIndices[t][vaultPattern.indexOf("Bonus Type")];
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
          if (tierUnlock.includes(key)) {
            // Store update for later batch processing
            var cellAddress =
              shared.columnToLetter(uIdx + 1) + (r + skipRows + 2);
            updates.push({
              range: sheetName + "!" + cellAddress,
              value: u,
            });
          }
          if (!newVault.hasOwnProperty(uIdx)) {
            newVault[uIdx] = [];
          }
          newVault[uIdx].push([u]);
        }
      }

      // Apply single cell updates first
      updates.forEach(function (update) {
        try {
          SheetsAPI.setValue(spreadsheetId, update.range, update.value);
        } catch (error) {
          console.log(
            "Error updating cell " + update.range + ": " + error.toString()
          );
          return {
            success: false,
            message: "Error updating cell: " + error.message,
          };
        }
      });

      // Apply bulk updates for columns
      Object.keys(newVault).forEach(function (colKey) {
        var colIdx = parseInt(colKey, 10);
        var colLetter = shared.columnToLetter(colIdx + 1);
        var values = newVault[colKey];
        var range =
          sheetName + "!" + colLetter + "3:" + colLetter + (2 + values.length);

        try {
          SheetsAPI.setValues(spreadsheetId, range, values);
        } catch (error) {
          console.log("Error updating range " + range + ": " + error.toString());
          return {
            success: false,
            message: "Error updating range: " + error.message,
          };
        }
      });
      console.log("Vault updated successfully");
      return {
        success: true,
        message: "Vault updated successfully"
      };
    }

    function getOldVault(spreadsheetId, sheetName, oldVaultPattern, skipRows) {
      // Get sheet data using SheetsAPI
      var sheetData = SheetsAPI.getDataRange(spreadsheetId, sheetName);
      if (!sheetData) {
        console.log("Error getting old vault data");
        return {};
      }

      if (sheetData.length < skipRows + 1) return {}; // Not enough data

      var oldVaultHeaders = sheetData[skipRows];
      var oldVaultData = sheetData.slice(skipRows + 1);
      var oldVault = {};

      // Find the indices of all "U", "Value", and "Bonus Type" columns by header
      var oldHeaderIndices = findHeaderIndices(
        oldVaultHeaders,
        oldVaultPattern
      );

      // Iterate through data (excluding headers)
      for (var r = 0; r < oldVaultData.length; r++) {
        var row = oldVaultData[r];
        for (var t = 0; t < oldHeaderIndices.length; t++) {
          var uIdx = oldHeaderIndices[t][oldVaultPattern.indexOf("U")];
          var valueIdx = oldHeaderIndices[t][oldVaultPattern.indexOf("Value")];
          var bonusTypeIdx =
            oldHeaderIndices[t][oldVaultPattern.indexOf("Bonus Type")];
          var u = row[uIdx];
          var value = row[valueIdx];
          var bonusType = row[bonusTypeIdx];
          var key = bonusType || value;
          if (key && isNaN(key)) {
            if (!oldVault.hasOwnProperty(key)) {
              oldVault[key] = [];
            }
            oldVault[key].push(u);
          }
        }
      }
      return oldVault;
    }

    function findHeaderIndices(headers, pattern) {
      var indices = [];
      for (var i = 0; i <= headers.length - pattern.length; i++) {
        var match = true;
        for (var j = 0; j < pattern.length; j++) {
          if (headers[i + j] !== pattern[j]) {
            match = false;
            break;
          }
        }
        if (match) {
          indices.push(Array.from({ length: pattern.length }, (_, k) => i + k));
        }
      }
      return indices;
    }
    return importVaultData(sheetType, newVaultSpreadsheetId);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
