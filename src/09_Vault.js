const vault = {
  convertVersionFunctions: {},

  importData: function (versionDifference) {
    function importVaultData(versionDifference) {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!newSpreadsheet) {
          console.log(`New spreadsheet not found`);
          return {
            success: false,
            message: "New spreadsheet not found",
          };
        }
        var newSheetID = newSpreadsheet.spreadsheetId;

        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        if (!oldSpreadsheet) {
          console.log(`Old spreadsheet not found`);
          return {
            success: false,
            message: "Old spreadsheet not found",
          };
        }
        var oldSheetID = oldSpreadsheet.spreadsheetId;
        
        if (versionDifference === 0) {
          // console.log(`Same Version`);

          // Get old vault data using Sheets API
          var vaultHarmonyHeaderPattern = ["U", "Value", "Bonus Type"];
          var oldVaultHarmony = getOldVault(
            oldSheetID,
            "Harmony",
            vaultHarmonyHeaderPattern,
            1
          );

          var vaultPowerHeaderPattern = ["U", "", "Value", "Bonus Type"];
          var oldVaultPower = getOldVault(
            oldSheetID,
            "Power",
            vaultPowerHeaderPattern,
            1
          );

          var result = updateVault(
            newSheetID,
            "Harmony",
            oldVaultHarmony,
            vaultHarmonyHeaderPattern,
            1
          );
          if (!result || !result.success) {
            console.log(`Error updating Harmony vault: ${result.message}`);
            return result;
          }

          var result = updateVault(
            newSheetID,
            "Power",
            oldVaultPower,
            vaultPowerHeaderPattern,
            1
          );
          if (!result || !result.success) {
            console.log(`Error updating Power vault: ${result.message}`);
            return result;
          }
          // console.log(`Vault data imported successfully`);
          return {
            success: true,
            message: `Vault data imported successfully`
          };
        }
        // else {// Else do something to convert old version to new one (Future me problem)
        // }
      } catch (error) {
        console.log(`Error importing vault data: ${error.toString()}`);
        return {
          success: false,
          message: `Error importing vault data: ${error.message}`
        };
      }
    }

    function updateVault(
      newSheetID,
      sheetName,
      oldVault,
      vaultPattern,
      skipRows
    ) {
      var tierUnlock = ["Tier x2 Unlock", "Tier x3 Unlock"];

      // Get sheet data using SheetsAPI
      var sheetData = SheetsAPI.getDataRange(newSheetID, sheetName);
      if (!sheetData) {
        console.log(`Error getting sheet data`);
        return {
          success: false,
          message: `Error getting sheet data`
        };
      }

      if (sheetData.length < skipRows + 2) return;
      if (sheetData.length < skipRows + 2) {
        console.log(`Not enough data in sheet`);
        return {
          success: false,
          message: `Not enough data in sheet`
        };
      }

      var newVaultHeaders = sheetData[skipRows];
      var newVaultData = sheetData.slice(skipRows + 1);

      var newHeaderIndices = findHeaderIndices(newVaultHeaders, vaultPattern);

      var newVault = {};
      var batchUpdate = [];

      for (var r = 0; r < newVaultData.length; r++) {
        var row = newVaultData[r];
        for (var t = 0; t < newHeaderIndices.length; t++) {
          var uIdx = newHeaderIndices[t][vaultPattern.indexOf("U")];
          var valueIdx = newHeaderIndices[t][vaultPattern.indexOf("Value")];
          var bonusTypeIdx = newHeaderIndices[t][vaultPattern.indexOf("Bonus Type")];
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
          // if (tierUnlock.includes(key)) {
          //   // Store update for later batch processing
          //   var cellAddress = shared.columnToLetter(uIdx + 1) + (r + skipRows + 2);
          //   batchUpdate.push({
          //     range: sheetName + "!" + cellAddress,
          //     value: u,
          //   });
          // }
          if (!newVault.hasOwnProperty(uIdx)) {
            newVault[uIdx] = [];
          }
          newVault[uIdx].push([u]);
        }
      }

      // Apply bulk updates for columns (preserve this block)
      Object.keys(newVault).forEach(function (colKey) {
        var colIdx = parseInt(colKey, 10);
        var colLetter = shared.columnToLetter(colIdx + 1);
        var values = newVault[colKey];
        var lastRow = values.length + 2;
        var range = `${sheetName}!${colLetter}3:${colLetter}${lastRow}`;
        batchUpdate.push({
          range: range,
          values: values,
        });
      });

      if (batchUpdate.length !== 0) {
        SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        // console.log(`Vault updated successfully`);
        return {
          success: true,
          message: `Vault updated successfully`
        };
      }
      // console.log(`No updates needed for vault`);
      return {
        success: true,
        message: `No updates needed for vault`
      };
    }

    function getOldVault(oldSheetID, sheetName, oldVaultPattern, skipRows) {
      // Get sheet data using SheetsAPI
      var sheetData = SheetsAPI.getDataRange(oldSheetID, sheetName);
      if (!sheetData) {
        console.log(`Error getting old vault data`);
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
    return importVaultData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
