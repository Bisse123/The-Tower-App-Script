const vault = {
  importData: function (versionDifference) {
    function importVaultData(versionDifference) {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!newSpreadsheet) {
          console.log(`New spreadsheet not found`);
          return {
            success: false,
            message: "New spreadsheet™ not found",
          };
        }
        var newSheetID = newSpreadsheet.spreadsheetId;

        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        if (!oldSpreadsheet) {
          console.log(`Old spreadsheet not found`);
          return {
            success: false,
            message: "Old spreadsheet™ not found",
          };
        }
        var oldSheetID = oldSpreadsheet.spreadsheetId;

        var getVersionFunction = convertVersionFunctions[versionDifference];
        if (!getVersionFunction) {
          console.log(`Unsupported version difference: ${versionDifference}`);
          return {
            success: false,
            message: `Unsupported version difference: ${versionDifference}`,
          };
        }
        var result = getVersionFunction();
        if (!result || !result.success) {
          console.log(`Error processing vault data: ${result.message}`);
          return result;
        }

        var oldVaultHarmony = result.oldVaultHarmony || {};
        var oldVaultPower = result.oldVaultPower || {};
        var vaultHarmonyHeaderPattern = ["U", "Value", "Bonus Type"];
        var vaultPowerHeaderPattern = ["U", "", "Value", "", "Bonus Type"];

        var harmonyResult = updateVault(
          newSheetID,
          "Harmony",
          oldVaultHarmony,
          vaultHarmonyHeaderPattern,
          1
        );
        if (!harmonyResult || !harmonyResult.success) {
          console.log(`Error updating Harmony vault: ${harmonyResult.message}`);
          return harmonyResult;
        }

        var batchUpdate = harmonyResult.batchUpdate || [];

        var powerResult = updateVault(
          newSheetID,
          "Power",
          oldVaultPower,
          vaultPowerHeaderPattern,
          1
        );
        if (!powerResult || !powerResult.success) {
          console.log(`Error updating Power vault: ${powerResult.message}`);
          return powerResult;
        }

        batchUpdate = batchUpdate.concat(powerResult.batchUpdate || []);
        if (batchUpdate.length > 0) {
          // Apply batch updates to the new spreadsheet
          var updateResult = SheetsAPI.batchUpdateValues(
            newSheetID,
            batchUpdate
          );
          if (!updateResult) {
            console.log(`Error applying batch updates to new spreadsheet`);
            return {
              success: false,
              message: "Error applying batch updates to new spreadsheet™",
            };
          }
          // console.log(`Vault data imported successfully`);
          return {
            success: true,
            message: `Vault data imported successfully`,
          };
        }
        // console.log(`No updates needed for vault`);
        return {
          success: true,
          message: `No updates needed for vault`,
        };
      } catch (error) {
        console.log(`Error importing vault data: ${error.toString()}`);
        return {
          success: false,
          message: `Error importing vault data: ${error.message}`,
        };
      }
    }

    function version142() {
      try {
        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        var oldSheetID = oldSpreadsheet.spreadsheetId;
        
        // Get old vault data using Sheets API
        var vaultHarmonyHeaderPattern = ["U", "Value", "Bonus Type"];
        var oldVaultHarmony = getOldVault(
          oldSheetID,
          "Harmony",
          vaultHarmonyHeaderPattern,
          1
        );

        var vaultPowerHeaderPattern = ["U", "", "Value", "", "Bonus Type"];
        var oldVaultPower = getOldVault(
          oldSheetID,
          "Power",
          vaultPowerHeaderPattern,
          1
        );

        return {
          success: true,
          oldVaultHarmony: oldVaultHarmony,
          oldVaultPower: oldVaultPower,
        };
      } catch (error) {
        console.log("Error in version10: " + error.toString());
        return {
          success: false,
          message: "Error in version10: " + error.message,
        };
      }
    }

    function version10() {
      try {
        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        var oldSheetID = oldSpreadsheet.spreadsheetId;
        
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

        return {
          success: true,
          oldVaultHarmony: oldVaultHarmony,
          oldVaultPower: oldVaultPower,
        };
      } catch (error) {
        console.log("Error in version10: " + error.toString());
        return {
          success: false,
          message: "Error in version10: " + error.message,
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
      var sheetBatchResult = SheetsAPI.batchGetValues(newSheetID, [sheetName]);
      if (!sheetBatchResult || sheetBatchResult.length === 0 || !sheetBatchResult[0].values) {
        console.log(`Error getting sheet data`);
        return {
          success: false,
          message: `Error getting sheet data`,
        };
      }
      var sheetData = sheetBatchResult[0].values;

      if (sheetData.length < skipRows + 2) return;
      if (sheetData.length < skipRows + 2) {
        console.log(`Not enough data in sheet`);
        return {
          success: false,
          message: `Not enough data in sheet`,
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
        // console.log(`Vault updated successfully`);
        return {
          success: true,
          message: `Vault updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      // console.log(`No updates needed for vault`);
      return {
        success: true,
        message: `No updates needed for vault`,
      };
    }

    function getOldVault(oldSheetID, sheetName, oldVaultPattern, skipRows) {
      // Get sheet data using SheetsAPI
      var sheetBatchResult = SheetsAPI.batchGetValues(oldSheetID, [sheetName]);
      if (!sheetBatchResult || sheetBatchResult.length === 0 || !sheetBatchResult[0].values) {
        console.log(`Error getting old vault data`);
        return {};
      }
      var sheetData = sheetBatchResult[0].values;

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
    var convertVersionFunctions = {
      "v1.4.2": version142,
      "v1.0": version10,
    };

    return importVaultData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = [
      "v1.4.2",
      "v1.0"
    ];
    
    var sortedThresholds = versionCompatibility.slice().sort(function(a, b) {
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