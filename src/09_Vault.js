const vault = {
  // #region Export Functions
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

  // #endregion
  // #region Import Functions
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
      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Vault",
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

  // #endregion
  // #region Update Functions
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
          bonusTypeIndices.push(col + 1);
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
          var u = null;
          var value = row[valueIdx];
          var bonusType = row[bonusTypeIdx];
          var key = bonusType || value;
          if (oldVault.hasOwnProperty(key)) {
            u = oldVault[key];
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

  // #endregion
  // #region Convert Versions
  version3_1: function () {
    try {
      console.log("Called: vault.version3_1");
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

      var harmonyResult = this.getVersion1_0Vault(harmonyData);
      if (!harmonyResult || !harmonyResult.success) {
        console.log(
          `Error getting harmony vault data: ${harmonyResult.message}`
        );
        return harmonyResult;
      }

      var powerResult = this.getVersion1_0Vault(powerData, true);
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
      console.log("Error in version3_1: " + error.toString());
      return {
        success: false,
        message: "Error in version3_1: " + error.message,
      };
    }
  },

  version1_0: function () {
    try {
      console.log("Called: vault.version1_0");
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

      var harmonyResult = this.getVersion3_1Vault(harmonyData);
      if (!harmonyResult || !harmonyResult.success) {
        console.log(
          `Error getting harmony vault data: ${harmonyResult.message}`
        );
        return harmonyResult;
      }

      var powerResult = this.getVersion3_1Vault(powerData);
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
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Vault
  getVersion3_1Vault: function (oldSheetData) {
    try {
      console.log("Called: vault.getVersion3_1Vault");
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
          bonusTypeIndices.push(col + 1);
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
          var value = row[group.valueIdx];
          var bonusType = row[group.bonusTypeIdx];
          var key = bonusType || value;
          if (!key) {
            continue;
          }
          var uVal = row[group.uIdx];
          var u = (uVal === true || uVal === "TRUE" || uVal === "true" || (typeof uVal === "string" && uVal.includes("x"))) ? uVal : null;
          oldVault[key] = u;
        }
      }

      return { success: true, oldVault: oldVault };
    } catch (error) {
      console.log("Error in getVersion3_1Vault: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion3_1Vault: " + error.message,
      };
    }
  },

  getVersion1_0Vault: function (oldSheetData, addIndexToKey = false) {
    try {
      console.log("Called: vault.getVersion1_0Vault");
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
      var oldVaultValues = {};
      
      for (var r = 0; r < oldVaultData.length; r++) {
        var row = oldVaultData[r];
        for (var g = 0; g < columnGroups.length; g++) {
          var group = columnGroups[g];
          var value = row[group.valueIdx];
          var bonusType = row[group.bonusTypeIdx];
          var key = bonusType || value;
          if (!key) {
            continue;
          }
          var uVal = row[group.uIdx];
          var u = (uVal === true || uVal === "TRUE" || uVal === "true" || (typeof uVal === "string" && uVal.includes("x"))) ? uVal : null;
          if (!oldVaultValues.hasOwnProperty(key)) {
            oldVaultValues[key] = [];
          }
          oldVaultValues[key].push(u);
        }
      }

      Object.keys(oldVaultValues).forEach(function (key) {
        if (oldVaultValues[key].length === 1 && (!addIndexToKey || (typeof key === "string" && key.includes("Tier x")))) {
          oldVault[key] = oldVaultValues[key][0];
          return;
        };

        oldVaultValues[key].forEach(function (u, index) {
          var newKey = key + " " + (index + 1);
          oldVault[newKey] = u;
        });
      });

      return { success: true, oldVault: oldVault };
    } catch (error) {
      console.log("Error in getVersion1_0Vault: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0Vault: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseVaultData: function (data) {
    const harmonyIndices = [
      "Discount Enhancements 1",          // 0
      "Additional Card Slot 1",           // 1
      "Discount Rerolls 1",               // 2
      "Demon Mode Automation",            // 3
      "Free Mission Reroll",              // 4
      "Discount Enhancements 2",          // 5
      "Smart Demon Mode Automation",      // 6
      "Nuke Automation",                  // 7
      "Discount Rerolls 2",               // 8
      "Workshop Respec Discount 1",       // 9
      "Smart Nuke Automation",            // 10
      "Discount Enhancements 3",          // 11
      "Workshop Respec Discount 2",       // 12
      "Additional Card Slot 2",           // 13
      "Discount Rerolls 3",               // 14
      "Workshop Respec Discount 3",       // 15
      "Ad gems Stack x2",                 // 16
      "Discount Enhancements 4",          // 17
      "+5 Workshop Presets",              // 18
      "Ad gems Stack x3",                 // 19
      "Discount Rerolls 4",               // 20
      "Additional Card Slot 3",           // 21
      "Ad gems Stack x5",                 // 22
      "Discount Enhancements 5",          // 23
      "Missile Barrage Automation",       // 24
      "Additional Card Slot 4",           // 25
      "Discount Rerolls 5",               // 26
      "Smart Missile Barrage Automation", // 27
      "Discount Enhancements 6",          // 28
      "Daily Mission - Set Shard Type",   // 29
      "Auto Shatter Rare Modules",        // 30
      "Additional Card Slot 5",           // 31
      "Discount Rerolls 6",               // 32
      "Auto Restart Run",                 // 33
      "Auto Charge Berzerker",            // 34
      "Discount Enhancements 7",          // 35
      "Damage Cap Slider",                // 36
      "Discount Rerolls 7",               // 37
      "Bot Respec Discount 1",            // 38
      "Workshop Orb Adjuster",            // 39
      "Discount Enhancements 8",          // 40
      "Bot Respec Discount 2",            // 41
      "Additional Card Slot 6",           // 42
      "Discount Rerolls 8",               // 43
      "Bot Respec Discount 3",            // 44
      "Bot Cooldown Sliders",             // 45
      "Discount Enhancements 9",          // 46
      "3 bot Presets",                    // 47
    ];

    const powerIndices = [
      "Ultimate Weapon Damage 1", // 0
      "Bot Range 1",              // 1
      "Defense Absolute 1",       // 2
      "Damage / Meter 1",         // 3
      "Cash 1",                   // 4
      "Health Regen 1",           // 5
      "Critical Chance 1",        // 6
      "Coins / Kill 1",           // 7
      "Health 1",                 // 8
      "Damage 1",                 // 9
      "Enemy Attack Skip 1",      // 10
      "Defense % 1",              // 11
      "Super Crit Chance 1",      // 12
      "Enemy Health Skip 1",      // 13
      "Ultimate Weapon Damage 2", // 14
      "Tier x2 Unlock",           // 15
      "Bot Range 2",              // 16
      "Thorn Damage 1",           // 17
      "Rend Armor Mult 1",        // 18
      "Recovery Amount 1",        // 19
      "Knockback Force 1",        // 20
      "Critical Factor 1",        // 21
      "Free Attack Upgrade 1",    // 22
      "Orb Speed 1",              // 23
      "Attack Speed 1",           // 24
      "Free Defense Upgrade 1",   // 25
      "Wall Rebuild 1",           // 26
      "Super Crit Mult 1",        // 27
      "Free Utility Upgrade 1",   // 28
      "Ultimate Weapon Damage 3", // 29
      "Tier x3 Unlock",           // 30
      "Bot Range 3",              // 31
      "Knockback Chance 1",       // 32
      "Rend Armor Chance 1",      // 33
      "Max Recovery 1",           // 34
      "Shockwave Frequency 1",    // 35
      "Rapid Fire Chance 1",      // 36
      "Interest / Wave 1",        // 37
      "Death Defy 1",             // 38
      "Multishot Chance 1",       // 39
      "Cash / Wave 1",            // 40
      "Orbs 1",                   // 41
      "Bounce Shot Chance 1",     // 42
      "Coins / Wave 1",           // 43
      "Ultimate Weapon Damage 4", // 44
      "Bot Range 4",              // 45
    ];

    const harmonyData = data.harmonyNodesUnlocked || [];
    const powerData = data.powerNodesUnlocked || [];
    const powerLevels = data.powerNodesLevel || [];

    var oldVaultHarmony = {};
    var oldVaultPower = {};

    harmonyData.forEach(function (isUnlocked, index) {
      if (isUnlocked && harmonyIndices[index]) {
        oldVaultHarmony[harmonyIndices[index]] = true;
      };
    });

    powerData.forEach(function (isUnlocked, index) {
      if (isUnlocked && powerIndices[index]) {
        var powerLevel = powerIndices[index].includes("Unlock") ? true : "x" + (powerLevels[index] + 1);
        oldVaultPower[powerIndices[index]] = powerLevel;
      }
    });

    return {
      oldVaultHarmony: oldVaultHarmony,
      oldVaultPower: oldVaultPower,
      harmonyIndices: harmonyIndices,
      powerIndices: powerIndices,
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v3.1": this.version3_1.bind(this),
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
