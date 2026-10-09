const vaultReader = {

  /**
   * Reads Vault data from a v4.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_0: function (oldSheetID) {
    try {
      console.log("Called: vaultReader.version4_0");
      var oldRanges = ["EXPORT!B4:C"];

      var oldVaultBatchResult = SheetsAPI.batchGetValues(oldSheetID, oldRanges);
      if (!oldVaultBatchResult || oldVaultBatchResult.length === 0) {
        console.log(`Error getting old vault sheet data`);
        return {
          success: false,
          message: `Error getting old vault sheet data`,
        };
      }

      var oldVaultData = oldVaultBatchResult[0].values

      var oldVaultResult = this.getVersion4_0Vault(oldVaultData);

      if (!oldVaultResult || !oldVaultResult.success) {
        console.log(`Error getting old vault data: ${oldVaultResult.message}`);
        return oldVaultResult;
      }

      return {
        success: true,
        message: "Vault processed successfully",
        oldVault: oldVaultResult.oldVault,
      };
    } catch (error) {
      var errorReport = errors.report("vaultReader.version4_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Vault data from a v3.1 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_1: function (oldSheetID) {
    try {
      console.log("Called: vaultReader.version3_1");
      return {
        success: true,
        message: "Vault is from an old version - no data to transfer",
      };

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
      var errorReport = errors.report("vaultReader.version3_1", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Vault data from a v1.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: vaultReader.version1_0");
      return {
        success: true,
        message: "Vault is from an old version - no data to transfer",
      };

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
      var errorReport = errors.report("vaultReader.version1_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Vault from a v4.0 sheet's values.
   * @param {Array<Array<*>>} oldVaultData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0Vault: function (oldVaultData) {
    try {
      console.log("Called: vaultReader.getVersion4_0Vault");
      if (!oldVaultData || oldVaultData.length === 0) {
        console.log(`No sheet data provided for vault`);
        return { success: false, message: `No sheet data provided for vault` };
      }

      var oldVault = {
        unlockedGroups: [],
        upgradesLevel: {},
      };

      for (var i = 0; i < oldVaultData.length; i++) {
        var row = oldVaultData[i] || [];
        var name = String(row[0]).trim();
        var value = row[1];

        if (!name) {
          continue;
        }

        if (name.toLowerCase().endsWith("section")) {
          if ((value === true || value === "TRUE" || value === "true") && !(name.toLowerCase().includes("gameplay") || name.toLowerCase().includes("simple"))) {
            oldVault.unlockedGroups.push(name);
          }
          continue;
        }

        if (value === null || value === undefined || value === "") {
          continue;
        }
        oldVault.upgradesLevel[name] = value;
      }
      return {
        success: true,
        message: "Vault processed successfully",
        oldVault: oldVault
      };
    } catch (error) {
      var errorReport = errors.report("vaultReader.getVersion4_0Vault", error, {
        oldVaultData: oldVaultData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Vault from a v3.1 sheet's values.
   * @param {Array<Array<*>>} oldSheetData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion3_1Vault: function (oldSheetData) {
    try {
      console.log("Called: vaultReader.getVersion3_1Vault");
      return {
        success: true,
        message: "Vault is from an old version - no data to transfer",
      };
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
      var errorReport = errors.report("oldVault.getVersion3_1Vault", error, {
        oldSheetData: oldSheetData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Vault from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldSheetData
   * @param {*} addIndexToKey = false
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0Vault: function (oldSheetData, addIndexToKey = false) {
    try {
      console.log("Called: vaultReader.getVersion1_0Vault");
      return {
        success: true,
        message: "Vault is from an old version - no data to transfer",
      };
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
      var errorReport = errors.report("oldVault.getVersion1_0Vault", error, {
        oldSheetData: oldSheetData,
      });
      return errors.fail(errorReport);
    }
  },
};
