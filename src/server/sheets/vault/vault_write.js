const vaultWriter = {

  /**
   * Builds the batch update that writes Vault into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldVault
   * @param {Object} newVaultData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateVault: function (sheetName, oldVault, newVaultData) {
    try {
      console.log("Called: vaultWriter.updateVault");
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

      var unlockedGroups = oldVault.unlockedGroups;
      var upgradesLevel = oldVault.upgradesLevel;

      var newVaultHeaders = null;
      var headerRowIndex = -1;

      for (var i = 0; i < newVaultData.length; i++) {
        var row = newVaultData[i];

        var uIdx = row.indexOf("U");
        var upgradesIdx = row.indexOf("Upgrades");
        var levelIdx = row.indexOf("Level");
        if (uIdx !== -1 && upgradesIdx !== -1 && levelIdx !== -1) {
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

      var uIndices = [];
      var upgradeIndices = [];
      var levelIndices = [];

      for (var col = 0; col < newVaultHeaders.length; col++) {
        if (newVaultHeaders[col] === "U") {
          uIndices.push(col);
        } else if (newVaultHeaders[col] === "Upgrades") {
          upgradeIndices.push(col);
        } else if (newVaultHeaders[col] === "Level") {
          levelIndices.push(col);
        }
      }

      var newVault = {};
      var batchUpdate = [];

      for (var r = headerRowIndex + 1; r < newVaultData.length; r++) {
        var row = newVaultData[r] || [];

        var breakLoop = true;
        for (var idx = 0; idx < upgradeIndices.length; idx++) {
          const UIdx = uIndices[idx];
          const upgradeIdx = upgradeIndices[idx];
          const levelIdx = levelIndices[idx];
          const sectionName = row[UIdx];
          const upgradeName = row[upgradeIdx];

          if (sectionName || upgradeName) {
            breakLoop = false;
          } else if (!sectionName && !upgradeName) {
            continue;
          }

          if (sectionName && unlockedGroups.indexOf(sectionName) !== -1) {
            const uLetter = sheetRefs.columnToLetter(UIdx + 1);
            batchUpdate.push({
              range: `${sheetName}!${uLetter}${r + 2}`,
              values: [[true]],
            });
          }

          if (!newVault.hasOwnProperty(levelIdx)) {
            newVault[levelIdx] = [];
          }
          var level = null;
          if (upgradeName && upgradesLevel.hasOwnProperty(upgradeName)) {
            level = upgradesLevel[upgradeName];
          }
          newVault[levelIdx].push([level]);
        }

        if (breakLoop) {
          break;
        }
      }

      Object.keys(newVault).forEach(function (colKey) {
        var colIdx = parseInt(colKey, 10);
        var colLetter = sheetRefs.columnToLetter(colIdx + 1);
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
      var errorReport = errors.report("vaultWriter.updateVault", error, {
        sheetName: sheetName,
        oldVault: oldVault,
        newVaultData: newVaultData,
      });
      return errors.fail(errorReport);
    }
  },
};
