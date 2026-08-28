const vault = {
  // #region Export Functions
  exportData: function (versionDifference, oldSheetID) {
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

      var oldDataResult = getVersionFunction(oldSheetID);
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
  importData: function (data, newSheetID) {
    try {
      console.log("Called: vault.importData");


      var requiredRanges = ["IDS", "Master Sheet"];

      var batchResults = SheetsAPI.batchGetValues(
        newSheetID,
        requiredRanges
      );
      if (!batchResults || batchResults.length === 0) {
        console.log("Error getting vault sheet data");
        return {
          success: false,
          message: "Error getting vault sheet data",
        };
      }

      var idsData = batchResults[0].values;
      var masterSheetData = batchResults[1].values;

      var batchUpdate = [];
      console.log("data", JSON.stringify(data, null, 2));
      if (data.hasOwnProperty("oldVault")) {
        var updateResult = this.updateVault(
          "Master Sheet",
          data.oldVault,
          masterSheetData
        );
        if (!updateResult || !updateResult.success) {
          console.log(`Error updating vault: ${updateResult.message}`);
          return updateResult;
        }
        batchUpdate = batchUpdate.concat(updateResult.batchUpdate || []);
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
  updateVault: function (sheetName, oldVault, newVaultData) {
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

      // Find all occurrences of each column type
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
            const uLetter = shared.columnToLetter(UIdx + 1);
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
  version4_0: function (oldSheetID) {
    try {
      console.log("Called: vault.version4_0");
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
      console.log("Error in version4_0: " + error.toString());
      return {
        success: false,
        message: "Error in version4_0: " + error.message,
      };
    }
  },

  version3_1: function (oldSheetID) {
    try {
      console.log("Called: vault.version3_1");
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
      console.log("Error in version3_1: " + error.toString());
      return {
        success: false,
        message: "Error in version3_1: " + error.message,
      };
    }
  },

  version1_0: function (oldSheetID) {
    try {
      console.log("Called: vault.version1_0");
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
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Vault
  getVersion4_0Vault: function (oldVaultData) {
    try {
      console.log("Called: vault.getVersion4_0Vault");
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
      console.log("Error in getVersion4_0Vault: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion4_0Vault: " + error.message,
      };
    }
  },

  getVersion3_1Vault: function (oldSheetData) {
    try {
      console.log("Called: vault.getVersion3_1Vault");
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
    const vaultGroupsByIndex = {
      // Harmony
      100: {
        group: "Harmony",
        section: "Gameplay Section",
        alwaysUnlocked: true,
        upgrades: {
          1000: "Auto Restart Run",
          1010: "Ad Gem Stack",
          1020: "Damage Cap Slider",
          1030: "Global Presets",
        },
      },
      110: {
        group: "Harmony",
        section: "Cards Section",
        upgrades: {
          1200: "Additional Card Slot",
          1210: "Charge Berzerker",
          1220: "Demon Mode Automation",
          1230: "Smart Demon Mode Automation",
          1240: "Nuke Automation",
          1250: "Smart Nuke Automation",
          1260: "Death Ray Toggle",
          1270: "Spawn Accelerator Toggle",
          1280: "Life Saving Ordering",
          1290: "Energy Shield Restriction",
          // 1300: "Bastion Automation",
          // 1310: "Smart Bastion Automation",
        },
      },
      120: {
        group: "Harmony",
        section: "UW Section",
        upgrades: {
          1400: "Missile Barrage Automation",
          1410: "Smart Missile Barrage Automation",
          1420: "Black Hole Size Slider",
          1430: "Swamp Range Slider",
        },
      },
      130: {
        group: "Harmony",
        section: "Modules Section",
        upgrades: {
          1600: "Auto Shatter Rare Modules",
          1610: "Daily Mission Shard Type",
          1620: "Free Mission Reroll",
          1630: "Module Reroll Discount",
          1640: "Module Presets",
        },
      },
      140: {
        group: "Harmony",
        section: "Bots Section",
        upgrades: {
          1800: "Bot Respec Discount",
          1805: "Bot Presets",
          1810: "Bot Range",
          1820: "Bot Cooldown Sliders",
        },
      },
      150: {
        group: "Harmony",
        section: "Workshop Section",
        upgrades: {
          2000: "Workshop Respec Discount",
          2005: "Workshop Presets",
          2010: "Enhancements Discount",
          2020: "Workshop Orb Adjuster",
        },
      },
      160: {
        group: "Harmony",
        section: "Guardian Section",
        upgrades: {
          2200: "Guardian Respec Discount",
          2205: "Guardian Presets",
        },
      },
      // Power
      1000: {
        group: "Power",
        section: "Attack Section",
        upgrades: {
          1: "Damage",
          10: "Attack Speed",
          20: "Critical Chance",
          30: "Crit Factor",
          40: "Damage / Meter",
          50: "Multishot Chance",
          60: "Rapid Fire Chance",
          70: "Bounce Shot Chance",
          80: "Super Crit Chance",
          90: "Super Crit Mult",
          100: "Rend Armor Chance",
          110: "Rend Armor Mult",
        },
      },
      1010: {
        group: "Power",
        section: "Defense Section",
        upgrades: {
          200: "Health",
          210: "Health Regen",
          220: "Defense %",
          230: "Defense Absolute",
          240: "Thorn Damage",
          250: "Knockback Chance",
          260: "Knockback Force",
          270: "Orb Speed",
          280: "Orbs",
          290: "Shockwave Size",
          300: "Shockwave Frequency",
          310: "Death Defy",
          320: "Wall Health",
          330: "Wall Rebuild",
        },
      },
      1020: {
        group: "Power",
        section: "Utility Section",
        upgrades: {
          400: "Cash Bonus",
          410: "Cash / Wave",
          420: "Coins / Kill",
          430: "Coins / Wave",
          440: "Free Attack Upgrade",
          450: "Free Defense Upgrade",
          460: "Free Utility Upgrade",
          470: "Interest / Wave",
          480: "Recovery Amount",
          490: "Max Recovery",
          500: "Enemy Attack Skip",
          510: "Enemy Health Skip",
        },
      },
      1030: {
        group: "Power",
        section: "Ultimate Weapons Section",
        upgrades: {
          600: "Ultimate Damage",
          610: "Chain Lightning Damage",
          620: "Smart Missile Damage",
          630: "Death Wave Damage",
          640: "Inner Land Mine Damage",
          650: "Golden Tower Bonus",
          660: "Swamp Damage",
          670: "Spotlight Bonus",
          680: "Black Hole Size",
        },
      },
      // Enemy
      2000: {
        group: "Enemy",
        section: "Simple Section",
        alwaysUnlocked: true,
        upgrades: {
          3000: "Basic Attack",
          3010: "Basic Health",
          3020: "Basic Coin Bonus",
          3100: "Fast Attack",
          3110: "Fast Health",
          3120: "Fast Speed",
          3200: "Tank Attack",
          3210: "Tank Health",
          3220: "Tank Mass",
        },
      },
      2010: {
        group: "Enemy",
        section: "Advanced Enemies Section",
        upgrades: {
          4000: "Boss Attack",
          4010: "Boss Health",
          4020: "Boss Common Drop Chance",
          4100: "Ranged Attack",
          4110: "Ranged Health",
          4120: "Ranged Attack Range",
          4200: "Protector Attack",
          4210: "Protector Health",
          4220: "Protector Radius",
        },
      },
      2020: {
        group: "Enemy",
        section: "Elites Enemies Section",
        upgrades: {
          5000: "Vampire Attack",
          5010: "Vampire Health",
          5020: "Vampire Attack Speed",
          5100: "Scatter Attack",
          5110: "Scatter Health",
          5120: "Scatter Attack Speed",
          5200: "Ray Attack",
          5210: "Ray Health",
          5220: "Ray Attack Speed",
        },
      },
      2030: {
        group: "Enemy",
        section: "Fleets Enemies Section",
        upgrades: {
          6000: "Saboteur Attack Speed",
          6010: "Saboteur Health",
          6020: "Saboteur Miss Chance",
          6100: "Commander Pulse Effect",
          6110: "Commander Health",
          6120: "Commander Pulse Speed",
          6200: "Overcharge Attack",
          6210: "Overcharge Health",
          6220: "Overcharge Attack Exponent",
        },
      },
    };

    const vault = data.vault || {};

    const unlockedGroupsName = "UnlockedGroups";
    const upgradesLevelName = "UpgradesLevel";

    var oldVault = {
      unlockedGroups: [],
      upgradesLevel: {},
    };

    var upgradeNamesByIndex = {};
    Object.keys(vaultGroupsByIndex).forEach((groupKey) => {
      const groupUpgrades = vaultGroupsByIndex[groupKey].upgrades || {};
      Object.keys(groupUpgrades).forEach((upgradeKey) => {
        upgradeNamesByIndex[upgradeKey] = groupUpgrades[upgradeKey];
      });
    });

    Object.keys(vault).forEach((key) => {
      if (key.includes(unlockedGroupsName)) {
        const unlockedGroups = vault[key];
        const unlockedElements = unlockedGroups.hasOwnProperty("Elements")
          ? unlockedGroups.Elements
          : [];
        unlockedElements.forEach((element) => {
          const groupInfo = vaultGroupsByIndex[element];
          if (groupInfo && !groupInfo.alwaysUnlocked) {
            oldVault.unlockedGroups.push(groupInfo.section);
          }
        });
      } else if (key.includes(upgradesLevelName)) {
        const upgradesLevel = vault[key];
        Object.keys(upgradesLevel).forEach((upgradeKey) => {
          const upgradeName = upgradeNamesByIndex[upgradeKey];
          if (upgradeName) {
            oldVault.upgradesLevel[upgradeName] = upgradesLevel[upgradeKey];
          }
        });
      }
    });

    var vaultIndices = [];
    var alwaysUnlocked = [];
    var columnsByName = {};

    Object.keys(vaultGroupsByIndex)
      .map(Number)
      .sort((a, b) => a - b)
      .forEach((groupIndex) => {
        const groupInfo = vaultGroupsByIndex[groupIndex];
        if (groupInfo.alwaysUnlocked) {
          alwaysUnlocked.push(groupInfo.section);
        }
        if (!columnsByName.hasOwnProperty(groupInfo.group)) {
          columnsByName[groupInfo.group] = {
            group: groupInfo.group,
            sections: [],
          };
          vaultIndices.push(columnsByName[groupInfo.group]);
        }

        const groupUpgrades = groupInfo.upgrades || {};
        const upgradeNames = Object.keys(groupUpgrades)
          .map(Number)
          .sort((a, b) => a - b)
          .map((upgradeIndex) => groupUpgrades[upgradeIndex]);

        columnsByName[groupInfo.group].sections.push({
          section: groupInfo.section,
          upgrades: upgradeNames,
        });
      });

    return {
      oldVault: oldVault,
      vaultIndices: vaultIndices,
      alwaysUnlocked: alwaysUnlocked,
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v3.1": this.version3_1.bind(this),
      "v4.0": this.version4_0.bind(this),
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
