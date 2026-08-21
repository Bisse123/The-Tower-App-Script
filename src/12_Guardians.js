const guardians = {
  // #region Export Functions
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: guardians.exportData");
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
        message: "Guardians export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting guardians data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data, newSheetID) {
    try {
      console.log("Called: guardians.importData");

      var requiredRanges = ["Master Sheet", "IDS"];
      var dvtIndex = requiredRanges.length;
      var dvtNamedRanges = {
        Attack: {
          Percentage: "DVT_GAR_UG_AT_PER",
          Cooldown: "DVT_GAR_UG_AT_COO",
          Targets: "DVT_GAR_UG_AT_TAR",
        },
        Ally: {
          "Recovery Amount": "DVT_GAR_UG_AL_REC",
          "Max Recovery": "DVT_GAR_UG_AL_MAX",
          Cooldown: "DVT_GAR_UG_AL_COO",
        },
        Bounty: {
          Multiplier: "DVT_GAR_UG_BO_MUL",
          Cooldown: "DVT_GAR_UG_BO_COO",
          Targets: "DVT_GAR_UG_BO_TAR",
        },
        Fetch: {
          Cooldown: "DVT_GAR_UG_FE_COO",
          "Find Chance": "DVT_GAR_UG_FE_FIN",
          "Double Find Chance": "DVT_GAR_UG_FE_DOU",
        },
        Summon: {
          Cooldown: "DVT_GAR_UG_SU_COO",
          Duration: "DVT_GAR_UG_SU_DUR",
          "Cash Bonus": "DVT_GAR_UG_SU_CAS",
        },
        Scout: {
          Cooldown: "DVT_GAR_UG_SC_COO",
          "Range Bonus": "DVT_GAR_UG_SC_RAN",
          Duration: "DVT_GAR_UG_SC_DUR",
        },
      };

      Object.keys(dvtNamedRanges).forEach(function (guardian) {
        Object.keys(dvtNamedRanges[guardian]).forEach(function (prop) {
          requiredRanges.push(dvtNamedRanges[guardian][prop]);
        });
      });

      var batchResult = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log("Error getting guardians sheet data");
        return {
          success: false,
          message: "Error getting guardians sheet data",
        };
      }

      var masterSheetData = batchResult[0].values;
      var idsData = batchResult[1].values;

      var dvtNamedRangesData = {};
      Object.keys(dvtNamedRanges).forEach(function (guardian) {
        dvtNamedRangesData[guardian] = {};
        Object.keys(dvtNamedRanges[guardian]).forEach(function (prop) {
          if (batchResult[dvtIndex]) {
            dvtNamedRangesData[guardian][prop] = batchResult[dvtIndex].values;
          } else {
            dvtNamedRangesData[guardian][prop] = [];
          }
          dvtIndex++;
        });
      });

      var batchUpdate = [];

      if (data.hasOwnProperty("oldGuardians")) {
        var oldGuardians = data.oldGuardians;
        var guardiansResult = this.updateGuardianLevels(
          "Master Sheet",
          oldGuardians,
          masterSheetData,
          dvtNamedRangesData
        );
        if (!guardiansResult || !guardiansResult.success) {
          console.log(`Error updating guardians: ${guardiansResult.message}`);
          return guardiansResult;
        }
        batchUpdate = batchUpdate.concat(guardiansResult.batchUpdate || []);
      }

      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Guardians",
        newSheetID,
        idsData,
        data.idMasterID
      );

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
        message: `Guardians import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing guardians data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Update Functions
  updateGuardianLevels: function (
    sheetName,
    oldGuardians,
    masterSheetData,
    dvtNamedRangesData
  ) {
    try {
      console.log("Called: guardians.updateGuardianLevels");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch", "Summon", "Scout"];
      if (!masterSheetData || masterSheetData.length < 2) {
        return {
          success: false,
          message: "Not enough data in Master Sheet™",
        };
      }

      var headerRow = masterSheetData[0];
      var guardianCol = headerRow.indexOf("Chips") + 1;

      if (guardianCol === 0) {
        console.log(`Guardian Weapon column not found`);
        return {
          success: false,
          message: `Guardian Weapon column not found`,
        };
      }

      var startCol = guardianCol + 1;
      var endCol = startCol;

      var presetColumnMapping = [];
      var firstPresetIndex = startCol + 2;

      if (firstPresetIndex === -1) {
        console.log(`Preset columns not found in Master Sheet`);
        return {
          success: false,
          message: `Preset columns not found in Master Sheet™`,
        };
      }

      var batchUpdate = [];

      var presetSlots = [];
      headerRow.forEach(function (header, index) {
        if (index < firstPresetIndex || !String(header).trim() || presetSlots.length >= oldGuardians.presetNames.length) {
          return;
        }
        presetSlots.push({ header: String(header).trim(), colIndex: index });
      });

      oldGuardians.presetNames.forEach(function (presetName, slot) {
        var presetSlot = presetSlots[slot];
        if (!presetName || !presetSlot) {
          return;
        }
        var colIndex = presetSlot.colIndex;
        endCol = colIndex + 1;
        presetColumnMapping.push({
          presetName: presetName,
          equippedColIndex: colIndex,
          levelColIndex: endCol,
        });
        if (presetSlot.header !== presetName) {
          batchUpdate.push({
            range: `${sheetName}!${shared.columnToLetter(colIndex + 1)}1`,
            values: [[presetName]],
          });
        }
      });

      var newGuardianData = masterSheetData
        .slice(1)
        .map(function (row) {
          return row.slice(startCol - 1, endCol);
        })
        .filter(function (row) {
          return row.some(function (cell) {
            return (
              cell !== null &&
              cell !== undefined &&
              String(cell || "").trim() !== ""
            );
          });
        });

      if (!newGuardianData || newGuardianData.length === 0) {
        return {
          success: false,
          message: "Could not read guardian data",
        };
      }

      var newGuardianUnlocked = [];
      var newGuardianLevels = {};
      var newGuardianEquipped = {};

      presetColumnMapping.forEach(function (presetMap) {
        newGuardianLevels[presetMap.presetName] = [];
        newGuardianEquipped[presetMap.presetName] = [];
      });

      var currentGuardianName = null;
      var currentGuardian = null;
      var guardianRow = -1;

      for (var row = 0; row < newGuardianData.length; row++) {
        var rowData = newGuardianData[row];
        var rowGuardianName = String(rowData[0] || "").trim();

        if (rowGuardianName && targetGuardians.includes(rowGuardianName)) {
          currentGuardianName = rowGuardianName;
          currentGuardian = oldGuardians.data.hasOwnProperty(rowGuardianName)
            ? oldGuardians.data[rowGuardianName]
            : null;
          guardianRow = row;
        }

        if (currentGuardian && row === guardianRow) {
          newGuardianUnlocked.push([currentGuardianName]);
        } else if (currentGuardian && row === guardianRow + 2) {
          newGuardianUnlocked.push([currentGuardian.unlocked]);
        } else {
          newGuardianUnlocked.push([null]);
        }

        var newGuardianProp = rowData[2];
        var dvtGuardianRanges = currentGuardianName
          ? dvtNamedRangesData[currentGuardianName]
          : null;

        presetColumnMapping.forEach(function (presetMap) {
          var presetName = presetMap.presetName;
          var presetData =
            currentGuardian &&
            currentGuardian.presets &&
            currentGuardian.presets[presetName]
              ? currentGuardian.presets[presetName]
              : null;

          if (
            presetData &&
            row === guardianRow &&
            presetData.hasOwnProperty("equipped")
          ) {
            newGuardianEquipped[presetName].push([presetData.equipped]);
          } else {
            newGuardianEquipped[presetName].push([null]);
          }

          if (
            !presetData ||
            !presetData.hasOwnProperty("props") ||
            !newGuardianProp ||
            !presetData.props.hasOwnProperty(newGuardianProp) ||
            !dvtGuardianRanges
          ) {
            newGuardianLevels[presetName].push([null]);
            return;
          }

          var dvtPropValue = shared.getDVTValue(
            presetData.props[newGuardianProp],
            dvtGuardianRanges[newGuardianProp]
          );
          newGuardianLevels[presetName].push([dvtPropValue]);
        });
      }

      if (newGuardianUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(guardianCol + 1);
        batchUpdate.push({
          range: `${sheetName}!${unlockedCol}2:${unlockedCol}${
            newGuardianUnlocked.length + 1
          }`,
          values: newGuardianUnlocked,
        });
      }

      presetColumnMapping.forEach(function (presetMap) {
        var levels = newGuardianLevels[presetMap.presetName];
        if (!levels || levels.length === 0) {
          return;
        }
        var levelCol = shared.columnToLetter(presetMap.levelColIndex + 1);
        batchUpdate.push({
          range: `${sheetName}!${levelCol}2:${levelCol}${levels.length + 1}`,
          values: levels,
        });
      });

      presetColumnMapping.forEach(function (presetMap) {
        var equipped = newGuardianEquipped[presetMap.presetName];
        if (
          !equipped ||
          !equipped.some(function (value) {
            return value[0] !== null && value[0] !== undefined;
          })
        ) {
          return;
        }
        var equippedCol = shared.columnToLetter(presetMap.equippedColIndex + 1);
        batchUpdate.push({
          range: `${sheetName}!${equippedCol}2:${equippedCol}${
            equipped.length + 1
          }`,
          values: equipped,
        });
      });

      if (batchUpdate.length !== 0) {
        return {
          success: true,
          message: `Guardians updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for guardians`,
      };
    } catch (error) {
      console.log("Error in updateGuardianLevels: " + error.toString());
      return {
        success: false,
        message: "Error updating guardian levels: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version3_1: function (oldSheetID) {
    try {
      console.log("Called: guardians.version3_1");

      var guardianLevelsRange = "EXPORT!B4:O";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion3_1Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      console.log("Error in version3_1: " + error.toString());
      return {
        success: false,
        message: "Error in version3_1: " + error.message,
      };
    }
  },

  version2_2: function (oldSheetID) {
    try {
      console.log("Called: guardians.version2_2");

      var guardianLevelsRange = "EXPORT!B5:F";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion2_2Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      console.log("Error in version2_2: " + error.toString());
      return {
        success: false,
        message: "Error in version2_2: " + error.message,
      };
    }
  },

  version2_1: function (oldSheetID) {
    try {
      console.log("Called: guardians.version2_1");

      var guardianLevelsRange = "EXPORT!B5:F";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion2_1Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      console.log("Error in version2_1: " + error.toString());
      return {
        success: false,
        message: "Error in version2_1: " + error.message,
      };
    }
  },

  version1_0: function (oldSheetID) {
    try {
      console.log("Called: guardians.version1_0");

      var guardianLevelsRange = "EXPORT!B5:F";
      var guardianBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        guardianLevelsRange,
      ]);
      if (
        !guardianBatchResult ||
        guardianBatchResult.length === 0 ||
        !guardianBatchResult[0].values
      ) {
        console.log(`Could not read guardian levels data`);
        return {
          success: false,
          message: `Could not read guardian levels data`,
        };
      }
      var oldGuardianLevelsData = guardianBatchResult[0].values;

      var guardiansData = this.getVersion1_0Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Guardians
  getVersion3_1Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardians.getVersion3_1Guardians");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch", "Summon", "Scout"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      if (!oldGuardianLevels || oldGuardianLevels.length === 0) {
        return {
          success: false,
          message: "No guardian levels data found",
        };
      }

      var oldGuardiansHeaderRow = oldGuardianLevels[0] || [];

      // That range is laid out as guardian name, unlocked, chip name, spacer,
      // then one (equipped, level) pair per preset - so the presets always
      // start at the fifth column no matter what they were renamed to.
      var firstPresetIndex = 4;

      if (oldGuardiansHeaderRow.length <= firstPresetIndex) {
        console.log(`Could not find the preset header row in guardian data`);
        return {
          success: false,
          message: "Could not find the preset header row in guardian data",
        };
      }

      var oldGuardiansPresetNames = [];
      var presetColumnMapping = [];

      for (
        var colIdx = firstPresetIndex;
        colIdx < oldGuardiansHeaderRow.length;
        colIdx++
      ) {
        var presetName = String(oldGuardiansHeaderRow[colIdx] || "").trim();
        if (!presetName) {
          continue;
        }

        oldGuardiansPresetNames.push(presetName);
        presetColumnMapping.push({
          presetName: presetName,
          equippedColIndex: colIdx,
          levelColIndex: colIdx + 1,
        });
      }

      var oldGuardians = {
        presetNames: shared.resolvePresetOrder(
          oldGuardiansPresetNames,
          shared.templatePresetNames,
        ).order,
        data: {},
      };

      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianRowData = oldGuardianLevels[row];
        var guardianName = String(guardianRowData[0] || "").trim();
        if (!guardianName || !targetGuardians.includes(guardianName)) {
          continue;
        }

        var unlocked;
        if (guardianName === "Attack" || guardianName === "Ally") {
          unlocked = null;
        } else {
          unlocked = oldGuardianLevels[row + 2]
            ? oldGuardianLevels[row + 2][0]
            : null;
        }

        var guardian = {
          unlocked: unlocked,
          presets: {},
        };

        presetColumnMapping.forEach(function (presetMap) {
          guardian.presets[presetMap.presetName] = {
            props: {},
            equipped: guardianRowData[presetMap.equippedColIndex],
          };
        });

        for (var nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
          var nextRowData = oldGuardianLevels[nextRow];
          if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
            row = nextRow - 1;
            break;
          }
          var key = String(nextRowData[2] || "").trim();
          if (!key) {
            continue;
          }
          
          // var defaultLevelColIndex = presetColumnMapping[0].levelColIndex;
          // var defaultLevelValue = nextRowData[defaultLevelColIndex];
          presetColumnMapping.forEach(function (presetMap) {
            var levelValue = nextRowData[presetMap.levelColIndex];
            // if (
            //   presetMap.levelColIndex !== defaultLevelColIndex &&
            //   levelValue === defaultLevelValue
            // ) {
            //   levelValue = null;
            // }
            guardian.presets[presetMap.presetName].props[key] = levelValue;
          });
        }

        oldGuardians.data[guardianName] = guardian;
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      console.log("Error in getVersion3_1Guardians: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion3_1Guardians: " + error.message,
      };
    }
  },

  getVersion2_2Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardians.getVersion2_2Guardians");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch", "Summon", "Scout"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {
        presetNames: ["Farming"],
        data: {},
      };
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            presets: {
              Farming: {
                props: {},
              },
            },
          };

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              guardian.presets.Farming.props[key] = value;
            }
          }
          oldGuardians.data[guardianName] = guardian;
        }
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      console.log("Error in getVersion2_2Guardians: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion2_2Guardians: " + error.message,
      };
    }
  },

  getVersion2_1Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardians.getVersion2_1Guardians");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {
        presetNames: ["Farming"],
        data: {},
      };
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            presets: {
              Farming: {
                props: {},
              },
            },
          };

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              guardian.presets.Farming.props[key] = value;
            }
          }
          oldGuardians.data[guardianName] = guardian;
        }
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      console.log("Error in getVersion2_1Guardians: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion2_1Guardians: " + error.message,
      };
    }
  },

  getVersion1_0Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardians.getVersion1_0Guardians");
      var targetGuardians = ["Attack", "Ally", "Steal", "Fetch"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {
        presetNames: ["Farming"],
        data: {},
      };
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            presets: {
              Farming: {
                props: {},
              },
            },
          };

          for (nextRow = row; nextRow < oldGuardianLevels.length; nextRow++) {
            var nextRowData = oldGuardianLevels[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              value = (value - 1).toString().padStart(2, "0");
              guardian.presets.Farming.props[key] = value;
            }
          }
          guardianName = guardianName === "Steal" ? "Bounty" : guardianName;
          oldGuardians.data[guardianName] = guardian;
        }
      }
      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      console.log("Error in getVersion1_0Guardians: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0Guardians: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseGuardiansData: function (data) {
    const targetGuardians = {
      "Bounty":  { upgrades: ["Multiplier", "Cooldown", "Targets"] },
      "Catch":   { upgrades: [null, null, null] },
      "Attack":  { upgrades: ["Percentage", "Cooldown", "Targets"], alwaysUnlocked: true },
      "Scare":   { upgrades: [null, null, null] },
      "Rush":    { upgrades: [null, null, null] },
      "Ally":    { upgrades: ["Recovery Amount", "Max Recovery", "Cooldown"], alwaysUnlocked: true },
      "Fetch":   { upgrades: ["Cooldown", "Find Chance", "Double Find Chance"] },
      "Summon":  { upgrades: ["Cooldown", "Duration", "Cash Bonus"] },
      "Scout":   { upgrades: ["Cooldown", "Range Bonus", "Duration"] },
    };

    const guardianSlotData = data.guardianChipSlot || [];
    const guardianUnlockedData = data.guardianChipUnlocked || [];
    const guardianLevelData = data.guardianChipLevel || [];

    var oldGuardians = {
      presetNames: ["Farming"],
      data: {},
    };

    Object.keys(targetGuardians).forEach(function (guardianName, i) {
      const { upgrades, alwaysUnlocked } = targetGuardians[guardianName];
      if (upgrades.every(attr => attr === null)) {
        return;
      }
      var chipLevels = {};
      upgrades.forEach(function (attr, j) {
        if (attr === null) return;
        const idx = i * upgrades.length + j;
        const level = guardianLevelData[idx];
        chipLevels[attr] = level ? String(level).padStart(2, "0") : "00";
      });
      oldGuardians.data[guardianName] = {
        unlocked: alwaysUnlocked ? null : (guardianUnlockedData[i] || false),
        presets: {
          Farming: {
            props: chipLevels,
            equipped: guardianSlotData.includes(i),
          },
        },
      };
    });
    
    const guardianOrder = Object.keys(targetGuardians).filter(function (guardianName) {
      return targetGuardians[guardianName].upgrades.some(function (attr) {return attr});
    });

    return {
      oldGuardians: oldGuardians,
      targetGuardians: targetGuardians,
      guardianOrder: guardianOrder,
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v2.1": this.version2_1.bind(this),
      "v2.2": this.version2_2.bind(this),
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
