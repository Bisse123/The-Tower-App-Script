const guardians = {
  exportData: function (versionDifference) {
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
      var oldDataResult = getVersionFunction();
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

  importData: function (data) {
    try {
      console.log("Called: guardians.importData");
      var newSpreadsheet = spreadsheets("Guardians newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      // Batch fetch required sheet data
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

      // Only update guardians if key exists
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

  updateGuardianLevels: function (
    sheetName,
    oldGuardians,
    masterSheetData,
    dvtNamedRangesData
  ) {
    try {
      console.log("Called: guardians.updateGuardianLevels");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch", "Summon"];
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
      var endCol = guardianCol + 5;

      var newGuardianDataValues = masterSheetData
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

      if (!newGuardianDataValues || newGuardianDataValues.length === 0) {
        return {
          success: false,
          message: "Could not read guardian data",
        };
      }

      var newGuardianData = newGuardianDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var newGuardianUnlocked = [];
      var newGuardianLevel = [];

      for (var row = 0; row < newGuardianData.length; row++) {
        var rowData = newGuardianData[row];
        var guardianName = rowData[0];
        if (oldGuardians.hasOwnProperty(guardianName)) {
          var oldGuardian = oldGuardians[guardianName];
          newGuardianUnlocked.push([guardianName]);
          newGuardianUnlocked.push([""]);
          newGuardianUnlocked.push([oldGuardian.unlocked]);

          for (var nextRow = row; nextRow < newGuardianData.length; nextRow++) {
            var nextRowData = newGuardianData[nextRow];
            if (nextRow !== row && targetGuardians.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var newGuardianProp = nextRowData[2];
            if (oldGuardian.props.hasOwnProperty(newGuardianProp)) {
              var dvtPropValue = shared.getDVTValue(
                oldGuardian.props[newGuardianProp],
                dvtNamedRangesData[guardianName][newGuardianProp]
              );
              newGuardianLevel.push([dvtPropValue]);
            } else {
              newGuardianLevel.push([nextRowData[4]]);
            }
            if (nextRow == newGuardianData.length - 1) {
              row = nextRow;
            }
          }
        } else {
          newGuardianUnlocked.push([rowData[0]]);
        }
      }

      var batchUpdate = [];
      if (newGuardianUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(guardianCol + 1);
        var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
          newGuardianUnlocked.length + 1
        }`;
        batchUpdate.push({
          range: unlockedRange,
          values: newGuardianUnlocked,
        });
      }

      if (newGuardianLevel.length > 0) {
        var levelCol = shared.columnToLetter(guardianCol + 5);
        var levelRange = `${sheetName}!${levelCol}2:${levelCol}${
          newGuardianLevel.length + 1
        }`;
        batchUpdate.push({
          range: levelRange,
          values: newGuardianLevel,
        });
      }
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

  version22: function () {
    try {
      console.log("Called: guardians.version22");
      var oldSpreadsheet = spreadsheets("Guardians oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }

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

      var guardiansData = this.getVersion22Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      console.log("Error in version22: " + error.toString());
      return {
        success: false,
        message: "Error in version22: " + error.message,
      };
    }
  },

  version21: function () {
    try {
      console.log("Called: guardians.version21");
      var oldSpreadsheet = spreadsheets("Guardians oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }

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

      var guardiansData = this.getVersion21Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      console.log("Error in version21: " + error.toString());
      return {
        success: false,
        message: "Error in version21: " + error.message,
      };
    }
  },

  version10: function () {
    try {
      console.log("Called: guardians.version10");
      var oldSpreadsheet = spreadsheets("Guardians oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old spreadsheet",
        };
      }

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

      var guardiansData = this.getVersion10Guardians(oldGuardianLevelsData);
      return guardiansData;
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardians.getVersion10Guardians");
      var targetGuardians = ["Attack", "Ally", "Steal", "Fetch"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {};
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        // Only proceed if guardianName is in targetGuardians
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            props: {},
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
              guardian.props[key] = value;
            }
          }
          guardianName = guardianName === "Steal" ? "Bounty" : guardianName;
          oldGuardians[guardianName] = guardian;
        }
      }
      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      console.log("Error in getVersion10Guardians: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10Guardians: " + error.message,
      };
    }
  },

  getVersion21Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardians.getVersion21Guardians");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {};
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        // Only proceed if guardianName is in targetGuardians
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            props: {},
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
              guardian.props[key] = value;
            }
          }
          oldGuardians[guardianName] = guardian;
        }
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      console.log("Error in getVersion21Guardians: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion21Guardians: " + error.message,
      };
    }
  },

  getVersion22Guardians: function (oldGuardianLevelsData) {
    try {
      console.log("Called: guardians.getVersion22Guardians");
      var targetGuardians = ["Attack", "Ally", "Bounty", "Fetch", "Summon"];
      var oldGuardianLevels = oldGuardianLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldGuardians = {};
      for (var row = 0; row < oldGuardianLevels.length; row++) {
        var guardianName = oldGuardianLevels[row][0];
        // Only proceed if guardianName is in targetGuardians
        if (guardianName && targetGuardians.includes(guardianName)) {
          var unlocked;
          if (guardianName === "Attack" || guardianName === "Ally") {
            unlocked = null;
          } else {
            unlocked = oldGuardianLevels[row + 2][0];
          }
          var guardian = {
            unlocked: unlocked,
            props: {},
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
              guardian.props[key] = value;
            }
          }
          oldGuardians[guardianName] = guardian;
        }
      }

      return {
        success: true,
        oldGuardians: oldGuardians,
      };
    } catch (error) {
      console.log("Error in getVersion22Guardians: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion22Guardians: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
      "v2.1": this.version21.bind(this),
      "v2.2": this.version22.bind(this),
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
