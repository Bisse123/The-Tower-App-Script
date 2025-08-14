const modules = {
  exportData: function (versionDifference) {
    try {
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
      console.log(oldDataResult.oldModulesPresets);
      return {
        success: true,
        message: "Modules export completed successfully",
        data: {
          oldModulesInventory: oldDataResult.oldModulesInventory || {},
          oldModulesPresets: oldDataResult.oldModulesPresets || {},
          oldModulesObtained: oldDataResult.oldModulesObtained || {}
        }
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting modules data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      // Use sheet type-based naming for parallel execution support
      var newSpreadsheet = spreadsheets("Modules newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      var oldModulesInventory = data.oldModulesInventory || {};
      var oldModulesPresets = data.oldModulesPresets || {};
      var oldModulesObtained = data.oldModulesObtained || {};

      // Batch fetch all required sheet data
      var requiredRanges = [
        "Modules Inventory",
        "Modules Presets",
        "Modules Tracker",
        "IDS"
      ];
      var batchResult = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResult || batchResult.length < 3) {
        console.log("Error getting modules sheet data");
        return {
          success: false,
          message: "Error getting modules sheet data",
        };
      }

      var newModuleInventoryValues =
        batchResult[0] && batchResult[0].values
          ? batchResult[0].values
          : null;
      var newModulePresetsValues =
        batchResult[1] && batchResult[1].values
          ? batchResult[1].values
          : null;
      var newModulesObtainedValues =
        batchResult[2] && batchResult[2].values
          ? batchResult[2].values
          : null;
      var idsData = batchResult[3].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(newSheetID, "IDS", "IDS Master's", idsData);
      if (!newSheetInfo || !newSheetInfo.importStatus || !newSheetInfo.importStatus.range) {
        console.log(`Could not find import status range in IDS sheet`);
        return {
          success: false,
          message: "Could not find import status range in IDS sheet",
        };
      }

      var inventoryResult = this.updateModulesInventory(
        "Modules Inventory",
        oldModulesInventory,
        newModuleInventoryValues
      );
      if (!inventoryResult || !inventoryResult.success) {
        return {
          success: false,
          message: inventoryResult.message,
        };
      }

      var batchUpdate = inventoryResult.batchUpdate || [];

      var presetsResult = this.updateModulesPresets(
        "Modules Presets",
        oldModulesPresets,
        newModulePresetsValues
      );

      if (!presetsResult || !presetsResult.success) {
        return {
          success: false,
          message: presetsResult.message,
        };
      }

      batchUpdate = batchUpdate.concat(presetsResult.batchUpdate || []);

      var obtainedResult = this.updateModulesObtained(
        "Modules Tracker",
        oldModulesObtained,
        newModulesObtainedValues
      );

      if (!obtainedResult || !obtainedResult.success) {
        return {
          success: false,
          message: obtainedResult.message,
        };
      }

      batchUpdate = batchUpdate.concat(obtainedResult.batchUpdate || []);

      // Add import status update to batch
      batchUpdate.push({
        range: newSheetInfo.importStatus.range,
        values: [["✅"]],
      });
      
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

      return {
        success: true,
        message: `Modules data imported successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing modules data: ${error.message}`,
      };
    }
  },

  updateModulesPresets: function (
    sheetName,
    oldModulesPresets,
    newModulePresetsValues
  ) {
    if (!newModulePresetsValues) {
      console.log(`Could not read modules presets data`);
      return {
        success: false,
        message: "Could not read modules presets data",
      };
    }

    var targetModuleTypes = ["cannon", "armor", "generator", "core"];
    var newModuleTypeIndex = this.findModuleTypesRowIndex(
      targetModuleTypes,
      newModulePresetsValues
    );

    var batchUpdate = [];

    targetModuleTypes.forEach(function (moduleType) {
      if (oldModulesPresets.hasOwnProperty(moduleType)) {
        var rowIdx = newModuleTypeIndex[moduleType] + 1;
        if (typeof rowIdx === "undefined") return;
        var row = newModulePresetsValues[rowIdx];
        Object.keys(oldModulesPresets[moduleType]).forEach(function (presetName) {
          var presetCol = row.indexOf(presetName);
          if (presetCol !== -1) {
            if (presetName === "Assist Slot") {
              var rarityRange = `${sheetName}!${shared.columnToLetter(presetCol + 2)}${rowIdx + 2}:${shared.columnToLetter(presetCol + 2)}${rowIdx + 2}`;
              var rarityValues = [
                [oldModulesPresets[moduleType][presetName].rarity || ""],
              ];
              var multiSubRange = `${sheetName}!${shared.columnToLetter(presetCol + 3)}${rowIdx + 3}:${shared.columnToLetter(presetCol + 3)}${rowIdx + 4}`;
              var multiSubValues = [
                [oldModulesPresets[moduleType][presetName].multiplier || ""],
                [oldModulesPresets[moduleType][presetName].substat || ""],
              ];
              batchUpdate.push({
                range: rarityRange,
                values: rarityValues,
              });
              batchUpdate.push({
                range: multiSubRange,
                values: multiSubValues,
              });
            } else {
              var range = `${sheetName}!${shared.columnToLetter(presetCol + 2)}${rowIdx + 3}:${shared.columnToLetter(presetCol + 2)}${rowIdx + 4}`;
              var values = [
                [oldModulesPresets[moduleType][presetName].primary || ""],
                [oldModulesPresets[moduleType][presetName].secondary || ""],
              ];
              batchUpdate.push({
                range: range,
                values: values,
              });
            }
          }
        });
      }
    });

    if (batchUpdate.length > 0) {
      // console.log(`Modules presets updated successfully`);
      return {
        success: true,
        message: `Modules presets updated successfully`,
        batchUpdate: batchUpdate,
      };
    }
    // console.log(`No updates needed for modules presets`);
    return {
      success: true,
      message: `No updates needed for modules presets`,
    };
  },

  updateModulesInventory: function (
    sheetName,
    oldModulesInventory,
    newModuleInventoryValues
  ) {
    if (!newModuleInventoryValues) {
      console.log("Could not read modules inventory data");
      return {
        success: false,
        message: "Could not read modules inventory data",
      };
    }

    var targetModuleTypes = ["cannon", "armor", "generator", "core"];
    var newModuleTypeIndex = this.findModuleTypesRowIndex(
      targetModuleTypes,
      newModuleInventoryValues
    );

    var batchUpdate = [];

    targetModuleTypes.forEach(function (moduleType) {
      if (oldModulesInventory.hasOwnProperty(moduleType)) {
        var rowIdx = newModuleTypeIndex[moduleType];
        if (typeof rowIdx === "undefined") return;
        var row = newModuleInventoryValues[rowIdx];
        for (var col = 1; col < row.length; col++) {
          var cellValue = String(row[col]);
          if (
            cellValue.trim() === "Any Other" &&
            !oldModulesInventory[moduleType].hasOwnProperty(cellValue)
          ) {
            for (var spare in oldModulesInventory[moduleType]) {
              if (spare.includes("Spare")) {
                batchUpdate.push({
                  range: `${sheetName}!${shared.columnToLetter(col + 1)}${
                    rowIdx + 1
                  }`,
                  values: [[spare]],
                });
                cellValue = spare;
                break;
              }
            }
          }
          if (
            cellValue.trim() !== "" &&
            oldModulesInventory[moduleType].hasOwnProperty(cellValue)
          ) {
            var rarityCell = shared.columnToLetter(col + 1) + (rowIdx + 3);
            batchUpdate.push({
              range: `${sheetName}!${rarityCell}`,
              values: [[oldModulesInventory[moduleType][cellValue].rarity]],
            });

            var substats =
              oldModulesInventory[moduleType][cellValue].substats;
            if (substats && substats.length > 0) {
              var numRows = substats.length;
              var numCols = substats[0].length;
              var startCell = shared.columnToLetter(col + 1) + (rowIdx + 5);
              var endCell =
                shared.columnToLetter(col + numCols) +
                (rowIdx + 5 + numRows - 1);
              batchUpdate.push({
                range: `${sheetName}!${startCell}:${endCell}`,
                values: substats,
              });
            }
          }
        }

        var maxLevel = oldModulesInventory[moduleType]["Highest Level"] || 0;
        var assistLevel = oldModulesInventory[moduleType]["Assist Level"] || 0;
        var highestLevelCol =
          newModuleInventoryValues[rowIdx + 1].indexOf("Highest Level");
        var assistLevelCol =
          newModuleInventoryValues[rowIdx + 4].indexOf("Assist Level");

        // Update highest level
        if (highestLevelCol !== -1) {
          var highestLevelCell =
            shared.columnToLetter(highestLevelCol + 1) + (rowIdx + 3);
          batchUpdate.push({
            range: `${sheetName}!${highestLevelCell}`,
            values: [[maxLevel]],
          });
        }
        if (assistLevelCol !== -1) {
          var assistLevelCell =
            shared.columnToLetter(assistLevelCol + 1) + (rowIdx + 5);
          batchUpdate.push({
            range: `${sheetName}!${assistLevelCell}`,
            values: [[assistLevel]],
          });
        }
      }
    });

    // console.log(`Modules inventory updated successfully`);
    return {
      success: true,
      message: `Modules inventory updated successfully`,
      batchUpdate: batchUpdate,
    };
  },

  updateModulesObtained: function (
    sheetName,
    oldModulesObtained,
    newModulesObtainedValues
  ) {
    if (!newModulesObtainedValues) {
      console.log("Could not read modules obtained data");
      return {
        success: false,
        message: "Could not read modules obtained data",
      };
    }

    var targetModuleTypes = ["cannon", "armor", "generator", "core"];
    var batchUpdate = [];
    for (var row = 0; row < newModulesObtainedValues.length; row++) {
      for (var col = 0; col < newModulesObtainedValues[row].length; col++) {
        var cellValue = String(newModulesObtainedValues[row][col]).trim();
        if (cellValue) {
          var moduleType = targetModuleTypes.find(function (type) {
            return cellValue.toLowerCase().includes(`${type} summary`);
          });
          moduleType = moduleType ? `${moduleType} summary` : null;
          if (moduleType && oldModulesObtained.hasOwnProperty(moduleType)) {
            var valuesArray = [];
            var startRow = row + 3;
            for (
              var modIdx = row + 2;
              modIdx < newModulesObtainedValues.length;
              modIdx++
            ) {
              var moduleName = String(
                newModulesObtainedValues[modIdx][col]
              ).trim();
              if (!moduleName || moduleName === "Total") {
                break;
              } else if (
                oldModulesObtained[moduleType].hasOwnProperty(moduleName)
              ) {
                valuesArray.push([
                  oldModulesObtained[moduleType][moduleName],
                ]);
              }
            }
            if (valuesArray.length > 0) {
              var startCell = shared.columnToLetter(col + 2) + startRow;
              var endCell =
                shared.columnToLetter(col + 2) +
                (startRow + valuesArray.length - 1);
              batchUpdate.push({
                range: `${sheetName}!${startCell}:${endCell}`,
                values: valuesArray,
              });
            }
          }
        }
      }
    }
    return {
      success: true,
      message: `Modules Obtained updated successfully`,
      batchUpdate: batchUpdate,
    };
  },

  findModuleTypesRowIndex: function (targetModuleTypes, moduleRange) {
    var moduleTypeIndex = {};
    var moduleFound = {};
    targetModuleTypes.forEach(function (moduleType) {
      moduleFound[moduleType] = false;
    });

    for (var i = 0; i < moduleRange.length; i++) {
      var cellValue = String(moduleRange[i][0]).toLowerCase();
      targetModuleTypes.forEach(function (moduleType) {
        if (
          !moduleFound[moduleType] &&
          cellValue &&
          cellValue.indexOf(moduleType) !== -1
        ) {
          moduleTypeIndex[moduleType] = i;
          moduleFound[moduleType] = true;
        }
      });
      // If all terms are found, we can break early
      if (Object.values(moduleFound).every(Boolean)) {
        break;
      }
    }

    return moduleTypeIndex;
  },
  version50: function () {
    try {
      var oldSpreadsheet = spreadsheets("Modules oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var ranges = ["Modules Inventory", "Modules Presets", "Modules Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length < 3) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesObtainedValues = batchResult[2].values;

      var inventoryData = this.getVersion50ModulesInventory(oldModulesInventoryValues);
      var presetsData = this.getVersion50ModulesPresets(oldModulesPresetsValues);
      var obtainedData = this.getVersion47ModulesObtained(oldModulesObtainedValues);

      var success = inventoryData.success && presetsData.success && obtainedData.success;
      
      return {
        success: success,
        message: success ? "Modules data retrieved successfully" : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesObtained: obtainedData.oldModulesObtained || {}
      };
    } catch (error) {
      console.log("Error in version50: " + error.toString());
      return {
        success: false,
        message: "Error in version50: " + error.message,
      };
    }
  },

  version47: function () {
    try {
      var oldSpreadsheet = spreadsheets("Modules oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var ranges = ["Modules Inventory", "Modules Presets", "Modules Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length < 3) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesObtainedValues = batchResult[2].values;

      var inventoryData = this.getVersion40ModulesInventory(oldModulesInventoryValues);
      var presetsData = this.getVersion40ModulesPresets(oldModulesPresetsValues);
      var obtainedData = this.getVersion47ModulesObtained(oldModulesObtainedValues);

      var success = inventoryData.success && presetsData.success && obtainedData.success;
      
      return {
        success: success,
        message: success ? "Modules data retrieved successfully" : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesObtained: obtainedData.oldModulesObtained || {}
      };
    } catch (error) {
      console.log("Error in version47: " + error.toString());
      return {
        success: false,
        message: "Error in version47: " + error.message,
      };
    }
  },

  version40: function () {
    try {
      var oldSpreadsheet = spreadsheets("Modules oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var ranges = ["Modules Inventory", "Modules Presets", "Mods Obtained"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length < 3) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesObtainedValues = batchResult[2].values;

      var inventoryData = this.getVersion40ModulesInventory(oldModulesInventoryValues);
      var presetsData = this.getVersion40ModulesPresets(oldModulesPresetsValues);
      var obtainedData = this.getVersion40ModulesObtained(oldModulesObtainedValues);

      var success = inventoryData.success && presetsData.success && obtainedData.success;
      
      return {
        success: success,
        message: success ? "Modules data retrieved successfully" : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesObtained: obtainedData.oldModulesObtained || {}
      };
    } catch (error) {
      console.log("Error in version40: " + error.toString());
      return {
        success: false,
        message: "Error in version40: " + error.message,
      };
    }
  },

  getVersion40ModulesInventory: function (oldModulesInventoryValues) {
    try {
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = this.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesInventoryValues
      );

      var oldModulesInventory = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType];
        if (typeof rowIdx === "undefined") return;
        oldModulesInventory[moduleType] = {};
        var row = oldModulesInventoryValues[rowIdx];
        var highestLevelCol =
          oldModulesInventoryValues[rowIdx + 1].indexOf("Highest Level");
        if (highestLevelCol !== -1) {
          oldModulesInventory[moduleType]["Highest Level"] =
            oldModulesInventoryValues[rowIdx + 2][highestLevelCol];
        }
        for (var col = 1; col < row.length; col++) {
          var cellValue = row[col] != null ? String(row[col]) : "";
          if (cellValue.trim() !== "") {
            var moduleName = cellValue;
            if (moduleName) {
              var removedRarity = ["Common", "Rare", "Rare+"];
              var moduleRarity =
                oldModulesInventoryValues[rowIdx + 2][col] != null
                  ? String(oldModulesInventoryValues[rowIdx + 2][col]).trim()
                  : "";
              if (removedRarity.includes(moduleRarity)) {
                moduleRarity = "Epic";
              }
              oldModulesInventory[moduleType][moduleName] = {
                rarity: moduleRarity,
                substats: [],
              };
              for (
                var substat = rowIdx + 4;
                substat < oldModulesInventoryValues.length;
                substat++
              ) {
                var substatRow = oldModulesInventoryValues[substat];
                var substatName =
                  substatRow && substatRow[col] != null
                    ? String(substatRow[col]).trim()
                    : "";
                var substatRarity =
                  substatRow && substatRow[col + 1] != null
                    ? String(substatRow[col + 1]).trim()
                    : "";
                if (substatName === "" && substatRarity === "") {
                  break;
                }
                var substatData = [
                  substatRow ? substatRow[col] : "",
                  substatRow ? substatRow[col + 1] : "",
                ];
                oldModulesInventory[moduleType][moduleName]["substats"].push(
                  substatData
                );
              }
            }
          }
        }
      });

      return {
        success: true,
        oldModulesInventory: oldModulesInventory,
      };
    } catch (error) {
      console.log("Error in getVersion40ModulesInventory: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion40ModulesInventory: " + error.message,
      };
    }
  },

  getVersion40ModulesPresets: function (oldModulesPresetsValues) {
    try {
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = this.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesPresetsValues
      );

      var oldModulesPresets = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        if (typeof rowIdx === "undefined") return;
        oldModulesPresets[moduleType] = {};
        var row = oldModulesPresetsValues[rowIdx];
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Module Name") {
            var presetName = String(
              oldModulesPresetsValues[rowIdx - 1][col]
            ).trim();
            var moduleName = String(
              oldModulesPresetsValues[rowIdx + 1][col]
            ).trim();
            if (presetName && moduleName) {
              oldModulesPresets[moduleType][presetName] = {"primary": moduleName, "secondary": ""};
            }
          }
        }
      });

      return {
        success: true,
        oldModulesPresets: oldModulesPresets,
      };
    } catch (error) {
      console.log("Error in getVersion40ModulesPresets: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion40ModulesPresets: " + error.message,
      };
    }
  },

  getVersion40ModulesObtained: function (oldModulesObtainedValues) {
    try {
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModulesObtained = {};
      
      for (var row = 0; row < oldModulesObtainedValues.length; row++) {
        for (var col = 0; col < oldModulesObtainedValues[row].length; col++) {
          var cellValue = String(oldModulesObtainedValues[row][col]).trim();
          if (cellValue) {
            var moduleType = targetModuleTypes.find(function (type) {
              return cellValue.toLowerCase().includes(type);
            });
            moduleType = moduleType ? `${moduleType} summary` : null;
            if (moduleType) {
              if (!oldModulesObtained[moduleType]) {
                oldModulesObtained[moduleType] = {};
              }
              for (
                var modIdx = row + 2;
                modIdx < oldModulesObtainedValues.length;
                modIdx++
              ) {
                var moduleName = String(
                  oldModulesObtainedValues[modIdx][col]
                ).trim();
                if (!moduleName || moduleName === "Total") {
                  break;
                } else {
                  oldModulesObtained[moduleType][moduleName] =
                    oldModulesObtainedValues[modIdx][col + 1];
                }
              }
              if (
                Object.keys(oldModulesObtained).length ===
                targetModuleTypes.length
              ) {
                break;
              }
            }
          }
        }
        if (
          Object.keys(oldModulesObtained).length === targetModuleTypes.length
        ) {
          break;
        }
      }

      return {
        success: true,
        oldModulesObtained: oldModulesObtained,
      };
    } catch (error) {
      console.log("Error in getVersion40ModulesObtained: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion40ModulesObtained: " + error.message,
      };
    }
  },

  getVersion47ModulesObtained: function (oldModulesObtainedValues) {
    try {
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModulesObtained = {};
      
      for (var row = 0; row < oldModulesObtainedValues.length; row++) {
        for (var col = 0; col < oldModulesObtainedValues[row].length; col++) {
          var cellValue = String(oldModulesObtainedValues[row][col]).trim();
          if (cellValue) {
            var moduleType = targetModuleTypes.find(function (type) {
              return cellValue.toLowerCase().includes(`${type} summary`);
            });
            moduleType = moduleType ? `${moduleType} summary` : null;
            if (moduleType) {
              if (!oldModulesObtained[moduleType]) {
                oldModulesObtained[moduleType] = {};
              }
              for (
                var modIdx = row + 2;
                modIdx < oldModulesObtainedValues.length;
                modIdx++
              ) {
                var moduleName = String(
                  oldModulesObtainedValues[modIdx][col]
                ).trim();
                if (!moduleName || moduleName === "Total") {
                  break;
                } else {
                  oldModulesObtained[moduleType][moduleName] =
                    oldModulesObtainedValues[modIdx][col + 1];
                }
              }
              if (
                Object.keys(oldModulesObtained).length ===
                targetModuleTypes.length
              ) {
                break;
              }
            }
          }
        }
        if (
          Object.keys(oldModulesObtained).length === targetModuleTypes.length
        ) {
          break;
        }
      }

      return {
        success: true,
        oldModulesObtained: oldModulesObtained,
      };
    } catch (error) {
      console.log("Error in getVersion47ModulesObtained: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion47ModulesObtained: " + error.message,
      };
    }
  },

  getVersion50ModulesInventory: function (oldModulesInventoryValues) {
    try {
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = this.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesInventoryValues
      );

      var oldModulesInventory = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType];
        if (typeof rowIdx === "undefined") return;
        oldModulesInventory[moduleType] = {};
        var row = oldModulesInventoryValues[rowIdx];
        var highestLevelCol =
          oldModulesInventoryValues[rowIdx + 1].indexOf("Highest Level");
        if (highestLevelCol !== -1) {
          oldModulesInventory[moduleType]["Highest Level"] =
            oldModulesInventoryValues[rowIdx + 2][highestLevelCol];
        }
        var assistLevelCol =
          oldModulesInventoryValues[rowIdx + 4].indexOf("Assist Level");
        if (assistLevelCol !== -1) {
          oldModulesInventory[moduleType]["Assist Level"] =
            oldModulesInventoryValues[rowIdx + 5][assistLevelCol];
        }
        for (var col = 1; col < row.length; col++) {
          var cellValue = row[col] != null ? String(row[col]) : "";
          if (cellValue.trim() !== "") {
            var moduleName = cellValue;
            if (moduleName) {
              var moduleRarity =
                oldModulesInventoryValues[rowIdx + 2][col] != null
                  ? String(oldModulesInventoryValues[rowIdx + 2][col]).trim()
                  : "";
              oldModulesInventory[moduleType][moduleName] = {
                rarity: moduleRarity,
                substats: [],
              };
              for (
                var substat = rowIdx + 4;
                substat < oldModulesInventoryValues.length;
                substat++
              ) {
                var substatRow = oldModulesInventoryValues[substat];
                var substatName =
                  substatRow && substatRow[col] != null
                    ? String(substatRow[col]).trim()
                    : "";
                var substatRarity =
                  substatRow && substatRow[col + 1] != null
                    ? String(substatRow[col + 1]).trim()
                    : "";
                if (substatName === "" && substatRarity === "") {
                  break;
                }
                var substatData = [
                  substatName || "",
                  substatRarity || "",
                ];
                oldModulesInventory[moduleType][moduleName]["substats"].push(
                  substatData
                );
              }
            }
          }
        }
      });

      return {
        success: true,
        oldModulesInventory: oldModulesInventory,
      };
    } catch (error) {
      console.log("Error in getVersion50ModulesInventory: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion50ModulesInventory: " + error.message,
      };
    }
  },

  getVersion50ModulesPresets: function (oldModulesPresetsValues) {
    try {
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = this.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesPresetsValues
      );

      var oldModulesPresets = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        if (typeof rowIdx === "undefined") return;
        oldModulesPresets[moduleType] = {};
        var row = oldModulesPresetsValues[rowIdx + 2];
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Primary Slot") {
            var presetName = oldModulesPresetsValues[rowIdx][col] ? String(oldModulesPresetsValues[rowIdx][col]).trim() : "";
            var primaryName = oldModulesPresetsValues[rowIdx + 2][col + 1] ? String(oldModulesPresetsValues[rowIdx + 2][col + 1]).trim() : "";
            var secondaryName = oldModulesPresetsValues[rowIdx + 3][col + 1] ? String(oldModulesPresetsValues[rowIdx + 3][col + 1]).trim() : "";
            oldModulesPresets[moduleType][presetName] = {"primary": primaryName, "secondary": secondaryName};
          }
        }
        var assistSlotCol = oldModulesPresetsValues[rowIdx].indexOf("Assist Slot");
        if (assistSlotCol !== -1) {
          var assistRarity = String(
            oldModulesPresetsValues[rowIdx + 1][assistSlotCol + 1])
            .trim();
          var assistMultiplier = String(
            oldModulesPresetsValues[rowIdx + 2][assistSlotCol + 2])
            .trim();
          var assistSubstat = String(
            oldModulesPresetsValues[rowIdx + 3][assistSlotCol + 2])
            .trim();
          oldModulesPresets[moduleType]["Assist Slot"] = {
            "rarity": assistRarity,
            "multiplier": assistMultiplier,
            "substat": assistSubstat
          };
        }
      });

      return {
        success: true,
        oldModulesPresets: oldModulesPresets,
      };
    } catch (error) {
      console.log("Error in getVersion50ModulesPresets: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion50ModulesPresets: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v4.0": this.version40.bind(this),
      "v4.7": this.version47.bind(this),
      "v5.0": this.version50.bind(this),
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
