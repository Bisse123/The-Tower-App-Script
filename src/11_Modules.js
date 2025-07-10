const modules = {
  importData: function (versionDifference) {
    function importModulesData(versionDifference) {
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
          console.log(`Error processing modules data: ${result.message}`);
          return result;
        }

        var targetModuleTypes = result.targetModuleTypes || [];
        var oldModulesInventory = result.oldModulesInventory || {};
        var oldModulesPresets = result.oldModulesPresets || {};

        var inventoryResult = updateModulesInventory(
          targetModuleTypes,
          newSheetID,
          "Modules Inventory",
          oldModulesInventory
        );
        if (!inventoryResult || !inventoryResult.success) {
          return {
            success: false,
            message: inventoryResult.message,
          };
        }

        var batchUpdate = inventoryResult.batchUpdate || [];

        var presetsResult = updateModulesPresets(
          targetModuleTypes,
          newSheetID,
          "Modules Presets",
          oldModulesPresets
        );

        if (!presetsResult || !presetsResult.success) {
          return {
            success: false,
            message: presetsResult.message,
          };
        }

        batchUpdate = batchUpdate.concat(presetsResult.batchUpdate || []);

        if (batchUpdate.length > 0) {
          SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
          // Update the sheet with the new data
          // console.log(`Modules data imported successfully`);
          return {
            success: true,
            message: `Modules data imported successfully`,
          };
        }
        // No updates needed
        // console.log(`No updates needed for modules data`);
        return {
          success: true,
          message: `No updates needed for modules data`,
        };
      } catch (error) {
        console.log(`Error importing modules data: ${error.toString()}`);
        return {
          success: false,
          message: `Error importing modules data: ${error.message}`,
        };
      }
    }

    function version40() {
      try {
        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        var oldSheetID = oldSpreadsheet.spreadsheetId;
        
        var targetModuleTypes = ["cannon", "armor", "generator", "core"];

        var oldModulesInventoryValues = SheetsAPI.getDataRange(
          oldSheetID,
          "Modules Inventory"
        );

        var oldModulesInventory = getOldModulesInventory(
          targetModuleTypes,
          oldModulesInventoryValues
        );

        // Get old modules presets data using Sheets API
        var oldModulesPresetsValues = SheetsAPI.getDataRange(
          oldSheetID,
          "Modules Presets"
        );

        var oldModulesPresets = getOldModulesPresets(
          targetModuleTypes,
          oldModulesPresetsValues
        );

        return {
          success: true,
          targetModuleTypes: targetModuleTypes,
          oldModulesInventory: oldModulesInventory,
          oldModulesPresets: oldModulesPresets,
        };
      } catch (error) {
        console.log("Error in version40: " + error.toString());
        return {
          success: false,
          message: "Error in version40: " + error.message,
        };
      }
    }

    function updateModulesPresets(
      targetModuleTypes,
      newSheetID,
      sheetName,
      oldModulesPresets
    ) {
      // Get sheet data using Sheets API
      var newModulePresetsValues;
      try {
        newModulePresetsValues = SheetsAPI.getDataRange(newSheetID, sheetName);
      } catch (error) {
        console.log("Error getting modules presets data: " + error.toString());
        return {
          success: false,
          message: "Error getting modules presets data: " + error.message,
        };
      }

      var newModuleTypeIndex = findModuleTypesRowIndex(
        targetModuleTypes,
        newModulePresetsValues
      );

      var batchUpdate = [];

      targetModuleTypes.forEach(function (moduleType) {
        if (oldModulesPresets.hasOwnProperty(moduleType)) {
          var rowIdx = newModuleTypeIndex[moduleType] + 1;
          if (typeof rowIdx === "undefined") return;
          var row = newModulePresetsValues[rowIdx];
          for (var col = 0; col < row.length; col++) {
            if (String(row[col]).trim() === "Module Name") {
              var presetName = String(
                newModulePresetsValues[rowIdx - 1][col]
              ).trim();
              if (
                presetName &&
                oldModulesPresets[moduleType].hasOwnProperty(presetName)
              ) {
                var cellAddress = shared.columnToLetter(col + 1) + (rowIdx + 2);
                batchUpdate.push({
                  range: sheetName + "!" + cellAddress,
                  values: [[oldModulesPresets[moduleType][presetName]]],
                });
              }
            }
          }
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
    }

    function getOldModulesPresets(targetModuleTypes, oldModulesPresetsValues) {
      var oldModuleTypeIndex = findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesPresetsValues
      );
      var oldModules = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        if (typeof rowIdx === "undefined") return;
        oldModules[moduleType] = {};
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
              oldModules[moduleType][presetName] = moduleName;
            }
          }
        }
      });
      return oldModules;
    }

    function updateModulesInventory(
      targetModuleTypes,
      newSheetID,
      sheetName,
      oldModulesInventory
    ) {
      // Get sheet data using Sheets API
      var newModuleInventoryValues;
      try {
        newModuleInventoryValues = SheetsAPI.getDataRange(
          newSheetID,
          sheetName
        );
      } catch (error) {
        console.log(
          "Error getting modules inventory data: " + error.toString()
        );
        return {
          success: false,
          message: "Error getting modules inventory data: " + error.message,
        };
      }

      var newModuleTypeIndex = findModuleTypesRowIndex(
        targetModuleTypes,
        newModuleInventoryValues
      );

      var singleCellUpdates = [];
      var rangeUpdates = [];

      targetModuleTypes.forEach(function (moduleType) {
        if (oldModulesInventory.hasOwnProperty(moduleType)) {
          var rowIdx = newModuleTypeIndex[moduleType];
          if (typeof rowIdx === "undefined") return;
          var row = newModuleInventoryValues[rowIdx];
          var maxLevel = oldModulesInventory[moduleType]["Highest Level"] || 0;
          var highestLevelCol =
            newModuleInventoryValues[rowIdx + 1].indexOf("Highest Level");

          for (var col = 1; col < row.length; col++) {
            var cellValue = String(row[col]);
            if (
              cellValue.trim() !== "" &&
              oldModulesInventory[moduleType].hasOwnProperty(cellValue)
            ) {
              maxLevel = Math.max(
                maxLevel,
                oldModulesInventory[moduleType][cellValue].level
              );

              // Update rarity
              var rarityCell = shared.columnToLetter(col + 1) + (rowIdx + 3);
              singleCellUpdates.push({
                range: sheetName + "!" + rarityCell,
                value: oldModulesInventory[moduleType][cellValue].rarity,
              });

              // var levelCell = shared.columnToLetter(col + 2) + (rowIdx + 3);
              // singleCellUpdates.push({
              //   range: sheetName + "!" + levelCell,
              //   value: oldModulesInventory[moduleType][cellValue].level,
              // });

              // Update substats if available
              var substats =
                oldModulesInventory[moduleType][cellValue].substats;
              if (substats && substats.length > 0) {
                var numRows = substats.length;
                var numCols = substats[0].length;
                var startCell = shared.columnToLetter(col + 1) + (rowIdx + 5);
                var endCell =
                  shared.columnToLetter(col + numCols) +
                  (rowIdx + 5 + numRows - 1);
                rangeUpdates.push({
                  range: sheetName + "!" + startCell + ":" + endCell,
                  values: substats,
                });
              }
            }
          }

          // Update highest level
          if (highestLevelCol !== -1) {
            var highestLevelCell =
              shared.columnToLetter(highestLevelCol + 1) + (rowIdx + 3);
            singleCellUpdates.push({
              range: sheetName + "!" + highestLevelCell,
              value: maxLevel,
            });
          }
        }
      });

      var batchUpdate = [];
      singleCellUpdates.forEach(function (update) {
        try {
          batchUpdate.push({
            range: update.range,
            values: [[update.value]],
          });
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

      // Apply range updates
      rangeUpdates.forEach(function (update) {
        try {
          batchUpdate.push({
            range: update.range,
            values: update.values,
          });
        } catch (error) {
          console.log(
            "Error updating range " + update.range + ": " + error.toString()
          );
          return {
            success: false,
            message: "Error updating range: " + error.message,
          };
        }
      });

      // console.log(`Modules inventory updated successfully`);
      return {
        success: true,
        message: `Modules inventory updated successfully`,
        batchUpdate: batchUpdate,
      };
    }

    function getOldModulesInventory(
      targetModuleTypes,
      oldModulesInventoryValues
    ) {
      var oldModuleTypeIndex = findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesInventoryValues
      );
      var oldModules = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType];
        if (typeof rowIdx === "undefined") return;
        oldModules[moduleType] = {};
        var row = oldModulesInventoryValues[rowIdx];
        var highestLevelCol =
          oldModulesInventoryValues[rowIdx + 1].indexOf("Highest Level");
        if (highestLevelCol !== -1) {
          oldModules[moduleType]["Highest Level"] =
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
              var moduleLevel =
                oldModulesInventoryValues[rowIdx + 2][col + 1] != null
                  ? String(
                      oldModulesInventoryValues[rowIdx + 2][col + 1]
                    ).trim()
                  : "";
              oldModules[moduleType][moduleName] = {
                rarity: moduleRarity,
                level: moduleLevel,
                substats: [],
              };
              for (
                var substat = rowIdx + 4;
                substat < oldModulesInventoryValues.length;
                substat++
              ) {
                var substatRow = oldModulesInventoryValues[substat];
                var substatColVal =
                  substatRow && substatRow[col] != null
                    ? String(substatRow[col]).trim()
                    : "";
                var substatCol1Val =
                  substatRow && substatRow[col + 1] != null
                    ? String(substatRow[col + 1]).trim()
                    : "";
                if (substatColVal === "" && substatCol1Val === "") {
                  break;
                }
                var substats = [
                  substatRow ? substatRow[col] : "",
                  substatRow ? substatRow[col + 1] : "",
                ];
                oldModules[moduleType][moduleName]["substats"].push(substats);
              }
            }
          }
        }
      });
      return oldModules;
    }

    function findModuleTypesRowIndex(targetModuleTypes, moduleRange) {
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
    }

    var convertVersionFunctions = {
      "v4.0": version40
    };

    return importModulesData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = [
      "v4.0"
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