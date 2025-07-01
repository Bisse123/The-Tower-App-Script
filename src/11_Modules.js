const modules = {
  convertVersionFunctions: {
    3: this.convertVersion3,
  },

  importData: function (sheetType, newModulesSpreadsheetId) {
    function importModulesData(sheetType, newModulesSpreadsheetId) {
      var idType = sheetType + " ID";
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];

      // Get new modules version using Sheets API
      var newModulesVersion;
      try {
        newModulesVersion = SheetsAPI.getValue(newModulesSpreadsheetId, "EXPORT!A1");
      } catch (error) {
        console.log("Error getting new modules version: " + error.toString());
        return {
          success: false,
          message: "Error getting new modules version: " + error.message
        };
      }

      // Get ID Master spreadsheet info
      var idMasterSpreadsheetInfo = shared.findSheetTypeID(
        newModulesSpreadsheetId,
        "IDS"
      );
      var idMasterSpreadsheetId = shared.extractSheetId(
        idMasterSpreadsheetInfo.id
      );
      if (!idMasterSpreadsheetId) {
      console.log("Could not find ID Master spreadsheet");
      return {
        success: false,
        message: "Could not find ID Master spreadsheet"
      };
      }

      var oldModulesSpreadsheetInfo = shared.findSheetTypeID(
        idMasterSpreadsheetId,
        "IDS",
        idType
      );
      var oldModulesSpreadsheetId = shared.extractSheetId(
        oldModulesSpreadsheetInfo.id
      );
      if (!oldModulesSpreadsheetId) {
      console.log("Could not find old modules spreadsheet");
      return {
        success: false,
        message: "Could not find old modules spreadsheet"
      };
      }

      // Get old modules version using Sheets API
      var oldModulesVersion;
      try {
        oldModulesVersion = SheetsAPI.getValue(oldModulesSpreadsheetId, "EXPORT!A1");
      } catch (error) {
        console.log("Error getting old modules version: " + error.toString());
        return {
          success: false,
          message: "Error getting old modules version: " + error.message
        };
      }

      var versionCheck = shared.compareVersions(
        oldModulesVersion,
        newModulesVersion
      );
      if (versionCheck === 0) {
        console.log("Same Version");

        // Get old modules inventory data using Sheets API
        var oldModulesInventoryValues;
        try {
          oldModulesInventoryValues = SheetsAPI.getDataRange(
            oldModulesSpreadsheetId,
            "Modules Inventory"
          ) || [];
        } catch (error) {
          console.log(
            "Error getting old modules inventory: " + error.toString()
          );
          return {
            success: false,
            message: "Error getting old modules inventory: " + error.message
          };
        }

        var oldModulesInventory = getOldModulesInventory(
          targetModuleTypes,
          oldModulesInventoryValues
        );
        var result = updateModulesInventory(
          targetModuleTypes,
          newModulesSpreadsheetId,
          "Modules Inventory",
          oldModulesInventory
        );
        if (!result || !result.success) {
          return {
            success: false,
            message: result.message
          };
        }

        // Get old modules presets data using Sheets API
        var oldModulesPresetsValues;
        try {
          oldModulesPresetsValues = SheetsAPI.getDataRange(
            oldModulesSpreadsheetId,
            "Modules Presets"
          ) || [];
        } catch (error) {
          console.log("Error getting old modules presets: " + error.toString());
          return {
            success: false,
            message: "Error getting old modules presets: " + error.message
          };
        }

        var oldModulesPresets = getOldModulesPresets(
          targetModuleTypes,
          oldModulesPresetsValues
        );
        var result = updateModulesPresets(
          targetModuleTypes,
          newModulesSpreadsheetId,
          "Modules Presets",
          oldModulesPresets
        );
        if (!result || !result.success) {
          return {
            success: false,
            message: result.message
          };
        }
        console.log("Modules data imported successfully");
        return {
          success: true,
          message: "Modules data imported successfully"
        };
      }
      // Else do something to convert old version to new one (Future me problem)
      else {
        var { oldModulesInventory, oldModulesPresets } =
          this.convertVersionFunctions[versionCheck](oldModulesSpreadsheetId);
        var result = updateModulesInventory(
          targetModuleTypes,
          newModulesSpreadsheetId,
          "Modules Inventory",
          oldModulesInventory
        );
        if (!result || !result.success) {
          return {
            success: false,
            message: result.message
          };
        }

        var result = updateModulesPresets(
          targetModuleTypes,
          newModulesSpreadsheetId,
          "Modules Presets",
          oldModulesPresets
        );
        if (!result || !result.success) {
          return {
            success: false,
            message: result.message
          };
        }
        console.log("Modules data imported successfully");
        return {
          success: true,
          message: "Modules data imported successfully"
        };
      }
      // Check version to figure out which convert function to use
      // Potentially do something where if new version is v4 and old is v2
      // you can do convert v2 -> v3 and then convert v3 -> v4
      // Means more calculations but less rewriting of convert functions when new version are released
    }

    function updateModulesPresets(
      targetModuleTypes,
      spreadsheetId,
      sheetName,
      oldModulesPresets
    ) {
      // Get sheet data using Sheets API
      var newModulePresetsValues;
      try {
        newModulePresetsValues = SheetsAPI.getDataRange(spreadsheetId, sheetName);
      } catch (error) {
        console.log("Error getting modules presets data: " + error.toString());
        return {
          success: false,
          message: "Error getting modules presets data: " + error.message
        };
      }

      var newModuleTypeIndex = findModuleTypesRowIndex(
        targetModuleTypes,
        newModulePresetsValues
      );

      var updates = []; // Collect updates for batch processing

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
                updates.push({
                  range: sheetName + "!" + cellAddress,
                  value: oldModulesPresets[moduleType][presetName],
                });
              }
            }
          }
        }
      });

      var batchUpdate = [];
      updates.forEach(function (update) {
        try {
          batchUpdate.push({
            range: update.range,
            values: [[update.value]]
          });
        } catch (error) {
          console.log(
            "Error updating preset " + update.range + ": " + error.toString()
          );
          return {
            success: false,
            message: "Error updating preset: " + error.message
          };
        }
      });
      SheetsAPI.batchUpdateValues(spreadsheetId, batchUpdate);
      console.log("Modules presets updated successfully");
      return {
        success: true,
        message: "Modules presets updated successfully"
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
      spreadsheetId,
      sheetName,
      oldModulesInventory
    ) {
      // Get sheet data using Sheets API
      var newModuleInventoryValues;
      try {
        newModuleInventoryValues = SheetsAPI.getDataRange(spreadsheetId, sheetName);
      } catch (error) {
        console.log("Error getting modules inventory data: " + error.toString());
        return {
          success: false,
          message: "Error getting modules inventory data: " + error.message
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

              // Update level (check if there's no formula first)
              var levelCell = shared.columnToLetter(col + 2) + (rowIdx + 3);
              // Note: We can't easily check for formulas with Sheets API, so we'll just update
              singleCellUpdates.push({
                range: sheetName + "!" + levelCell,
                value: oldModulesInventory[moduleType][cellValue].level,
              });

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

      var batchUpdate = []
      singleCellUpdates.forEach(function (update) {
        try {
          batchUpdate.push({
            range: update.range,
            values: [[update.value]]
          });
        } catch (error) {
          console.log(
            "Error updating cell " + update.range + ": " + error.toString()
          );
          return {
            success: false,
            message: "Error updating cell: " + error.message
          };
        }
      });

      // Apply range updates
      rangeUpdates.forEach(function (update) {
        try {
          batchUpdate.push({
            range: update.range,
            values: update.values
          });
        } catch (error) {
          console.log(
            "Error updating range " + update.range + ": " + error.toString()
          );
          return {
            success: false,
            message: "Error updating range: " + error.message
          };
        }
      });
      try {
        SheetsAPI.batchUpdateValues(spreadsheetId, batchUpdate);
      } catch (error) {
        console.log("Error updating modules inventory: " + error.toString());
        return {
          success: false,
          message: "Error updating modules inventory: " + error.message
        };
      }
      console.log("Modules inventory updated successfully");
      return {
        success: true,
        message: "Modules inventory updated successfully"
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
              var moduleRarity = oldModulesInventoryValues[rowIdx + 2][col] != null ? String(oldModulesInventoryValues[rowIdx + 2][col]).trim() : "";
              if (removedRarity.includes(moduleRarity)) {
                moduleRarity = "Epic";
              }
              var moduleLevel = oldModulesInventoryValues[rowIdx + 2][col + 1] != null ? String(oldModulesInventoryValues[rowIdx + 2][col + 1]).trim() : "";
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
                var substatColVal = substatRow && substatRow[col] != null ? String(substatRow[col]).trim() : "";
                var substatCol1Val = substatRow && substatRow[col + 1] != null ? String(substatRow[col + 1]).trim() : "";
                if (
                  substatColVal === "" &&
                  substatCol1Val === ""
                ) {
                  break;
                }
                var substats = [substatRow ? substatRow[col] : "", substatRow ? substatRow[col + 1] : ""];
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
    return importModulesData(sheetType, newModulesSpreadsheetId);
  },

  convertVersion3: function (oldModulesSpreadsheetId) {
    var targetModuleTypes = ["cannon", "armor", "generator", "core"];

    // Check if Modules Presets sheet exists
    if (!SheetsAPI.hasSheet(oldModulesSpreadsheetId, "Modules Presets")) {
      console.log("Modules Presets sheet not found");
      return {
        success: false,
        message: "Modules Presets sheet not found"
      };
    }

    var oldModulesPresetsValues = SheetsAPI.getDataRange(
      oldModulesSpreadsheetId,
      "Modules Presets"
    );

    var oldModuleTypeIndex = findModuleTypesRowIndex(
      targetModuleTypes,
      oldModulesPresetsValues
    );
    var oldModulesPresets = {};
    var oldModulesInventory = {};
    targetModuleTypes.forEach(function (moduleType) {
      var rowIdx = oldModuleTypeIndex[moduleType] + 1;
      if (typeof rowIdx === "undefined") return;
      oldModulesPresets[moduleType] = {};
      oldModulesInventory[moduleType] = {};
      var row = oldModulesPresetsValues[rowIdx];
      for (var col = 0; col < row.length; col++) {
        var colVal = row[col] != null ? String(row[col]).trim() : "";
        if (colVal === "Module Name") {
          var moduleName = oldModulesPresetsValues[rowIdx + 1][col] != null ? String(oldModulesPresetsValues[rowIdx + 1][col]).trim() : "";
          if (moduleName) {
            var presetName = oldModulesPresetsValues[rowIdx - 1][col] != null ? String(oldModulesPresetsValues[rowIdx - 1][col]).trim() : "";
            if (presetName) {
              oldModulesPresets[moduleType][presetName] = moduleName;
            }
            if (!oldModulesInventory[moduleType].hasOwnProperty(moduleName)) {
              var moduleRarity = oldModulesPresetsValues[rowIdx + 1][col + 1] != null ? String(oldModulesPresetsValues[rowIdx + 1][col + 1]).trim() : "";
              var moduleLevel = oldModulesPresetsValues[rowIdx + 1][col + 2] != null ? String(oldModulesPresetsValues[rowIdx + 1][col + 2]).trim() : "";
              oldModulesInventory[moduleType][moduleName] = {
                rarity: moduleRarity,
                level: moduleLevel,
                substats: [],
              };
              for (
                var substat = rowIdx + 3;
                substat < oldModulesPresetsValues.length;
                substat++
              ) {
                var substatRow = oldModulesPresetsValues[substat];
                var substatColVal = substatRow && substatRow[col] != null ? String(substatRow[col]).trim() : "";
                var substatCol1Val = substatRow && substatRow[col + 1] != null ? String(substatRow[col + 1]).trim() : "";
                if (
                  substatColVal === "" &&
                  substatCol1Val === ""
                ) {
                  break;
                }
                var substats = [substatRow ? substatRow[col] : "", substatRow ? substatRow[col + 1] : ""];
                oldModulesInventory[moduleType][moduleName]["substats"].push(
                  substats
                );
              }
            }
          }
        }
      }
    });
    return {
      oldModulesInventory: oldModulesInventory,
      oldModulesPresets: oldModulesPresets,
    };
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
