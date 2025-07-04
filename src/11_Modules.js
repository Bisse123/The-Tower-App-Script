const modules = {
  convertVersionFunctions: {
    3: this.convertVersion3,
  },

  importData: function (versionDifference) {
    function importModulesData(versionDifference) {
      try {
        var targetModuleTypes = [
          "cannon",
          "armor",
          "generator",
          "core",
        ];
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

        if (versionDifference === 0) {
          // console.log("Same Version");

          // Get old modules inventory data using Sheets API
          var oldModulesInventoryValues =
            SheetsAPI.getDataRange(
              oldSheetID,
              "Modules Inventory"
            )

          var oldModulesInventory = getOldModulesInventory(
            targetModuleTypes,
            oldModulesInventoryValues
          );
          var result = updateModulesInventory(
            targetModuleTypes,
            newSheetID,
            "Modules Inventory",
            oldModulesInventory
          );
          if (!result || !result.success) {
            return {
              success: false,
              message: result.message,
            };
          }

          var batchUpdate = result.batchUpdate || [];

          // Get old modules presets data using Sheets API
          var oldModulesPresetsValues =
            SheetsAPI.getDataRange(
              oldSheetID,
              "Modules Presets"
            )

          var oldModulesPresets = getOldModulesPresets(
            targetModuleTypes,
            oldModulesPresetsValues
          );

          var result = updateModulesPresets(
            targetModuleTypes,
            newSheetID,
            "Modules Presets",
            oldModulesPresets
          );

          if (!result || !result.success) {
            return {
              success: false,
              message: result.message,
            };
          }

          batchUpdate = batchUpdate.concat(result.batchUpdate || []);

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
        }
        // Else do something to convert old version to new one (Future me problem)
        else {
          var { oldModulesInventory, oldModulesPresets } =
            this.convertVersionFunctions[versionDifference](oldSheetID);
          var result = updateModulesInventory(
            targetModuleTypes,
            newSheetID,
            "Modules Inventory",
            oldModulesInventory
          );
          if (!result || !result.success) {
            return {
              success: false,
              message: result.message,
            };
          }

          var batchUpdate = result.batchUpdate || [];

          var result = updateModulesPresets(
            targetModuleTypes,
            newSheetID,
            "Modules Presets",
            oldModulesPresets
          );
          if (!result || !result.success) {
            return {
              success: false,
              message: result.message,
            };
          }
          
          batchUpdate = batchUpdate.concat(result.batchUpdate || []);

          if (batchUpdate.length > 0) {
            SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
            // console.log(`Modules data imported successfully`);
            return {
              success: true,
              message: `Modules data imported successfully`,
            };
          }
          // console.log(`No updates needed for modules data`);
          return {
            success: true,
            message: `No updates needed for modules data`,
          };
        }
        // Check version to figure out which convert function to use
        // Potentially do something where if new version is v4 and old is v2
        // you can do convert v2 -> v3 and then convert v3 -> v4
        // Means more calculations but less rewriting of convert functions when new version are released
      } catch (error) {
        console.log(`Error importing modules data: ${error.toString()}`);
        return {
          success: false,
          message: `Error importing modules data: ${error.message}`,
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
        newModulePresetsValues = SheetsAPI.getDataRange(
          newSheetID,
          sheetName
        );
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
    return importModulesData(versionDifference);
  },

  convertVersion3: function (oldSheetID) {
    var targetModuleTypes = ["cannon", "armor", "generator", "core"];

    var oldSpreadsheet = spreadsheets("oldSpreadsheet", oldSheetID);
    if (!oldSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${oldSheetID}`);
      return {
        success: false,
        message: `New spreadsheet not found with ID: ${oldSheetID}`,
      };
    }
    // Check if Modules Presets sheet exists
    if (!SheetsAPI.getSheetByName(oldSpreadsheet, "Modules Presets")) {
      console.log(`Modules Presets sheet not found`);
      return {
        success: false,
        message: `Modules Presets sheet not found`,
      };
    }

    var oldModulesPresetsValues = SheetsAPI.getDataRange(
      oldSheetID,
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
          var moduleName =
            oldModulesPresetsValues[rowIdx + 1][col] != null
              ? String(oldModulesPresetsValues[rowIdx + 1][col]).trim()
              : "";
          if (moduleName) {
            var presetName =
              oldModulesPresetsValues[rowIdx - 1][col] != null
                ? String(oldModulesPresetsValues[rowIdx - 1][col]).trim()
                : "";
            if (presetName) {
              oldModulesPresets[moduleType][presetName] = moduleName;
            }
            if (!oldModulesInventory[moduleType].hasOwnProperty(moduleName)) {
              var moduleRarity =
                oldModulesPresetsValues[rowIdx + 1][col + 1] != null
                  ? String(oldModulesPresetsValues[rowIdx + 1][col + 1]).trim()
                  : "";
              var moduleLevel =
                oldModulesPresetsValues[rowIdx + 1][col + 2] != null
                  ? String(oldModulesPresetsValues[rowIdx + 1][col + 2]).trim()
                  : "";
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
