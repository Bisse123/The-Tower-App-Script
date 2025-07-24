const modules = {
  importData: function (versionDifference) {
    function importModulesData(versionDifference) {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!newSpreadsheet) {
          console.log(`New spreadsheet not found`);
          return {
            success: false,
            message: "New spreadsheet™ not found",
          };
        }
        var newSheetID = newSpreadsheet.spreadsheetId;

        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        if (!oldSpreadsheet) {
          console.log(`Old spreadsheet not found`);
          return {
            success: false,
            message: "Old spreadsheet™ not found",
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
        var oldDataResult = getVersionFunction();
        if (!oldDataResult || !oldDataResult.success) {
          console.log(
            `Error processing modules data: ${oldDataResult.message}`
          );
          return oldDataResult;
        }

        var targetModuleTypes = oldDataResult.targetModuleTypes || [];
        var oldModulesInventory = oldDataResult.oldModulesInventory || {};
        var oldModulesPresets = oldDataResult.oldModulesPresets || {};
        var oldModulesObtained = oldDataResult.oldModulesObtained || {};

        // Batch fetch all required sheet data
        var requiredRanges = [
          "Modules Inventory",
          "Modules Presets",
          "Mods Obtained",
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

        var inventoryResult = updateModulesInventory(
          targetModuleTypes,
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

        var presetsResult = updateModulesPresets(
          targetModuleTypes,
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

        var obtainedResult = updateModulesObtained(
          targetModuleTypes,
          "Mods Obtained",
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

        if (batchUpdate.length > 0) {
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

    function updateModulesPresets(
      targetModuleTypes,
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
                  range: `${sheetName}!${cellAddress}`,
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

    function updateModulesInventory(
      targetModuleTypes,
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

      var newModuleTypeIndex = findModuleTypesRowIndex(
        targetModuleTypes,
        newModuleInventoryValues
      );

      var batchUpdate = [];

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
              maxLevel = Math.max(
                maxLevel,
                oldModulesInventory[moduleType][cellValue].level
              );

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

          // Update highest level
          if (highestLevelCol !== -1) {
            var highestLevelCell =
              shared.columnToLetter(highestLevelCol + 1) + (rowIdx + 3);
            batchUpdate.push({
              range: `${sheetName}!${highestLevelCell}`,
              values: [[maxLevel]],
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
    }

    function updateModulesObtained(
      targetModuleTypes,
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

      var batchUpdate = [];
      for (var row = 0; row < newModulesObtainedValues.length; row++) {
        for (var col = 0; col < newModulesObtainedValues[row].length; col++) {
          var cellValue = String(newModulesObtainedValues[row][col]).trim();
          if (cellValue) {
            var moduleType = targetModuleTypes.find(function (type) {
              return cellValue.toLowerCase().includes(type);
            });
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

    function version40() {
      try {
        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        var oldSheetID = oldSpreadsheet.spreadsheetId;

        var targetModuleTypes = ["cannon", "armor", "generator", "core"];

        // Batch get all three module sheets at once
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
        var oldModuleTypeIndex = findModuleTypesRowIndex(
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
                var moduleLevel =
                  oldModulesInventoryValues[rowIdx + 2][col + 1] != null
                    ? String(
                        oldModulesInventoryValues[rowIdx + 2][col + 1]
                      ).trim()
                    : "";
                oldModulesInventory[moduleType][moduleName] = {
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
                  oldModulesInventory[moduleType][moduleName]["substats"].push(
                    substats
                  );
                }
              }
            }
          }
        });

        var oldModulesPresetsValues = batchResult[1].values;
        var oldModuleTypeIndex = findModuleTypesRowIndex(
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
                oldModulesPresets[moduleType][presetName] = moduleName;
              }
            }
          }
        });

        var oldModulesObtainedValues = batchResult[2].values;

        var oldModulesObtained = {};
        for (var row = 0; row < oldModulesObtainedValues.length; row++) {
          for (var col = 0; col < oldModulesObtainedValues[row].length; col++) {
            var cellValue = String(oldModulesObtainedValues[row][col]).trim();
            if (cellValue) {
              var moduleType = targetModuleTypes.find(function (type) {
                return cellValue.toLowerCase().includes(type);
              });
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
          targetModuleTypes: targetModuleTypes,
          oldModulesInventory: oldModulesInventory,
          oldModulesPresets: oldModulesPresets,
          oldModulesObtained: oldModulesObtained,
        };
      } catch (error) {
        console.log("Error in version40: " + error.toString());
        return {
          success: false,
          message: "Error in version40: " + error.message,
        };
      }
    }

    var convertVersionFunctions = {
      "v4.0": version40,
    };

    return importModulesData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = ["v4.0"];

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
