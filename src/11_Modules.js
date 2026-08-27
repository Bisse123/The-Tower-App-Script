const modules = {
  // #region Export Functions
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: modules.exportData");
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
        message: "Modules export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting modules data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data, newSheetID) {
    try {
      console.log("Called: modules.importData");

      // Batch get required data for update function only
      var requiredRanges = ["Inventory", "Presets", "Planner v2", "Tracker", "IDS"];
      var dvtIndex = requiredRanges.length;
      var dvtNamedRanges = {
        "Main Efficiency": "DVT_Mod_Assist_Bonus_Level",
        "Substat Efficiency": "DVT_Mod_Assist_Substat_Level",
      };

      Object.keys(dvtNamedRanges).forEach(function (item) {
        requiredRanges.push(dvtNamedRanges[item]);
      });

      var batchUpdate = [];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log("Error getting modules sheet data");
        return {
          success: false,
          message: "Error getting modules sheet data",
        };
      }

      var newModuleInventoryValues = batchResults[0].values;
      var newModulePresetsValues = batchResults[1].values;
      var newModulesPlannerValues = batchResults[2].values;
      var newModulesTrackerValues = batchResults[3].values;
      var idsData = batchResults[4].values;

      var dvtNamedRangesData = {};
      Object.keys(dvtNamedRanges).forEach(function (item) {
        dvtNamedRangesData[item] = batchResults[dvtIndex]
          ? batchResults[dvtIndex].values
          : [];
        dvtIndex++;
      });

      var batchUpdate = [];
      // Only update modules inventory if key exists
      if (data.hasOwnProperty("oldModulesInventory")) {
        var oldModulesInventory = data.oldModulesInventory;
        var inventoryResult = this.updateModulesInventory(
          "Inventory",
          oldModulesInventory,
          newModuleInventoryValues,
        );
        if (!inventoryResult || !inventoryResult.success) {
          return {
            success: false,
            message: inventoryResult.message,
          };
        }
        batchUpdate = batchUpdate.concat(inventoryResult.batchUpdate || []);
      }

      // Only update modules presets if key exists
      if (data.hasOwnProperty("oldModulesPresets")) {
        var oldModulesPresets = data.oldModulesPresets;
        var presetsResult = this.updateModulesPresets(
          "Presets",
          oldModulesPresets,
          newModulePresetsValues,
          dvtNamedRangesData,
        );
        if (!presetsResult || !presetsResult.success) {
          return {
            success: false,
            message: presetsResult.message,
          };
        }
        batchUpdate = batchUpdate.concat(presetsResult.batchUpdate || []);
      }

      // Only update modules planner if key exists
      if (data.hasOwnProperty("oldModulesPlanner")) {
        var oldModulesPlanner = data.oldModulesPlanner;
        var PlannerResult = this.updateModulesInventory(
          "Planner v2",
          oldModulesPlanner,
          newModulesPlannerValues,
        );
        if (!PlannerResult || !PlannerResult.success) {
          return {
            success: false,
            message: PlannerResult.message,
          };
        }
        batchUpdate = batchUpdate.concat(PlannerResult.batchUpdate || []);
      }

      // Only update modules tracker if key exists
      if (data.hasOwnProperty("oldModulesTracker")) {
        var oldModulesTracker = data.oldModulesTracker;
        var trackerResult = this.updateModulesTracker(
          "Tracker",
          oldModulesTracker,
          newModulesTrackerValues,
        );
        if (!trackerResult || !trackerResult.success) {
          return {
            success: false,
            message: trackerResult.message,
          };
        }
        batchUpdate = batchUpdate.concat(trackerResult.batchUpdate || []);
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Modules",
        newSheetID,
        idsData,
        data.idMasterID,
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
        message:
          batchUpdate.length > 1
            ? `Modules data imported successfully`
            : "No modules data to update, but ID setting completed",
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing modules data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Update Functions
  updateModulesPresets: function (
    sheetName,
    oldModulesPresets,
    newModulePresetsValues,
    dvtNamedRangesData,
  ) {
    try {
      console.log("Called: modules.updateModulesPresets");
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
        newModulePresetsValues,
      );

      var presetColIndices = [];
      for (var row = 0; row < newModulePresetsValues.length; row++) {
        (newModulePresetsValues[row] || []).forEach(function (cell, index) {
          if (String(cell).trim() === "Primary Slot") {
            presetColIndices.push(index);
          }
        });
        if (presetColIndices.length > 0) {
          break;
        }
      }

      if (presetColIndices.length === 0) {
        console.log(`Could not find "Primary Slot" columns in sheet`);
        return {
          success: false,
          message: `Could not find "Primary Slot" columns in sheet`,
        };
      }

      var presetNames = oldModulesPresets.presetNames || [];

      var moduleContexts = {};
      targetModuleTypes.forEach(function (moduleType) {
        if (!oldModulesPresets.hasOwnProperty(moduleType)) return;
        if (typeof newModuleTypeIndex[moduleType] === "undefined") return;
        var rowIdx = newModuleTypeIndex[moduleType] + 1;
        moduleContexts[moduleType] = {
          rowIdx: rowIdx,
          row: newModulePresetsValues[rowIdx],
        };
      });

      var batchUpdate = [];
      if (Object.keys(moduleContexts).length === 0) {
        return {
          success: false,
          message: `Could not find preset names or module rows for updating presets`,
        };
      }

      presetNames.forEach(function (presetName, slot) {
        var presetCol = presetColIndices[slot];
        if (!presetName || typeof presetCol === "undefined") {
          return;
        }

        targetModuleTypes.forEach(function (moduleType) {
          var context = moduleContexts[moduleType];
          if (!context) {
            return;
          }
          var modulePresets = oldModulesPresets[moduleType];
          if (!modulePresets.hasOwnProperty(presetName)) {
            return;
          }
          var presetData = modulePresets[presetName];
          if (!presetData) {
            return;
          }
          var rowIdx = context.rowIdx;
          var currentName = String(context.row[presetCol] || "").trim();
          if (
            currentName !== String(presetName).trim() &&
            moduleType === "cannon"
          ) {
            var presetNameRange = `${sheetName}!${shared.columnToLetter(presetCol + 1)}${rowIdx + 1}`;
            var presetNameValues = [[presetName]];
            batchUpdate.push({
              range: presetNameRange,
              values: presetNameValues,
            });
          }

          var presetModuleRange = `${sheetName}!${shared.columnToLetter(
            presetCol + 2,
          )}${rowIdx + 3}:${shared.columnToLetter(presetCol + 2)}${rowIdx + 4}`;
          var presetModuleValues = [
            [presetData.primary || ""],
            [presetData.secondary || ""],
          ];
          batchUpdate.push({
            range: presetModuleRange,
            values: presetModuleValues,
          });
        });
      });

      targetModuleTypes.forEach(function (moduleType) {
        var context = moduleContexts[moduleType];
        if (!context) {
          return;
        }
        var modulePresets = oldModulesPresets[moduleType];
        if (!modulePresets.hasOwnProperty("Assist Slot")) {
          return;
        }
        var presetData = modulePresets["Assist Slot"];
        if (!presetData) {
          return;
        }
        var row = context.row;
        var rowIdx = context.rowIdx;
        var presetCol = row.indexOf("Assist Slot");
        if (presetCol === -1) {
          return;
        }
        var lockedRange = `${sheetName}!${shared.columnToLetter(
          presetCol + 3,
        )}${rowIdx + 1}:${shared.columnToLetter(presetCol + 3)}${rowIdx + 1}`;
        var lockedValues = [[presetData.locked || false]];
        var rarityRange = `${sheetName}!${shared.columnToLetter(
          presetCol + 2,
        )}${rowIdx + 2}:${shared.columnToLetter(presetCol + 2)}${rowIdx + 2}`;
        var rarityValues = [[presetData.rarity || null]];
        var multiSubRange = `${sheetName}!${shared.columnToLetter(
          presetCol + 3,
        )}${rowIdx + 3}:${shared.columnToLetter(presetCol + 3)}${rowIdx + 4}`;

        // Transform values using DVT
        var dvtMultiplier = shared.getDVTValue(
          presetData.multiplier || null,
          dvtNamedRangesData["Main Efficiency"],
        );
        var dvtSubstat = shared.getDVTValue(
          presetData.substat || null,
          dvtNamedRangesData["Substat Efficiency"],
        );

        var multiSubValues = [[dvtMultiplier], [dvtSubstat]];
        batchUpdate.push({
          range: lockedRange,
          values: lockedValues,
        });
        batchUpdate.push({
          range: rarityRange,
          values: rarityValues,
        });
        batchUpdate.push({
          range: multiSubRange,
          values: multiSubValues,
        });
      });
      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: `Modules presets updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for modules presets`,
      };
    } catch (error) {
      console.log("Error in updateModulesPresets:", error.toString());
      return {
        success: false,
        message: "Error in updateModulesPresets: " + error.message,
      };
    }
  },

  updateModulesInventory: function (
    sheetName,
    oldModulesInventory,
    newModuleInventoryValues,
  ) {
    try {
      console.log("Called: modules.updateModulesInventory");
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
        newModuleInventoryValues,
      );
      var batchUpdate = [];
      var usedSpares = {};
      targetModuleTypes.forEach(function (moduleType) {
        if (oldModulesInventory.hasOwnProperty(moduleType)) {
          var rowIdx = newModuleTypeIndex[moduleType];
          if (typeof rowIdx === "undefined") return;
          var row = newModuleInventoryValues[rowIdx];
          for (var col = 1; col < row.length; col++) {
            var cellValue = String(row[col]);
            if (
              cellValue.trim().includes("Spare") ||
              (cellValue.trim().includes("Any Other") &&
                !oldModulesInventory[moduleType].hasOwnProperty(cellValue))
            ) {
              for (var spare in oldModulesInventory[moduleType]) {
                if (spare.includes("Spare") && !usedSpares[spare]) {
                  batchUpdate.push({
                    range: `${sheetName}!${shared.columnToLetter(col + 1)}${
                      rowIdx + 1
                    }`,
                    values: [[spare]],
                  });
                  cellValue = spare;
                  usedSpares[spare] = true;
                  break;
                }
              }
            }
            if (cellValue.trim().includes("Spare") && !usedSpares[cellValue]) {
              usedSpares[cellValue] = true;
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
          if (oldModulesInventory[moduleType].hasOwnProperty("Highest Level")) {
            var highestLevelCol = newModuleInventoryValues[rowIdx + 1].indexOf("Highest Level");
            if (highestLevelCol !== -1) {
              var maxLevel = oldModulesInventory[moduleType]["Highest Level"] || null;
              var highestLevelCell =
                shared.columnToLetter(highestLevelCol + 1) + (rowIdx + 3);
              batchUpdate.push({
                range: `${sheetName}!${highestLevelCell}`,
                values: [[maxLevel]],
              });
            }
          }
          if (oldModulesInventory[moduleType].hasOwnProperty("Assist Level")) {
            var assistLevelCol = newModuleInventoryValues[rowIdx + 4].indexOf("Assist Level");
            if (assistLevelCol !== -1) {
              var assistLevel = oldModulesInventory[moduleType]["Assist Level"] || null;
              var assistLevelCell =
                shared.columnToLetter(assistLevelCol + 1) + (rowIdx + 6);
              batchUpdate.push({
                range: `${sheetName}!${assistLevelCell}`,
                values: [[assistLevel]],
              });
            }
          }
          if (oldModulesInventory[moduleType].hasOwnProperty("Dice")) {
            var dicecol = newModuleInventoryValues[rowIdx + 7].indexOf("Dice");
            if (dicecol !== -1) {
              var diceValue = oldModulesInventory[moduleType]["Dice"] || null;
              var diceCell =
                shared.columnToLetter(dicecol + 1) + (rowIdx + 9);
              batchUpdate.push({
                range: `${sheetName}!${diceCell}`,
                values: [[diceValue]],
              });
            }
          }
        }
      });
      return {
        success: true,
        message: `Modules inventory updated successfully`,
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      console.log("Error in updateModulesInventory:", error.toString());
      return {
        success: false,
        message: "Error in updateModulesInventory: " + error.message,
      };
    }
  },

  updateModulesTracker: function (
    sheetName,
    oldModulesTracker,
    newModulesTrackerValues,
  ) {
    try {
      console.log("Called: modules.updateModulesTracker");
      if (!newModulesTrackerValues) {
        console.log("Could not read modules tracker data");
        return {
          success: false,
          message: "Could not read modules tracker data",
        };
      }
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var batchUpdate = [];
      for (var row = 0; row < newModulesTrackerValues.length; row++) {
        var rowData = newModulesTrackerValues[row];
        var inputColIdx = rowData.indexOf("Input values");
        if (inputColIdx !== -1) {
          var inputValues = [];
          for (
            var inputRow = row + 1;
            inputRow < newModulesTrackerValues.length;
            inputRow++
          ) {
            var inputRowData = newModulesTrackerValues[inputRow];
            var inputKey = inputRowData[inputColIdx] || null;
            if (!inputKey) {
              break;
            }
            if (oldModulesTracker["Input values"].hasOwnProperty(inputKey)) {
              var inputValue =
                oldModulesTracker["Input values"][inputKey] || null;
              inputValues.push([inputValue]);
            }
          }
          if (inputValues.length > 0) {
            var inputColLetter = shared.columnToLetter(inputColIdx + 5);
            var inputRowStart = row + 2;
            var inputRowEnd = inputRowStart + inputValues.length - 1;
            var inputRange = `${sheetName}!${inputColLetter}${inputRowStart}:${inputColLetter}${inputRowEnd}`;
            batchUpdate.push({
              range: inputRange,
              values: inputValues,
            });
          }
        }

        if (
          oldModulesTracker.hasOwnProperty("summary") &&
          rowData.some((cell) => String(cell).toLowerCase().includes("summary"))
        ) {
          rowData.forEach(function (cell, idx) {
            targetModuleTypes.forEach(function (type) {
              if (
                oldModulesTracker["summary"].hasOwnProperty(type) &&
                String(cell).toLowerCase().includes(type) &&
                String(cell).toLowerCase().includes("summary")
              ) {
                var summaryData = oldModulesTracker["summary"][type];
                for (
                  var rowIdx = row + 2;
                  rowIdx < newModulesTrackerValues.length;
                  rowIdx++
                ) {
                  var summaryRowData = newModulesTrackerValues[rowIdx];
                  var summaryModuleName = summaryRowData[idx] || null;
                  if (summaryModuleName.toLowerCase().includes("total")) {
                    break;
                  }
                  if (summaryData.hasOwnProperty(summaryModuleName)) {
                    var summaryModuleCount = summaryData[summaryModuleName];
                    batchUpdate.push({
                      range: `${sheetName}!${shared.columnToLetter(idx + 2)}${rowIdx + 1}`,
                      values: [[summaryModuleCount]],
                    });
                  }
                }
              }
            });
          });
        }

        var colIdx = rowData.findIndex(
          (cell) =>
            targetModuleTypes.some((type) =>
              String(cell).toLowerCase().includes(type),
            ) &&
            !String(cell).toLowerCase().includes("summary") &&
            !String(cell).toLowerCase().includes("quantity"),
        );
        if (colIdx === -1) {
          continue;
        }
        var targetModule = targetModuleTypes.find(function (type) {
          return String(rowData[colIdx]).toLowerCase().includes(type);
        });
        for (var col = colIdx; col < rowData.length; col++) {
          var moduleName = String(rowData[col]).trim();
          if (
            oldModulesTracker[targetModule] &&
            oldModulesTracker[targetModule].hasOwnProperty(moduleName)
          ) {
            var startRow = row + 4;
            var rangeCol = col + 2;
            if (moduleName === "Fodders") {
              startRow -= 1;
              rangeCol += 1;
            }
            var endRow =
              startRow + oldModulesTracker[targetModule][moduleName].length - 1;
            var range = `${sheetName}!${shared.columnToLetter(
              rangeCol,
            )}${startRow}:${shared.columnToLetter(rangeCol)}${endRow}`;
            var values = oldModulesTracker[targetModule][moduleName].map(
              function (copy) {
                return [copy || null];
              },
            );
            batchUpdate.push({
              range: range,
              values: values,
            });
            if (
              oldModulesTracker[targetModule].hasOwnProperty(
                moduleName + " Shattered",
              )
            ) {
              var shatteredRow = row + 9;
              var shatteredCol = col + 3;
              var shatteredRange = `${sheetName}!${shared.columnToLetter(shatteredCol)}${shatteredRow}`;
              var shatteredValue = [
                [
                  oldModulesTracker[targetModule][moduleName + " Shattered"] ||
                    null,
                ],
              ];
              batchUpdate.push({
                range: shatteredRange,
                values: shatteredValue,
              });
            }
          }
        }
      }
      return {
        success: true,
        message: `Modules Tracker updated successfully`,
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      console.log("Error in updateModulesTracker:", error.toString());
      return {
        success: false,
        message: "Error in updateModulesTracker: " + error.message,
      };
    }
  },

  findModuleTypesRowIndex: function (targetModuleTypes, moduleRange) {
    try {
      console.log("Called: modules.findModuleTypesRowIndex");
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
    } catch (error) {
      console.log("Error in findModuleTypesRowIndex:", error.toString());
      return {};
    }
  },

  // #endregion
  // #region Convert Versions
  version6_4_3: function (oldSheetID) {
    try {
      console.log("Called: modules.version6_4_3");

      var ranges = ["Inventory", "Presets", "Planner v2", "Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length === 0) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesPlannerValues = batchResult[2].values;
      var oldModulesTrackerValues = batchResult[3].values;

      var formulaRanges = ["Tracker"];
      var formulaBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulaRanges,
      );
      if (!formulaBatchResult || formulaBatchResult.length === 0) {
        console.log("Could not read module formulas from old spreadsheet");
        return {
          success: false,
          message: "Could not read module formulas from old spreadsheet",
        };
      }
      var oldModulesTrackerFormulas = formulaBatchResult[0].values;

      var inventoryData = this.getVersion5_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = this.getVersion5_0ModulesPresets(
        oldModulesPresetsValues,
      );
      var plannerData = this.getVersion5_0ModulesInventory(
        oldModulesPlannerValues,
      );
      var trackerData = this.getVersion4_7ModulesTracker(
        oldModulesTrackerValues,
        oldModulesTrackerFormulas,
      );

      var success =
        inventoryData.success && presetsData.success && plannerData.success && trackerData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesPlanner: plannerData.oldModulesInventory || {},
        oldModulesTracker: trackerData.oldModulesTracker || {},
      };
    } catch (error) {
      console.log("Error in version6_4_3: " + error.toString());
      return {
        success: false,
        message: "Error in version6_4_3: " + error.message,
      };
    }
  },

  version5_2_1: function (oldSheetID) {
    try {
      console.log("Called: modules.version5_2_1");

      var ranges = ["Inventory", "Presets", "Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length === 0) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesTrackerValues = batchResult[2].values;

      var formulaRanges = ["Tracker"];
      var formulaBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulaRanges,
      );
      if (!formulaBatchResult || formulaBatchResult.length === 0) {
        console.log("Could not read module formulas from old spreadsheet");
        return {
          success: false,
          message: "Could not read module formulas from old spreadsheet",
        };
      }
      var oldModulesTrackerFormulas = formulaBatchResult[0].values;

      var inventoryData = this.getVersion5_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = this.getVersion5_0ModulesPresets(
        oldModulesPresetsValues,
      );
      var trackerData = this.getVersion4_7ModulesTracker(
        oldModulesTrackerValues,
        oldModulesTrackerFormulas,
      );

      var success =
        inventoryData.success && presetsData.success && trackerData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesTracker: trackerData.oldModulesTracker || {},
      };
    } catch (error) {
      console.log("Error in version5_2_1: " + error.toString());
      return {
        success: false,
        message: "Error in version5_2_1: " + error.message,
      };
    }
  },

  version5_0: function (oldSheetID) {
    try {
      console.log("Called: modules.version5_0");

      var ranges = ["Modules Inventory", "Modules Presets", "Modules Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length === 0) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesTrackerValues = batchResult[2].values;

      var formulaRanges = ["Tracker"];
      var formulaBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulaRanges,
      );
      if (!formulaBatchResult || formulaBatchResult.length === 0) {
        console.log("Could not read module formulas from old spreadsheet");
        return {
          success: false,
          message: "Could not read module formulas from old spreadsheet",
        };
      }
      var oldModulesTrackerFormulas = formulaBatchResult[0].formulas;

      var inventoryData = this.getVersion5_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = this.getVersion5_0ModulesPresets(
        oldModulesPresetsValues,
      );
      var trackerData = this.getVersion4_7ModulesTracker(
        oldModulesTrackerValues,
        oldModulesTrackerFormulas,
      );

      var success =
        inventoryData.success && presetsData.success && trackerData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesTracker: trackerData.oldModulesTracker || {},
      };
    } catch (error) {
      console.log("Error in version5_0: " + error.toString());
      return {
        success: false,
        message: "Error in version5_0: " + error.message,
      };
    }
  },

  version4_7: function (oldSheetID) {
    try {
      console.log("Called: modules.version4_7");

      var ranges = ["Modules Inventory", "Modules Presets", "Modules Tracker"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length < 2) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;
      var oldModulesTrackerValues = batchResult[2].values;

      var formulaRanges = ["Tracker"];
      var formulaBatchResult = SheetsAPI.batchGetFormulas(
        oldSheetID,
        formulaRanges,
      );
      if (!formulaBatchResult || formulaBatchResult.length === 0) {
        console.log("Could not read module formulas from old spreadsheet");
        return {
          success: false,
          message: "Could not read module formulas from old spreadsheet",
        };
      }
      var oldModulesTrackerFormulas = formulaBatchResult[0].formulas;

      var inventoryData = this.getVersion4_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = this.getVersion4_0ModulesPresets(
        oldModulesPresetsValues,
      );
      var trackerData = this.getVersion4_7ModulesTracker(
        oldModulesTrackerValues,
        oldModulesTrackerFormulas,
      );

      var success =
        inventoryData.success && presetsData.success && trackerData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
        oldModulesTracker: trackerData.oldModulesTracker || {},
      };
    } catch (error) {
      console.log("Error in version4_7: " + error.toString());
      return {
        success: false,
        message: "Error in version4_7: " + error.message,
      };
    }
  },

  version4_0: function (oldSheetID) {
    try {
      console.log("Called: modules.version4_0");

      var ranges = ["Modules Inventory", "Modules Presets"];
      var batchResult = SheetsAPI.batchGetValues(oldSheetID, ranges);

      if (!batchResult || batchResult.length < 2) {
        console.log("Could not read module data from old spreadsheet");
        return {
          success: false,
          message: "Could not read module data from old spreadsheet",
        };
      }

      var oldModulesInventoryValues = batchResult[0].values;
      var oldModulesPresetsValues = batchResult[1].values;

      var inventoryData = this.getVersion4_0ModulesInventory(
        oldModulesInventoryValues,
      );
      var presetsData = this.getVersion4_0ModulesPresets(
        oldModulesPresetsValues,
      );

      var success = inventoryData.success && presetsData.success;

      return {
        success: success,
        message: success
          ? "Modules data retrieved successfully"
          : "Error retrieving Modules data",
        oldModulesInventory: inventoryData.oldModulesInventory || {},
        oldModulesPresets: presetsData.oldModulesPresets || {},
      };
    } catch (error) {
      console.log("Error in version4_0: " + error.toString());
      return {
        success: false,
        message: "Error in version4_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Modules Inventory
  getVersion5_0ModulesInventory: function (oldModulesInventoryValues) {
    try {
      console.log("Called: modules.getVersion5_0ModulesInventory");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = this.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesInventoryValues,
      );

      var oldModulesInventory = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType];
        if (typeof rowIdx === "undefined") return;
        oldModulesInventory[moduleType] = {};
        var row = oldModulesInventoryValues[rowIdx];

        const highestLevelRow = oldModulesInventoryValues[rowIdx + 1];
        var highestLevelCol = highestLevelRow ? highestLevelRow.indexOf("Highest Level") : -1;
        if (highestLevelCol !== -1) {
          oldModulesInventory[moduleType]["Highest Level"] =
            oldModulesInventoryValues[rowIdx + 2][highestLevelCol];
        }
        const assistLevelRow = oldModulesInventoryValues[rowIdx + 4];
        var assistLevelCol = assistLevelRow ? assistLevelRow.indexOf("Assist Level") : -1;
        if (assistLevelCol !== -1) {
          oldModulesInventory[moduleType]["Assist Level"] =
            oldModulesInventoryValues[rowIdx + 5][assistLevelCol];
        }
        const diceRow = oldModulesInventoryValues[rowIdx + 7];
        var diceCol = diceRow ? diceRow.indexOf("Dice") : -1;
        if (diceCol !== -1) {
          oldModulesInventory[moduleType]["Dice"] =
            oldModulesInventoryValues[rowIdx + 8][diceCol];
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
                substat < rowIdx + 4 + 8;
                substat++
              ) {
                var substatRow = oldModulesInventoryValues[substat];
                var substatData = [
                  substatRow ? substatRow[col] : "",
                  substatRow ? substatRow[col + 1] : "",
                ];
                oldModulesInventory[moduleType][moduleName]["substats"].push(
                  substatData,
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
      console.log(
        "Error in getVersion5_0ModulesInventory: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion5_0ModulesInventory: " + error.message,
      };
    }
  },

  getVersion4_0ModulesInventory: function (oldModulesInventoryValues) {
    try {
      console.log("Called: modules.getVersion4_0ModulesInventory");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = this.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesInventoryValues,
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
              if (
                removedRarity.includes(moduleRarity) &&
                !moduleName.includes("Any Other")
              ) {
                moduleRarity = "Epic";
              }
              oldModulesInventory[moduleType][moduleName] = {
                rarity: moduleRarity,
                substats: [],
              };
              for (
                var substat = rowIdx + 4;
                substat < rowIdx + 4 + 8;
                substat++
              ) {
                var substatRow = oldModulesInventoryValues[substat];
                var substatData = [
                  substatRow ? substatRow[col] : "",
                  substatRow ? substatRow[col + 1] : "",
                ];
                oldModulesInventory[moduleType][moduleName]["substats"].push(
                  substatData,
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
      console.log(
        "Error in getVersion4_0ModulesInventory: " + error.toString(),
      );
      return {
        success: false,
        message: "Error in getVersion4_0ModulesInventory: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Modules Presets
  getVersion5_0ModulesPresets: function (oldModulesPresetsValues) {
    try {
      console.log("Called: modules.getVersion5_0ModulesPresets");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = this.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesPresetsValues,
      );

      var oldModulesPresets = {};
      var presetNames = [];
      var presetNamesLocked = false;
      targetModuleTypes.forEach(function (moduleType) {
        if (typeof oldModuleTypeIndex[moduleType] === "undefined") return;
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        oldModulesPresets[moduleType] = {};
        var row = oldModulesPresetsValues[rowIdx + 2];
        var presetIndex = 0;
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Primary Slot") {
            var rawPresetName = oldModulesPresetsValues[rowIdx][col]
              ? String(oldModulesPresetsValues[rowIdx][col]).trim()
              : "";
            if (!rawPresetName) {
              continue;
            }
            var presetName;
            if (presetIndex < presetNames.length) {
              presetName = presetNames[presetIndex];
            } else if (!presetNamesLocked) {
              presetName = rawPresetName;
              presetNames.push(rawPresetName);
            } else {
              break;
            }
            var primaryName = oldModulesPresetsValues[rowIdx + 2][col + 1]
              ? String(oldModulesPresetsValues[rowIdx + 2][col + 1]).trim()
              : "";
            var secondaryName = oldModulesPresetsValues[rowIdx + 3][col + 1]
              ? String(oldModulesPresetsValues[rowIdx + 3][col + 1]).trim()
              : "";
            oldModulesPresets[moduleType][presetName] = {
              primary: primaryName,
              secondary: secondaryName,
            };
            presetIndex++;
          }
        }
        if (presetIndex > 0) {
          presetNamesLocked = true;
        }
        var assistSlotCol =
          oldModulesPresetsValues[rowIdx].indexOf("Assist Slot");
        if (assistSlotCol !== -1) {
          var assistLocked = oldModulesPresetsValues[rowIdx][assistSlotCol + 2];

          var assistRarity = String(
            oldModulesPresetsValues[rowIdx + 1][assistSlotCol + 1],
          ).trim();
          var assistMultiplier = String(
            oldModulesPresetsValues[rowIdx + 2][assistSlotCol + 2],
          ).trim();
          var assistSubstat = String(
            oldModulesPresetsValues[rowIdx + 3][assistSlotCol + 2],
          ).trim();
          oldModulesPresets[moduleType]["Assist Slot"] = {
            locked: assistLocked,
            rarity: assistRarity,
            multiplier: assistMultiplier,
            substat: assistSubstat,
          };
        }
      });
      oldModulesPresets.presetNames = shared.resolvePresetOrder(
        presetNames,
        shared.templatePresetNames,
      ).order;

      return {
        success: true,
        oldModulesPresets: oldModulesPresets,
      };
    } catch (error) {
      console.log("Error in getVersion5_0ModulesPresets: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion5_0ModulesPresets: " + error.message,
      };
    }
  },

  getVersion4_0ModulesPresets: function (oldModulesPresetsValues) {
    try {
      console.log("Called: modules.getVersion4_0ModulesPresets");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = this.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesPresetsValues,
      );

      var oldModulesPresets = {};
      var presetNames = [];
      var presetNamesLocked = false;
      targetModuleTypes.forEach(function (moduleType) {
        if (typeof oldModuleTypeIndex[moduleType] === "undefined") return;
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        oldModulesPresets[moduleType] = {};
        var row = oldModulesPresetsValues[rowIdx];
        var presetIndex = 0;
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Module Name") {
            var rawPresetName = String(
              oldModulesPresetsValues[rowIdx - 1][col],
            ).trim();
            var moduleName = String(
              oldModulesPresetsValues[rowIdx + 1][col],
            ).trim();
            if (rawPresetName && moduleName) {
              var presetName;
              if (presetIndex < presetNames.length) {
                presetName = presetNames[presetIndex];
              } else if (!presetNamesLocked) {
                presetName = rawPresetName;
                presetNames.push(rawPresetName);
              } else {
                break;
              }
              oldModulesPresets[moduleType][presetName] = {
                primary: moduleName,
                secondary: "",
              };
              presetIndex++;
            }
          }
        }
        if (presetIndex > 0) {
          presetNamesLocked = true;
        }
      });
      oldModulesPresets.presetNames = presetNames;

      return {
        success: true,
        oldModulesPresets: oldModulesPresets,
      };
    } catch (error) {
      console.log("Error in getVersion4_0ModulesPresets: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion4_0ModulesPresets: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Modules Tracker
  getVersion4_7ModulesTracker: function (
    oldModulesTrackerValues,
    oldModulesTrackerFormulas,
  ) {
    try {
      console.log("Called: modules.getVersion4_7ModulesTracker");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModulesTracker = {};

      for (var row = 0; row < oldModulesTrackerValues.length; row++) {
        var rowData = oldModulesTrackerValues[row];
        var inputColIdx = rowData.indexOf("Input values");
        if (inputColIdx !== -1) {
          oldModulesTracker["Input values"] = {};
          for (
            var inputRow = row + 1;
            inputRow < oldModulesTrackerValues.length;
            inputRow++
          ) {
            var inputRowData = oldModulesTrackerValues[inputRow];
            var inputKey = inputRowData[inputColIdx] || null;
            var inputValue = inputRowData[inputColIdx + 4] || null;
            if (inputKey) {
              oldModulesTracker["Input values"][inputKey] = inputValue;
            } else {
              break;
            }
          }
        }

        if (
          rowData.some((cell) => String(cell).toLowerCase().includes("summary"))
        ) {
          rowData.forEach(function (cell, idx) {
            targetModuleTypes.forEach(function (type) {
              if (
                String(cell).toLowerCase().includes(type) &&
                String(cell).toLowerCase().includes("summary")
              ) {
                for (
                  var rowIdx = row + 2;
                  rowIdx < oldModulesTrackerFormulas.length;
                  rowIdx++
                ) {
                  var summaryRowData = oldModulesTrackerFormulas[rowIdx];
                  var summaryModuleName = summaryRowData[idx] || null;
                  var summaryModuleCount = summaryRowData[idx + 1] || null;
                  if (summaryModuleName.toLowerCase().includes("total")) {
                    break;
                  }
                  if (
                    summaryModuleName &&
                    !String(summaryModuleCount).trim().startsWith("=")
                  ) {
                    if (!oldModulesTracker["summary"]) {
                      oldModulesTracker["summary"] = {};
                    }
                    if (!oldModulesTracker["summary"][type]) {
                      oldModulesTracker["summary"][type] = {};
                    }
                    oldModulesTracker["summary"][type][summaryModuleName] =
                      summaryModuleCount;
                  }
                }
              }
            });
          });
        }

        var colIdx = rowData.findIndex(
          (cell) =>
            targetModuleTypes.some((type) =>
              String(cell).toLowerCase().includes(type),
            ) &&
            !String(cell).toLowerCase().includes("summary") &&
            !String(cell).toLowerCase().includes("quantity") &&
            !String(cell).toLowerCase().includes("score"),
        );
        if (colIdx === -1) {
          continue;
        }
        var targetModule = targetModuleTypes.find(function (type) {
          return String(rowData[colIdx]).toLowerCase().includes(type);
        });
        for (var col = colIdx; col < rowData.length; col++) {
          var moduleName = String(rowData[col]).trim();
          if (
            moduleName &&
            oldModulesTrackerValues[row + 3][col] === "Copies"
          ) {
            if (!oldModulesTracker[targetModule]) {
              oldModulesTracker[targetModule] = {};
            }
            oldModulesTracker[targetModule][moduleName] = [];
            for (var subRow = row + 3; subRow < row + 7; subRow++) {
              var subRowData = oldModulesTrackerValues[subRow];

              var copy = subRowData[col + 1] || null;
              oldModulesTracker[targetModule][moduleName].push(copy);
            }

            if (oldModulesTrackerValues[row + 8][col] === "Shattered epic") {
              oldModulesTracker[targetModule][moduleName + " Shattered"] =
                oldModulesTrackerValues[row + 8][col + 2] || null;
            }
          } else if (moduleName === "Fodders") {
            if (!oldModulesTracker[targetModule]) {
              oldModulesTracker[targetModule] = {};
            }
            oldModulesTracker[targetModule][moduleName] = [];
            for (var subRow = row + 2; subRow < row + 8; subRow++) {
              var subRowData = oldModulesTrackerValues[subRow];

              var fodder = subRowData[col + 2] || null;
              oldModulesTracker[targetModule][moduleName].push(fodder);
            }
            row = row + 8;
            break;
          }
        }
        // summaryIndexes now contains all indexes for each type in this row
      }
      return {
        success: true,
        oldModulesTracker: oldModulesTracker,
      };
    } catch (error) {
      console.log("Error in getVersion4_7ModulesTracker: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion4_7ModulesTracker: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseModulesData: function (data) {
    const moduleNames = {
      7: { name: "Havoc Bringer", category: "Cannon" },
      8: { name: "Death Penalty", category: "Cannon" },
      9: { name: "Being Annihilator", category: "Cannon" },
      10: { name: "Astral Deliverance", category: "Cannon" },
      17: { name: "Wormhole Redirector", category: "Armor" },
      18: { name: "Negative Mass Projector", category: "Armor" },
      19: { name: "Space Displacer", category: "Armor" },
      20: { name: "Anti-Cube Portal", category: "Armor" },
      27: { name: "Black Hole Digestor", category: "Generator" },
      28: { name: "Pulsar Harvester", category: "Generator" },
      29: { name: "Galaxy Compressor", category: "Generator" },
      30: { name: "Singularity Harness", category: "Generator" },
      37: { name: "Multiverse Nexus", category: "Core" },
      38: { name: "Dimension Core", category: "Core" },
      39: { name: "Harmony Conductor", category: "Core" },
      40: { name: "Om Chip", category: "Core" },
      41: { name: "Shrink Ray", category: "Cannon" },
      42: { name: "Sharp Fortitude", category: "Armor" },
      43: { name: "Project Funding", category: "Generator" },
      44: { name: "Magnetic Hook", category: "Core" },
      45: { name: "Amplifying Strike", category: "Cannon" },
      46: { name: "Orbital Augment", category: "Armor" },
      47: { name: "Restorative Bonus", category: "Generator" },
      48: { name: "Primordial Collapse", category: "Core" },
      49: { name: "Gilded Sniper", category: "Cannon" },
      50: { name: "Sentry Protocol", category: "Armor" },
      51: { name: "New Generator", category: "Generator" },
      52: { name: "Tactical Barrage", category: "Core" },
    };

    const moduleRarities = {
      1: "Common",
      2: "Rare",
      3: "Rare+",
      4: "Epic",
      5: "Epic+",
      6: "Legendary",
      7: "Legendary+",
      8: "Mythic",
      9: "Mythic+",
      10: "Ancestral",
      11: "Ancestral 1*",
      12: "Ancestral 2*",
      13: "Ancestral 3*",
      14: "Ancestral 4*",
      15: "Ancestral 5*",
    };

    const moduleCategories = {
      0: "Cannon",
      1: "Armor",
      2: "Generator",
      3: "Core",
    };

    const assistRarities = ["Epic", "Legendary", "Mythic", "Ancestral"];

    function lookupEffectID(effectID) {
      const effectRarities = [
        "Common",
        "Rare",
        "Epic",
        "Legendary",
        "Mythic",
        "Ancestral",
      ];

      // [label, n] - rarityStart = 6 - n, effectIDs assigned sequentially
      const substatClusters = [
        // Cannon (C=1-17)
        ["Attack Speed", 6],
        ["Critical Chance", 6],
        ["Critical Factor", 6],
        ["Attack Range", 6],
        ["Damage / Meter", 6],
        ["Multishot Chance", 5],
        ["Multishot Targets", 4],
        ["Rapid Fire Chance", 5],
        ["Rapid Fire Duration", 5],
        ["Bounce Shot Chance", 5],
        ["Bounce Shot Targets", 4],
        ["Bounce Shot Range", 5],
        ["Super Crit Chance", 4],
        ["Super Crit Multi", 4],
        ["Rend Armor Chance", 3],
        ["Rend Armor Multi", 3],
        ["Max Rend Armor Multi", 3],
        // Armor (C=18-34)
        ["Health Regen", 6],
        ["Defense %", 6],
        ["Defense Absolute", 6],
        ["Thorns Damage", 4],
        ["Lifesteal", 4],
        ["Knockback Chance", 4],
        ["Knockback Force", 4],
        ["Orb Speed", 4],
        ["Orbs", 2],
        ["Shockwave Size", 4],
        ["Shockwave Frequency", 4],
        ["Land Mine Chance", 5],
        ["Land Mine Damage", 5],
        ["Land Mine Radius", 5],
        ["Death Defy", 3],
        ["Wall Health", 4],
        ["Wall Rebuild", 4],
        // Generator (C=35-46)
        ["Cash Bonus", 6],
        ["Cash / Wave", 6],
        ["Coins / Kill Bonus", 6],
        ["Coins / Wave", 6],
        ["Free Attack Upgrade", 6],
        ["Free Defense Upgrade", 6],
        ["Free Utility Upgrade", 6],
        ["Interest / Wave", 4],
        ["Recovery Amount", 4],
        ["Package Chance", 4],
        ["Enemy Attack Level Skip", 4],
        ["Enemy Health Level Skip", 4],
        // Core (C=47-73)
        ["Chain Lightning - Damage", 6],
        ["Chain Lightning - Quantity", 4],
        ["Chain Lightning - Chance", 6],
        ["Smart Missiles - Damage", 6],
        ["Smart Missiles - Quantity", 4],
        ["Smart Missiles - Cooldown", 3],
        ["Death Wave - Damage", 6],
        ["Death Wave - Quantity", 3],
        ["Death Wave - Cooldown", 3],
        ["Chrono Field - Duration", 3],
        ["Chrono Field - Speed Reduction", 4],
        ["Chrono Field - Cooldown", 3],
        ["Inner Land Mines - Damage", 6],
        ["Inner Land Mines - Quantity", 3],
        ["Inner Land Mines - Cooldown", 4],
        ["Golden Tower - Bonus", 4],
        ["Golden Tower - Duration", 3],
        ["Golden Tower - Cooldown", 3],
        ["Poison Swamp - Damage", 6],
        ["Poison Swamp - Duration", 3],
        ["Poison Swamp - Cooldown", 5],
        ["Black Hole - Size", 6],
        ["Black Hole - Duration", 3],
        ["Black Hole - Cooldown", 3],
        ["Spotlight - Bonus", 6],
        ["Spotlight - Angle", 4],
        ["Spotlight - Quantity", 1],
        // Generator appendix (C=74)
        ["Max Recovery", 4],
      ];
      if (effectID < 1 || effectID > 331) {
        return null;
      }
      var id = 1;
      for (var c = 0; c < substatClusters.length; c++) {
        var cluster = substatClusters[c];
        var numRarities = cluster[1];
        if (effectID < id + numRarities) {
          var label = cluster[0];
          return {
            label: label,
            rarity: effectRarities[6 - numRarities + (effectID - id)],
          };
        }
        id += numRarities;
      }
      return null;
    }

    const equippedModulesData = data.moduleEquipped || [];
    const assistSlotData = data.assistModuleSlots || [];
    const inventoryData = data.inventory || [];
    const modulePresetsData = data.modulePresets || [];
    const moduleLevelsData = data.moduleLevels || [];
    const moduleDice = data.moduleDice || null;

    var oldModuleInventory = {};
    var oldModulesPresets = {};
    var oldModulesPlanner = {};
    var moduleInstances = {};

    function collectModuleInstance(module, fallbackCategory) {
      if (!module || !module.guid) {
        return null;
      }
      if (!moduleInstances.hasOwnProperty(module.guid)) {
        const moduleInfo = moduleNames[module.infoIndex];
        const category = (moduleInfo && moduleInfo.category) || fallbackCategory;
        if (!category) {
          return null;
        }
        const name = (moduleInfo && moduleInfo.name) || "Any Other";
        moduleInstances[module.guid] = {
          guid: module.guid,
          name: name,
          category: category.toLowerCase(),
          rarityLevel: module.currentRarity,
          rarity: moduleRarities.hasOwnProperty(module.currentRarity)
            ? moduleRarities[module.currentRarity]
            : "Epic",
          effects: module.effects || [],
        };
      }
      return moduleInstances[module.guid];
    }

    function writeModuleEntry(instance, moduleName) {
      instance.displayName = moduleName;
      var moduleSubstats = [];
      instance.effects.forEach(function (effectID) {
        var substatInfo = lookupEffectID(effectID);
        if (substatInfo) {
          moduleSubstats.push([substatInfo.label, substatInfo.rarity]);
        } else {
          moduleSubstats.push([null, null]);
        }
      });
      if (!oldModuleInventory.hasOwnProperty(instance.category)) {
        oldModuleInventory[instance.category] = {};
      }
      oldModuleInventory[instance.category][moduleName] = {
        rarity: instance.rarity,
        substats: moduleSubstats,
      };

      // if (!oldModulesPlanner.hasOwnProperty(instance.category)) {
      //   oldModulesPlanner[instance.category] = {
      //     "Dice": moduleDice,
      //   };
      // }
    }

    function presetSlot(index, presetName) {
      const moduleCategory = (moduleCategories[index] || "").toLowerCase();
      if (!moduleCategory) {
        return null;
      }
      if (!oldModulesPresets.hasOwnProperty(moduleCategory)) {
        oldModulesPresets[moduleCategory] = {};
      }
      if (!oldModulesPresets[moduleCategory].hasOwnProperty(presetName)) {
        oldModulesPresets[moduleCategory][presetName] = {
          primary: "Any Other",
          secondary: "",
        };
      }
      return oldModulesPresets[moduleCategory][presetName];
    }
    
    function placeModuleInstance(guid) {
      const instance = guid ? moduleInstances[guid] : null;
      if (!instance) {
        return null;
      }
      if (instance.displayName) {
        return instance.displayName;
      }
      const placed = oldModuleInventory[instance.category] || {};
      if (!placed.hasOwnProperty(instance.name)) {
        writeModuleEntry(instance, instance.name);
        return instance.displayName;
      }
      // "Spare" is the marker updateModulesInventory looks for when filling the
      // sheet's spare columns, and there is room for only one of them.
      const spareName = `Spare ${instance.name}`;
      if (!placed.hasOwnProperty(spareName)) {
        writeModuleEntry(instance, spareName);
      } else {
        instance.displayName = spareName;
      }
      return instance.displayName;
    }

    moduleLevelsData.forEach(function (moduleLevel, index) {
      const moduleCategory = moduleCategories[index].toLowerCase();
      if (!moduleCategory) {
        return;
      }
      if (!oldModuleInventory.hasOwnProperty(moduleCategory)) {
        oldModuleInventory[moduleCategory] = {};
      }
      oldModuleInventory[moduleCategory]["Highest Level"] = moduleLevel || null;
    });

    equippedModulesData.forEach(function (module, index) {
      if (!module) {
        return;
      }
      const moduleCategory = moduleCategories[index].toLowerCase();
      if (!moduleCategory) {
        return;
      }
      collectModuleInstance(module, moduleCategory);
      if (!oldModuleInventory.hasOwnProperty(moduleCategory)) {
        oldModuleInventory[moduleCategory] = {};
      }
      if (!oldModulesPresets.hasOwnProperty(moduleCategory)) {
        oldModulesPresets[moduleCategory] = {};
      }
    });

    assistSlotData.forEach(function (assistSlot, index) {
      const assistCategory = moduleCategories[index].toLowerCase();
      if (!assistCategory) {
        return;
      }
      const assistUnlocked = assistSlot.unlocked || false;
      const assistMainEffiency = String(
        assistSlot.mainEffectEfficiencyLevel || 0,
      ).padStart(2, "0");
      const assistSubEffiency = String(
        assistSlot.substatEfficiencyLevel || 0,
      ).padStart(2, "0");
      const assistUniqueEffect = assistSlot.uniqueEffectEfficiencyLevel || 0;
      const assistRarity = assistRarities[assistUniqueEffect] || "Epic";
      if (!oldModulesPresets.hasOwnProperty(assistCategory)) {
        oldModulesPresets[assistCategory] = {};
      }
      oldModulesPresets[assistCategory]["Assist Slot"] = {
        locked: assistUnlocked,
        rarity: assistRarity,
        multiplier: assistMainEffiency,
        substat: assistSubEffiency,
      };

      if (!assistUnlocked) {
        return;
      }
      const equippedAssistModule = assistSlot.equippedModule || {};
      if (!equippedAssistModule) {
        return;
      }
      collectModuleInstance(equippedAssistModule, assistCategory);
      if (!oldModuleInventory.hasOwnProperty(assistCategory)) {
        oldModuleInventory[assistCategory] = {};
      }
      oldModuleInventory[assistCategory]["Assist Level"] =
        assistSlot.level || null;
      if (!oldModulesPresets.hasOwnProperty(assistCategory)) {
        oldModulesPresets[assistCategory] = {};
      }
    });

    inventoryData.forEach(function (module) {
      collectModuleInstance(module);
    });

    var presetNames = [];
    modulePresetsData.forEach(function (preset) {
      var presetName = preset.presetName || null;
      if (!presetName) {
        return;
      }
      if (presetName === "Preset 1") {
        presetName = "Farming";
      }
      presetNames.push(presetName);

      (preset.primaryModuleGuids || []).forEach(function (moduleGuid, index) {
        const slot = presetSlot(index, presetName);
        if (!slot) {
          return;
        }
        slot.primary = placeModuleInstance(moduleGuid) || "Any Other";
      });
      (preset.assistModuleGuids || []).forEach(function (moduleGuid, index) {
        const moduleName = placeModuleInstance(moduleGuid);
        if (!moduleName) {
          return;
        }
        const slot = presetSlot(index, presetName);
        if (slot) {
          slot.secondary = moduleName;
        }
      });
    });

    // Every other module the player owns is represented by its best copy. The
    // rest are "fodder" - duplicates kept only to raise another copy's rarity,
    // an Ancestral module being fed an Epic+ copy of itself to reach
    // Ancestral 1* - and are dropped, along with anything already placed above.
    Object.keys(moduleNames).forEach(function (infoIndex) {
      const moduleInfo = moduleNames[infoIndex];
      const moduleCategory = moduleInfo.category.toLowerCase();
      if (
        oldModuleInventory.hasOwnProperty(moduleCategory) &&
        oldModuleInventory[moduleCategory].hasOwnProperty(moduleInfo.name)
      ) {
        return;
      }
      var bestCopy = null;
      Object.keys(moduleInstances).forEach(function (guid) {
        const instance = moduleInstances[guid];
        if (instance.name !== moduleInfo.name || instance.displayName) {
          return;
        }
        if (
          !bestCopy ||
          (instance.rarityLevel || 0) > (bestCopy.rarityLevel || 0)
        ) {
          bestCopy = instance;
        }
      });
      if (bestCopy) {
        writeModuleEntry(bestCopy, moduleInfo.name);
      }
    });
    oldModulesPresets.presetNames = shared.resolvePresetOrder(
      presetNames,
      shared.templatePresetNames,
    ).order;

    // The order the game lists modules in, for anything that has to display
    // them - the inventory is keyed by name, which says nothing about order.
    const moduleOrder = Object.keys(moduleNames)
      .map(Number)
      .sort(function (a, b) {
        return a - b;
      })
      .map(function (infoIndex) {
        return moduleNames[infoIndex].name;
      });

    return {
      oldModulesInventory: oldModuleInventory,
      oldModulesPresets: oldModulesPresets,
      oldModulesPlanner: oldModulesPlanner,
      moduleOrder: moduleOrder,
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v4.0": this.version4_0.bind(this),
      "v4.7": this.version4_7.bind(this),
      "v5.0": this.version5_0.bind(this),
      "v5.2.1": this.version5_2_1.bind(this),
      // "v6.4.3": this.version6_4_3.bind(this),
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
