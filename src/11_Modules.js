const modules = {
  // #region Export Functions
  exportData: function (versionDifference) {
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

      var oldDataResult = getVersionFunction();
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
  importData: function (data) {
    try {
      console.log("Called: modules.importData");
      var newSpreadsheet = spreadsheets("Modules newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

      // Batch get required data for update function only
      var requiredRanges = ["Inventory", "Presets", "Tracker", "IDS"];
      var dvtIndex = requiredRanges.length;
      var dvtNamedRanges = {
        "Assist Level": "DVT_Mod_Assist_Level",
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
      var newModulesTrackerValues = batchResults[2].values;
      var idsData = batchResults[3].values;

      var dvtNamedRangesData = {};
      Object.keys(dvtNamedRanges).forEach(function (item) {
        dvtNamedRangesData[item] = batchResults[dvtIndex]
          ? batchResults[dvtIndex].values
          : [];
        dvtIndex++;
      });

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(
        newSheetID,
        "IDS",
        "IDS Master's",
        idsData,
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
      var batchUpdate = [];
      targetModuleTypes.forEach(function (moduleType) {
        if (oldModulesPresets.hasOwnProperty(moduleType)) {
          var rowIdx = newModuleTypeIndex[moduleType] + 1;
          if (typeof rowIdx === "undefined") return;
          var row = newModulePresetsValues[rowIdx];
          var usedPresets = [];
          Object.keys(oldModulesPresets[moduleType]).forEach(
            function (presetName) {
              var presetCol = row.indexOf(presetName);
              if (presetName === "") {
                return;
              }
              var presetData = oldModulesPresets[moduleType][presetName];
              if (!presetData) {
                return;
              }
              if (presetCol === -1) {
                presetCol = row.findIndex(
                  (cell) =>
                    String(cell).toLowerCase().includes("preset") &&
                    !String(cell).toLowerCase().includes("module") &&
                    !usedPresets.includes(cell) &&
                    !oldModulesPresets[moduleType].hasOwnProperty(cell)
                );
                if (presetCol === -1) {
                  return;
                }
              }
              usedPresets.push(row[presetCol]);
              if (presetName === "Assist Slot") {
                var lockedRange = `${sheetName}!${shared.columnToLetter(
                  presetCol + 3,
                )}${rowIdx + 1}:${shared.columnToLetter(presetCol + 3)}${
                  rowIdx + 1
                }`;
                var lockedValues = [
                  [presetData.locked || ""],
                ];
                var rarityRange = `${sheetName}!${shared.columnToLetter(
                  presetCol + 2,
                )}${rowIdx + 2}:${shared.columnToLetter(presetCol + 2)}${
                  rowIdx + 2
                }`;
                var rarityValues = [
                  [presetData.rarity || ""],
                ];
                var multiSubRange = `${sheetName}!${shared.columnToLetter(
                  presetCol + 3,
                )}${rowIdx + 3}:${shared.columnToLetter(presetCol + 3)}${
                  rowIdx + 4
                }`;

                // Transform values using DVT
                var dvtMultiplier = shared.getDVTValue(
                  presetData.multiplier || "",
                  dvtNamedRangesData["Assist Level"],
                );
                var dvtSubstat = shared.getDVTValue(
                  presetData.substat || "",
                  dvtNamedRangesData["Assist Level"],
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
              } else {
                var presetNameRange = `${sheetName}!${shared.columnToLetter(presetCol + 1,)}${rowIdx + 1}`;
                var presetNameValues = [[presetName]];
                batchUpdate.push({
                  range: presetNameRange,
                  values: presetNameValues,
                });

                var presetModuleRange = `${sheetName}!${shared.columnToLetter(
                  presetCol + 2,
                )}${rowIdx + 3}:${shared.columnToLetter(presetCol + 2)}${
                  rowIdx + 4
                }`;
                var presetModuleValues = [
                  [presetData.primary || ""],
                  [presetData.secondary || ""],
                ];
                batchUpdate.push({
                  range: presetModuleRange,
                  values: presetModuleValues,
                });
              }
            },
          );
        }
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
          var maxLevel = oldModulesInventory[moduleType]["Highest Level"] || 0;
          var assistLevel =
            oldModulesInventory[moduleType]["Assist Level"] || 0;
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
              shared.columnToLetter(assistLevelCol + 1) + (rowIdx + 6);
            batchUpdate.push({
              range: `${sheetName}!${assistLevelCell}`,
              values: [[assistLevel]],
            });
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
  version5_2_1: function () {
    try {
      console.log("Called: modules.version5_2_1");
      var oldSpreadsheet = spreadsheets("Modules oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

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

  version5_0: function () {
    try {
      console.log("Called: modules.version5_0");
      var oldSpreadsheet = spreadsheets("Modules oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

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

  version4_7: function () {
    try {
      console.log("Called: modules.version4_7");
      var oldSpreadsheet = spreadsheets("Modules oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

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

  version4_0: function () {
    try {
      console.log("Called: modules.version4_0");
      var oldSpreadsheet = spreadsheets("Modules oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

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
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        if (typeof rowIdx === "undefined") return;
        oldModulesPresets[moduleType] = {};
        var row = oldModulesPresetsValues[rowIdx + 2];
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Primary Slot") {
            var presetName = oldModulesPresetsValues[rowIdx][col]
              ? String(oldModulesPresetsValues[rowIdx][col]).trim()
              : "";
            if (!presetName) {
              continue;
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
          }
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
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        if (typeof rowIdx === "undefined") return;
        oldModulesPresets[moduleType] = {};
        var row = oldModulesPresetsValues[rowIdx];
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Module Name") {
            var presetName = String(
              oldModulesPresetsValues[rowIdx - 1][col],
            ).trim();
            var moduleName = String(
              oldModulesPresetsValues[rowIdx + 1][col],
            ).trim();
            if (presetName && moduleName) {
              oldModulesPresets[moduleType][presetName] = {
                primary: moduleName,
                secondary: "",
              };
            }
          }
        }
      });

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
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v4.0": this.version4_0.bind(this),
      "v4.7": this.version4_7.bind(this),
      "v5.0": this.version5_0.bind(this),
      "v5.2.1": this.version5_2_1.bind(this),
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
