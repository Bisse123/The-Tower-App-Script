const modulesWriter = {

  /**
   * Builds the batch update that writes ModulesPresets into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldModulesPresets
   * @param {Array<Array<*>>} newModulePresetsValues
   * @param {Object} dvtNamedRangesData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateModulesPresets: function (
    sheetName,
    oldModulesPresets,
    newModulePresetsValues,
    dvtNamedRangesData,
  ) {
    try {
      console.log("Called: modulesWriter.updateModulesPresets");
      if (!newModulePresetsValues) {
        console.log(`Could not read modules presets data`);
        return {
          success: false,
          message: "Could not read modules presets data",
        };
      }
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var newModuleTypeIndex = modulesLayout.findModuleTypesRowIndex(
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
            var presetNameRange = `${sheetName}!${sheetRefs.columnToLetter(presetCol + 1)}${rowIdx + 1}`;
            var presetNameValues = [[presetName]];
            batchUpdate.push({
              range: presetNameRange,
              values: presetNameValues,
            });
          }

          var presetModuleRange = `${sheetName}!${sheetRefs.columnToLetter(
            presetCol + 2,
          )}${rowIdx + 3}:${sheetRefs.columnToLetter(presetCol + 2)}${rowIdx + 4}`;
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
        var lockedRange = `${sheetName}!${sheetRefs.columnToLetter(
          presetCol + 3,
        )}${rowIdx + 1}:${sheetRefs.columnToLetter(presetCol + 3)}${rowIdx + 1}`;
        var lockedValues = [[presetData.locked || false]];
        var rarityRange = `${sheetName}!${sheetRefs.columnToLetter(
          presetCol + 2,
        )}${rowIdx + 2}:${sheetRefs.columnToLetter(presetCol + 2)}${rowIdx + 2}`;
        var rarityValues = [[presetData.rarity || null]];
        var multiSubRange = `${sheetName}!${sheetRefs.columnToLetter(
          presetCol + 3,
        )}${rowIdx + 3}:${sheetRefs.columnToLetter(presetCol + 3)}${rowIdx + 4}`;

        var dvtMultiplier = dropdownUtils.getDVTValue(
          presetData.multiplier || null,
          dvtNamedRangesData["Main Efficiency"],
        );
        var dvtSubstat = dropdownUtils.getDVTValue(
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
      var errorReport = errors.report("dvtNamedRanges.updateModulesPresets", error, {
        sheetName: sheetName,
        oldModulesPresets: oldModulesPresets,
        newModulePresetsValues: newModulePresetsValues,
        dvtNamedRangesData: dvtNamedRangesData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes ModulesInventory into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldModulesInventory
   * @param {Array<Array<*>>} newModuleInventoryValues
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateModulesInventory: function (
    sheetName,
    oldModulesInventory,
    newModuleInventoryValues,
  ) {
    try {
      console.log("Called: modulesWriter.updateModulesInventory");
      if (!newModuleInventoryValues) {
        console.log("Could not read modules inventory data");
        return {
          success: false,
          message: "Could not read modules inventory data",
        };
      }
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var newModuleTypeIndex = modulesLayout.findModuleTypesRowIndex(
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
                    range: `${sheetName}!${sheetRefs.columnToLetter(col + 1)}${
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
              var rarityCell = sheetRefs.columnToLetter(col + 1) + (rowIdx + 3);
              batchUpdate.push({
                range: `${sheetName}!${rarityCell}`,
                values: [[oldModulesInventory[moduleType][cellValue].rarity]],
              });
              var substats =
                oldModulesInventory[moduleType][cellValue].substats;
              if (substats && substats.length > 0) {
                var numRows = substats.length;
                var numCols = substats[0].length;
                var startCell = sheetRefs.columnToLetter(col + 1) + (rowIdx + 5);
                var endCell =
                  sheetRefs.columnToLetter(col + numCols) +
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
                sheetRefs.columnToLetter(highestLevelCol + 1) + (rowIdx + 3);
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
                sheetRefs.columnToLetter(assistLevelCol + 1) + (rowIdx + 6);
              batchUpdate.push({
                range: `${sheetName}!${assistLevelCell}`,
                values: [[assistLevel]],
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
      var errorReport = errors.report("dvtNamedRanges.updateModulesInventory", error, {
        sheetName: sheetName,
        oldModulesInventory: oldModulesInventory,
        newModuleInventoryValues: newModuleInventoryValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes ModulesTracker into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldModulesTracker
   * @param {Array<Array<*>>} newModulesTrackerValues
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateModulesTracker: function (
    sheetName,
    oldModulesTracker,
    newModulesTrackerValues,
  ) {
    try {
      console.log("Called: modulesWriter.updateModulesTracker");
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
            var inputRow = row + 2;
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
            var inputColLetter = sheetRefs.columnToLetter(inputColIdx + 5);
            var inputRowStart = row + 3;
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
                  var rowIdx = row + 3;
                  rowIdx < newModulesTrackerValues.length;
                  rowIdx++
                ) {
                  var summaryRowData = newModulesTrackerValues[rowIdx];
                  var summaryModuleName = summaryRowData[idx] || null;
                  if (!summaryModuleName) {
                    continue;
                  }
                  if (summaryModuleName.toLowerCase().includes("total")) {
                    break;
                  }
                  if (summaryData.hasOwnProperty(summaryModuleName)) {
                    var summaryModuleCount = summaryData[summaryModuleName];
                    batchUpdate.push({
                      range: `${sheetName}!${sheetRefs.columnToLetter(idx + 2)}${rowIdx + 1}`,
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
            oldModulesTracker[targetModule] &&
            oldModulesTracker[targetModule].hasOwnProperty(moduleName)
          ) {
            var startRow = row + 5;
            var rangeCol = col + 2;
            if (moduleName === "Fodders") {
              startRow -= 1;
              rangeCol += 1;
            }
            var endRow =
              startRow + oldModulesTracker[targetModule][moduleName].length - 1;
            var range = `${sheetName}!${sheetRefs.columnToLetter(
              rangeCol,
            )}${startRow}:${sheetRefs.columnToLetter(rangeCol)}${endRow}`;
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
              var shatteredRow = startRow + 5;
              var shatteredCol = col + 3;
              var shatteredRange = `${sheetName}!${sheetRefs.columnToLetter(shatteredCol)}${shatteredRow}`;
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
      var errorReport = errors.report("dvtNamedRanges.updateModulesTracker", error, {
        sheetName: sheetName,
        oldModulesTracker: oldModulesTracker,
        newModulesTrackerValues: newModulesTrackerValues,
      });
      return errors.fail(errorReport);
    }
  },
};
