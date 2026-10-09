const modulesTrackerReader = {

  /**
   * Extracts ModulesTracker from a v6.4.3 sheet's values.
   * @param {Array<Array<*>>} oldModulesTrackerValues
   * @param {Array<Array<string>>} oldModulesTrackerFormulas
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion6_4_3ModulesTracker: function (
    oldModulesTrackerValues,
    oldModulesTrackerFormulas,
  ) {
    try {
      console.log("Called: modulesTrackerReader.getVersion6_4_3ModulesTracker");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModulesTracker = {};

      for (var row = 0; row < oldModulesTrackerValues.length; row++) {
        var rowData = oldModulesTrackerValues[row];
        var inputColIdx = rowData.indexOf("Input values");
        if (inputColIdx !== -1) {
          oldModulesTracker["Input values"] = {};
          for (
            var inputRow = row + 2;
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
                  var rowIdx = row + 3;
                  rowIdx < oldModulesTrackerFormulas.length;
                  rowIdx++
                ) {
                  var summaryRowData = oldModulesTrackerFormulas[rowIdx];
                  var summaryModuleName = summaryRowData[idx] || null;
                  var summaryModuleCount = summaryRowData[idx + 1] || null;
                  if (!summaryModuleName) {
                    continue;
                  }
                  if (summaryModuleName.toLowerCase().includes("total")) {
                    break;
                  }
                  if (!String(summaryModuleCount).trim().startsWith("=")) {
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
        const subRowStart = row + 4;
        for (var col = colIdx; col < rowData.length; col++) {
          var moduleName = String(rowData[col]).trim();
          if (
            moduleName &&
            oldModulesTrackerValues[subRowStart][col] === "Copies"
          ) {
            if (!oldModulesTracker[targetModule]) {
              oldModulesTracker[targetModule] = {};
            }
            oldModulesTracker[targetModule][moduleName] = [];
            for (var subRow = subRowStart; subRow < subRowStart + 4; subRow++) {
              var subRowData = oldModulesTrackerValues[subRow];

              var copy = subRowData[col + 1] || null;
              oldModulesTracker[targetModule][moduleName].push(copy);
            }

            if (oldModulesTrackerValues[subRowStart + 5][col] === "Shattered epic") {
              oldModulesTracker[targetModule][moduleName + " Shattered"] =
                oldModulesTrackerValues[subRowStart + 5][col + 2] || null;
            }
          } else if (moduleName === "Fodders") {
            if (!oldModulesTracker[targetModule]) {
              oldModulesTracker[targetModule] = {};
            }
            const fodderRowStart = row + 3;
            oldModulesTracker[targetModule][moduleName] = [];
            for (var subRow = fodderRowStart; subRow < fodderRowStart + 6; subRow++) {
              var subRowData = oldModulesTrackerValues[subRow];

              var fodder = subRowData[col + 2] || null;
              oldModulesTracker[targetModule][moduleName].push(fodder);
            }
            break;
          }
        }
      }
      return {
        success: true,
        oldModulesTracker: oldModulesTracker,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion6_4_3ModulesTracker", error, {
        oldModulesTrackerValues: oldModulesTrackerValues,
        oldModulesTrackerFormulas: oldModulesTrackerFormulas,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts ModulesTracker from a v4.7 sheet's values.
   * @param {Array<Array<*>>} oldModulesTrackerValues
   * @param {Array<Array<string>>} oldModulesTrackerFormulas
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_7ModulesTracker: function (
    oldModulesTrackerValues,
    oldModulesTrackerFormulas,
  ) {
    try {
      console.log("Called: modulesTrackerReader.getVersion4_7ModulesTracker");
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

      }
      return {
        success: true,
        oldModulesTracker: oldModulesTracker,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion4_7ModulesTracker", error, {
        oldModulesTrackerValues: oldModulesTrackerValues,
        oldModulesTrackerFormulas: oldModulesTrackerFormulas,
      });
      return errors.fail(errorReport);
    }
  },
};
