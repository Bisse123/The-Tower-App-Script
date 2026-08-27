const lab = {
  // #region Export Functions
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: lab.exportData");
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
        message: "Laboratory export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting lab data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data, newSheetID) {
    try {
      console.log("Called: lab.importData");
      var newSpreadsheet = spreadsheets("Laboratory newSpreadsheet", newSheetID);
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }

      var requiredRanges = ["Master Sheet", "IDS"];
      var labPlannerSheetName = "";
      var batchUpdate = [];

      var labPlannerSheet = SheetsAPI.getSheetBySubstring(
        newSpreadsheet,
        "Lab Planner",
      );
      if (labPlannerSheet) {
        labPlannerSheetName = labPlannerSheet.title;
        requiredRanges.push(labPlannerSheetName);
      }

      // Batch get required data for update function only
      var batchResults = SheetsAPI.batchGetFormulas(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var idsData = batchResults[1].values;
      var labPlannerData = batchResults[2] ? batchResults[2].values : null;

      // Only update lab levels if key exists
      if (data.hasOwnProperty("oldLabLevels")) {
        var oldLabLevels = data.oldLabLevels;
        var labResult = this.updateLabLevels(
          "Master Sheet",
          oldLabLevels,
          masterSheetData,
        );
        if (!labResult || !labResult.success) {
          console.log(`Error updating lab levels: ${labResult.message}`);
          return labResult;
        }
        batchUpdate = batchUpdate.concat(labResult.batchUpdate || []);
      }

      // Only update lab planner if key exists
      if (data.hasOwnProperty("oldLabPlanner")) {
        var oldLabPlanner = data.oldLabPlanner;
        var labPlannerResult = this.updateLabPlanner(
          labPlannerSheetName,
          oldLabPlanner,
          labPlannerData,
        );
        if (!labPlannerResult || !labPlannerResult.success) {
          console.log(
            `Error updating lab planner: ${labPlannerResult.message}`,
          );
          return labPlannerResult;
        }
        batchUpdate = batchUpdate.concat(labPlannerResult.batchUpdate || []);
      }

      // Set sheet IDs and IDS Master ID (moved from copyFileTemplate for optimization)
      batchUpdate = shared.addIDUpdatesToBatch(
        batchUpdate,
        "Laboratory",
        newSheetID,
        idsData,
        data.idMasterID,
      );

      // Apply all updates
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
        message: `Laboratory import completed successfully`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: "Error importing lab data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Update Functions
  updateLabLevels: function (sheetName, oldLabLevels, masterSheetData) {
    try {
      console.log("Called: lab.updateLabLevels");
      var headerValues = ["Labs"];

      if (!masterSheetData || masterSheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: "Not enough data in Master Sheet",
        };
      }

      var headerRow = masterSheetData[0];
      var lastRow = masterSheetData.length;

      var columnsToCheck = [];
      for (var i = 0; i < headerRow.length; i++) {
        if (headerValues.includes(headerRow[i])) {
          columnsToCheck.push(i + 1);
        }
      }

      if (columnsToCheck.length === 0) {
        console.log(`No Labs columns found in Master Sheet`);
        return {
          success: false,
          message: "No Labs columns found in Master Sheet",
        };
      }

      var batchUpdate = [];
      columnsToCheck.forEach(function (col) {
        var newLabLevels = [];
        var numRows = lastRow - 2;

        for (var row = 1; row < numRows + 1; row++) {
          if (row >= masterSheetData.length) break;

          var cellValue = masterSheetData[row][col - 1];
          if (!cellValue || cellValue.trim() === "") break;

          if (oldLabLevels.hasOwnProperty(cellValue)) {
            var oldLabLevel = oldLabLevels[cellValue];
            newLabLevels.push(oldLabLevel);
          } else {
            var currentLevel = masterSheetData[row][col] || 0;
            var currentTarget = masterSheetData[row][col + 1] || "";
            newLabLevels.push([currentLevel, currentTarget]);
          }
        }
        if (newLabLevels.length > 0) {
          var startCol = shared.columnToLetter(col + 1);
          var endCol = shared.columnToLetter(col + 2);
          var range = `${sheetName}!${startCol}2:${endCol}${
            newLabLevels.length + 1
          }`;

          batchUpdate.push({
            range: range,
            values: newLabLevels,
          });
        }
      });

      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: "Lab levels updated successfully",
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: "No updates needed for lab levels",
      };
    } catch (error) {
      console.log(`Error in updateLabLevels: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating lab levels: ${error.message}`,
      };
    }
  },

  updateLabPlanner: function (sheetName, oldLabPlanner, labPlannerData) {
    try {
      console.log("Called: lab.updateLabPlanner");
      if (!labPlannerData || labPlannerData.length === 0) {
        console.log(`No lab planner data provided`);
        return {
          success: true,
          message: "No lab planner data provided",
        };
      }
      if (!oldLabPlanner || Object.keys(oldLabPlanner).length === 0) {
        console.log(`No lab planner updates provided`);
        return {
          success: true,
          message: "No lab planner updates needed",
        };
      }

      var labHeaders = [
        "Lab One",
        "Lab Two",
        "Lab Three",
        "Lab Four",
        "Lab Five",
      ];
      var reminderHeaders = [
        "Lab One Reminder",
        "Lab Two Reminder",
        "Lab Three Reminder",
        "Lab Four Reminder",
        "Lab Five Reminder",
      ];
      var miscHeaders = [
        "OPTIONS",
        "Estimated Daily Coins required to Sustain:",
      ];

      var batchUpdate = [];

      var estimatedCoinsHeader = [...labHeaders];
      for (var rowIndex = 0; rowIndex < labPlannerData.length; rowIndex++) {
        var row = labPlannerData[rowIndex];
        if (
          labHeaders.length === 0 &&
          reminderHeaders.length === 0 &&
          miscHeaders.length === 0
        ) {
          break;
        }
        labHeaders = labHeaders.filter(function (labHeader) {
          var colIndex = row.findIndex(function (cellValue) {
            return (
              cellValue &&
              typeof cellValue === "string" &&
              cellValue.startsWith("=") &&
              cellValue.includes(labHeader)
            );
          });

          if (colIndex !== -1) {
            var firstColIndex = colIndex + row[colIndex].split(",").length;
            var oldBoost = oldLabPlanner[labHeader]["Boost"];
            var boostRange = `${sheetName}!${shared.columnToLetter(
              firstColIndex + 3,
            )}${rowIndex + 1}`;
            batchUpdate.push({
              range: boostRange,
              values: [[oldBoost]],
            });

            var oldLabData = oldLabPlanner[labHeader]["Labs"];
            if (oldLabData && oldLabData.length !== 0) {
              var startCol = shared.columnToLetter(firstColIndex);
              var endCol = shared.columnToLetter(firstColIndex + 2);
              var startRow = rowIndex + 4;
              var endRow = startRow + oldLabData.length - 1;
              var labRange = `${sheetName}!${startCol}${startRow}:${endCol}${endRow}`;
              var labValues = oldLabData.map(function (dataRow) {
                return [dataRow[0] || "", dataRow[1] || "", dataRow[2] || ""];
              });
              batchUpdate.push({
                range: labRange,
                values: labValues,
              });
            }
            return false;
          }
          return true;
        });
        reminderHeaders = reminderHeaders.filter(function (reminderHeader) {
          var colIndex = row.findIndex(function (cellValue) {
            return (
              cellValue &&
              typeof cellValue === "string" &&
              cellValue.trim().toLowerCase() === reminderHeader.toLowerCase()
            );
          });
          if (colIndex !== -1) {
            var oldReminderData = oldLabPlanner[reminderHeader];
            if (oldReminderData && oldReminderData.length !== 0) {
              var startCol = shared.columnToLetter(colIndex + 3);
              var endCol = shared.columnToLetter(colIndex + 4);
              var startRow = rowIndex + 1;
              var endRow = startRow + oldReminderData.length - 1;
              var range = `${sheetName}!${startCol}${startRow}:${endCol}${endRow}`;
              batchUpdate.push({
                range: range,
                values: oldReminderData,
              });
            }
            return false;
          }
          return true;
        });
        miscHeaders = miscHeaders.filter(function (miscHeader) {
          var miscColIndex = row.findIndex(function (cellValue) {
            return (
              cellValue &&
              typeof cellValue === "string" &&
              cellValue.trim().toLowerCase() === miscHeader.toLowerCase()
            );
          });
          if (miscColIndex !== -1) {
            var miscData = oldLabPlanner[miscHeader];
            if (
              miscHeader === "Estimated Daily Coins required to Sustain:" &&
              miscData &&
              miscData.length !== 0
            ) {
              var labStartOption =
                oldLabPlanner["OPTIONS"]["I plan my labs starting at the: →"];
              if (labStartOption && labStartOption.length > 1) {
                for (var index = 0; index < miscData.length; index++) {
                  var dataRow = miscData[index];
                  if (dataRow && dataRow.length > 0) {
                    // var headers = Object.keys(oldLabPlanner);
                    var oldLabHeader =
                      oldLabPlanner[estimatedCoinsHeader[index]];
                    if (!oldLabHeader) {
                      console.log(
                        `No old lab header found for ${estimatedCoinsHeader[index]}`,
                      );
                      continue;
                    }
                    var oldLabData = oldLabHeader["Labs"];
                    if (!oldLabData || oldLabData.length === 0) {
                      console.log(
                        `No old lab data found for ${estimatedCoinsHeader[index]}`,
                      );
                      continue;
                    }
                    var oldLabDataFiltered = oldLabData.filter(
                      function (dataRow) {
                        return (
                          dataRow &&
                          dataRow.length > 2 &&
                          dataRow[2] &&
                          dataRow[2].trim() !== ""
                        );
                      },
                    );
                    var miscIndex =
                      labStartOption[1] === "Top"
                        ? 0
                        : oldLabDataFiltered.length - 1;
                    var labLevel = oldLabDataFiltered[miscIndex]
                      ? oldLabDataFiltered[miscIndex][2]
                      : null;
                    if (
                      labLevel &&
                      labLevel !== "" &&
                      labLevel === dataRow[0]
                    ) {
                      dataRow[0] = null;
                    }
                  }
                }
              }
              var col = shared.columnToLetter(miscColIndex + 1);
              var startCell = `${col}${rowIndex + 2}`;
              var endCell = `${col}${rowIndex + 6}`;
              var range = `${sheetName}!${startCell}:${endCell}`;
              batchUpdate.push({
                range: range,
                values: miscData,
              });
            } else if (
              miscHeader === "OPTIONS" &&
              miscData &&
              Object.keys(miscData).length !== 0
            ) {
              var plannerType = row[miscColIndex + 1] !== "" ? 1 : 2;
              var plannerRows =
                (labPlannerData[rowIndex + 1][plannerType] !== "" ? 1 : 2) * 4;
              var showLabColIndex = miscColIndex + 4 * plannerType - 2;
              var optionColIndex = miscColIndex + 5 * plannerType - 2;
              var showLabCol = shared.columnToLetter(showLabColIndex + 1);
              var optionCol = shared.columnToLetter(optionColIndex + 1);
              var startCell = `${showLabCol}${rowIndex + 1}`;
              var endCell = `${optionCol}${rowIndex + plannerRows}`;
              var range = `${sheetName}!${startCell}:${endCell}`;
              var values = [];
              for (var i = 0; i < plannerRows; i++) {
                var currentRowIndex = rowIndex + i;
                var optionKey =
                  labPlannerData[currentRowIndex][miscColIndex + plannerType] ||
                  "";
                if (optionKey.startsWith("=")) {
                  optionKey = optionKey
                    .split(",")
                    .pop()
                    .trim()
                    .replace(/['"]/g, "")
                    .replace(/[')]/g, "");
                }
                if (optionKey && miscData[optionKey]) {
                  if (plannerType === 1) {
                    values.push([
                      miscData[optionKey][0] || "",
                      miscData[optionKey][1] || "",
                    ]);
                  } else {
                    values.push([
                      miscData[optionKey][0] || "",
                      "",
                      miscData[optionKey][1] || "",
                    ]);
                  }
                } else {
                  if (plannerType === 1) {
                    values.push([
                      labPlannerData[currentRowIndex][showLabColIndex] || "",
                      labPlannerData[currentRowIndex][optionColIndex] || "",
                    ]);
                  } else {
                    values.push([
                      labPlannerData[currentRowIndex][showLabColIndex] || "",
                      "",
                      labPlannerData[currentRowIndex][optionColIndex] || "",
                    ]);
                  }
                }
              }

              batchUpdate.push({
                range: range,
                values: values,
              });
            }
            return false;
          }
          return true;
        });
      }

      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: `Lab planner updated successfully (${batchUpdate.length} cells updated)`,
          batchUpdate: batchUpdate,
        };
      }

      return {
        success: true,
        message: "No lab planner formulas found to update",
      };
    } catch (error) {
      console.log(`Error in updateLabPlanner: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating lab planner: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: lab.version1_0");
      var oldSpreadsheet = spreadsheets("Laboratory oldSpreadsheet", oldSheetID);

      var labLevelsRange = "EXPORT!B5:E";
      var rangesToFetch = [labLevelsRange];

      var oldLabPlannerSheet = SheetsAPI.getSheetBySubstring(
        oldSpreadsheet,
        "Lab Planner",
      );

      var oldLabPlannerValues = null;
      var oldLabPlannerFormulas = null;
      if (oldLabPlannerSheet) {
        var oldLabPlannerSheetName = oldLabPlannerSheet.title;
        if (oldLabPlannerSheetName) {
          rangesToFetch.push(oldLabPlannerSheetName);
          var oldLabPlannerData = SheetsAPI.batchGetFormulas(oldSheetID, [
            oldLabPlannerSheetName,
          ]);
          if (
            oldLabPlannerData &&
            oldLabPlannerData.length > 0 &&
            oldLabPlannerData[0].values
          ) {
            oldLabPlannerFormulas = oldLabPlannerData[0].values;
          }
        }
      }

      var labBatchResult = SheetsAPI.batchGetValues(oldSheetID, rangesToFetch);
      if (
        !labBatchResult ||
        labBatchResult.length === 0 ||
        !labBatchResult[0].values
      ) {
        console.log(`Could not read lab levels data`);
        return {
          success: false,
          message: "Could not read lab levels data",
        };
      }
      var oldLabLevelsValues = labBatchResult[0].values;

      if (labBatchResult[1] && labBatchResult[1].values) {
        oldLabPlannerValues = labBatchResult[1].values;
      }

      // Process lab levels first
      var labLevelsResult = this.getVersion1_0LabLevels(oldLabLevelsValues);
      if (!labLevelsResult || !labLevelsResult.success) {
        return labLevelsResult;
      }

      var oldLabLevels = labLevelsResult.oldLabLevels;
      var oldLabMax = labLevelsResult.oldLabMax;

      // Process lab planner if data exists
      var labPlannerResult = this.getVersion1_0LabPlanner(
        oldLabPlannerValues,
        oldLabPlannerFormulas,
        oldLabLevels,
        oldLabMax,
      );
      if (!labPlannerResult || !labPlannerResult.success) {
        return labPlannerResult;
      }

      if (!labPlannerResult.oldLabPlanner) {
        console.log(`No lab planner data found in old spreadsheet`);
        return {
          success: true,
          message: "No sheet containing 'Lab Planner' found in old spreadsheet",
          oldLabLevels: oldLabLevels,
        };
      }
      var oldLabPlanner = labPlannerResult.oldLabPlanner;

      return {
        success: true,
        message: "Laboratory processed successfully",
        oldLabLevels: oldLabLevels,
        oldLabPlanner: oldLabPlanner,
      };
    } catch (error) {
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Lab Levels
  getVersion1_0LabLevels: function (oldLabLevelsValues) {
    try {
      console.log("Called: lab.getVersion1_0LabLevels");
      var oldLabLevels = {};
      var oldLabMax = {};
      oldLabLevelsValues.forEach(function (row) {
        var hasData = row.some(function (cell) {
          return (
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
          );
        });

        if (hasData && row[0]) {
          oldLabLevels[row[0]] = [row[1] || 0, row[2] || ""];
          oldLabMax[row[0]] = row[3] || null;
        }
      });

      return {
        success: true,
        message: "Lab levels processed successfully",
        oldLabLevels: oldLabLevels,
        oldLabMax: oldLabMax,
      };
    } catch (error) {
      console.log("Error in getVersion1_0LabLevels: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0LabLevels: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Lab Planner
  getVersion1_0LabPlanner: function (
    oldLabPlannerValues,
    oldLabPlannerFormulas,
    oldLabLevels,
    oldLabMax,
  ) {
    try {
      console.log("Called: lab.getVersion1_0LabPlanner");
      if (!oldLabPlannerFormulas || !oldLabPlannerValues) {
        console.log(
          `No sheet containing "Lab Planner" found in old spreadsheet`,
        );
        return {
          success: true,
          message: "No sheet containing 'Lab Planner' found in old spreadsheet",
        };
      }

      var labHeaders = [
        "Lab One",
        "Lab Two",
        "Lab Three",
        "Lab Four",
        "Lab Five",
      ];
      var reminderHeaders = [
        "Lab One Reminder",
        "Lab Two Reminder",
        "Lab Three Reminder",
        "Lab Four Reminder",
        "Lab Five Reminder",
      ];
      var miscHeaders = [
        "OPTIONS",
        "Estimated Daily Coins required to Sustain:",
      ];

      var oldLabPlanner = {};
      for (
        var rowIndex = 0;
        rowIndex < oldLabPlannerFormulas.length;
        rowIndex++
      ) {
        var row = oldLabPlannerFormulas[rowIndex];
        if (
          labHeaders.length === 0 &&
          reminderHeaders.length === 0 &&
          miscHeaders.length === 0
        ) {
          break;
        }
        labHeaders = labHeaders.filter(function (labHeader) {
          var colIndex = row.findIndex(function (cellValue) {
            return (
              cellValue &&
              typeof cellValue === "string" &&
              cellValue.startsWith("=") &&
              cellValue.includes(labHeader)
            );
          });
          if (colIndex !== -1) {
            var firstColIndex = colIndex + row[colIndex].split(",").length - 1;
            if (!oldLabPlanner[labHeader]) {
              oldLabPlanner[labHeader] = {};
            }
            if (!oldLabPlanner[labHeader]["Labs"]) {
              oldLabPlanner[labHeader]["Labs"] = [];
            }

            oldLabPlanner[labHeader]["Boost"] =
              oldLabPlannerValues[rowIndex][firstColIndex + 3] || "";

            var lastNonEmptyRow = -1;
            for (var i = rowIndex + 3; i < oldLabPlannerFormulas.length; i++) {
              if (
                !oldLabPlannerFormulas[i][colIndex] ||
                oldLabPlannerFormulas[i][colIndex].trim() === ""
              ) {
                break;
              }

              var labName = oldLabPlannerValues[i][firstColIndex + 2] || "";
              if (labName.trim() === "") {
                oldLabPlanner[labHeader]["Labs"].push(["", "", ""]);
                continue;
              }
              lastNonEmptyRow = i - (rowIndex + 3);
              var plannerLevel = oldLabPlannerValues[i][firstColIndex] || "";
              if (
                oldLabLevels[labName] &&
                plannerLevel === oldLabLevels[labName][0]
              ) {
                plannerLevel = "";
              }
              var plannerTarget =
                oldLabPlannerValues[i][firstColIndex + 1] || "";
              if (
                oldLabLevels[labName] &&
                (plannerTarget === oldLabLevels[labName][1] ||
                  plannerTarget === oldLabMax[labName])
              ) {
                plannerTarget = "";
              }

              oldLabPlanner[labHeader]["Labs"].push([
                plannerLevel,
                plannerTarget,
                labName,
              ]);
            }
            if (lastNonEmptyRow === -1) {
              delete oldLabPlanner[labHeader]["Labs"];
            } else {
              oldLabPlanner[labHeader]["Labs"] = oldLabPlanner[labHeader][
                "Labs"
              ].slice(0, lastNonEmptyRow + 1);
            }
            return false;
          }
          return true;
        });
        reminderHeaders = reminderHeaders.filter(function (reminderHeader) {
          var colIndex = row.findIndex(function (cellValue) {
            return (
              cellValue &&
              typeof cellValue === "string" &&
              cellValue.trim().toLowerCase() === reminderHeader.toLowerCase()
            );
          });
          if (colIndex !== -1) {
            var reminderRowIndex = rowIndex;
            if (!oldLabPlanner[reminderHeader]) {
              oldLabPlanner[reminderHeader] = [];
            }
            while (
              oldLabPlannerValues[reminderRowIndex][colIndex] === reminderHeader
            ) {
              var reminderData = oldLabPlannerValues[reminderRowIndex];
              oldLabPlanner[reminderHeader].push([
                reminderData[colIndex + 2] || "",
                reminderData[colIndex + 3] || "",
              ]);
              reminderRowIndex++;
            }
            return false;
          }
          return true;
        });
        miscHeaders = miscHeaders.filter(function (miscHeader) {
          var miscColIndex = row.findIndex(function (cellValue) {
            return (
              cellValue &&
              typeof cellValue === "string" &&
              cellValue.trim().toLowerCase() === miscHeader.toLowerCase()
            );
          });
          if (miscColIndex !== -1) {
            if (miscHeader === "Estimated Daily Coins required to Sustain:") {
              oldLabPlanner[miscHeader] = oldLabPlannerValues
                .slice(rowIndex + 1, rowIndex + 6)
                .map(function (row) {
                  return [row[miscColIndex] || null];
                });
            } else if (miscHeader === "OPTIONS") {
              var plannerType = row[miscColIndex + 1] !== "" ? 1 : 2;
              var plannerRows =
                (oldLabPlannerFormulas[rowIndex + 1][
                  miscColIndex + plannerType
                ] !== ""
                  ? 1
                  : 2) * 4;
              var showLabColIndex = miscColIndex + (4 * plannerType - 2);
              var optionColIndex = miscColIndex + (5 * plannerType - 2);
              var optionDict = {};
              oldLabPlannerFormulas
                .slice(rowIndex, rowIndex + plannerRows)
                .forEach(function (row) {
                  if (row[miscColIndex + 1] !== "") {
                    var optionKey = row[miscColIndex + 1];
                    if (optionKey.startsWith("=")) {
                      optionKey = optionKey
                        .split(",")
                        .pop()
                        .trim()
                        .replace(/['"]/g, "")
                        .replace(/[')]/g, "");
                    }
                    optionDict[optionKey] = [
                      row[showLabColIndex] || "",
                      row[optionColIndex] || "",
                    ];
                  }
                });
              oldLabPlanner[miscHeader] = optionDict;
            }
            return false;
          }
          return true;
        });
      }

      return {
        success: true,
        message: "Lab planner processed successfully",
        oldLabPlanner: oldLabPlanner,
      };
    } catch (error) {
      console.log("Error in getVersion1_0LabPlanner: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0LabPlanner: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseLabData: function (data) {
    const labNamesByIndex = {
      0: "Damage",
      1: "Attack Speed",
      2: "Critical Factor",
      3: "Range",
      4: "Damage / Meter",
      5: "Super Crit Chance",
      6: "Super Crit Multi",
      10: "Health",
      11: "Health Regen",
      12: "Defense Absolute",
      13: "Defense %",
      14: "Orbs Speed",
      15: "Land Mine Damage",
      16: "Land Mine Decay",
      17: "Shockwave Size",
      18: "Orb Boss Hit",
      19: "Recovery Package Amount",
      20: "Cash Bonus",
      21: "Cash / Wave",
      22: "Coins / Kill Bonus",
      23: "Coins / Wave",
      24: "Interest",
      25: "Max Interest",
      26: "Package After Boss",
      30: "Game Speed",
      31: "Starting Cash",
      32: "Workshop Attack Discount",
      33: "Workshop Defense Discount",
      34: "Workshop Utility Discount",
      35: "Labs Coin Discount",
      36: "Labs Speed",
      37: "Buy Multiplier",
      38: "More Round Stats",
      39: "Target Priority",
      40: "Presets",
      41: "Workshop Respec",
      50: "Missile Despawn Time",
      51: "Missiles Explosion",
      52: "Missile Radius",
      53: "Chrono Field Duration",
      54: "Chrono Field Damage Reduction",
      55: "Chrono Field Reduction %",
      56: "Swamp Radius",
      57: "Swamp Stun",
      58: "Swamp Stun Chance",
      59: "Swamp Stun Time",
      60: "Golden Tower Bonus",
      61: "Golden Tower Duration",
      62: "Chain Lightning Shock",
      63: "Shock Chance",
      64: "Shock Multiplier",
      65: "Death Wave Health",
      66: "Death Wave Coin Bonus",
      67: "Inner Mine Blast Radius",
      68: "Inner Mine Rotation Speed",
      69: "Chrono Field Range",
      70: "Second Wind Blast",
      71: "Double Death Ray",
      72: "Extra Orb Adjuster",
      73: "Extra Extra Orbs",
      74: "Energy Shield Extra Hit",
      75: "Super Tower Bonus",
      80: "Unlock Perks",
      81: "Waves Required",
      82: "Auto Pick Perks",
      83: "Standard Perks Bonus",
      84: "Perk Option Quantity",
      85: "First Perk Choice",
      86: "First Trade-off Choice",
      87: "Ban Perks",
      88: "Improve Trade-off Perks",
      90: "Missile Amplifier",
      91: "Missile Barrage",
      92: "Missile Barrage Quantity",
      93: "Inner Mine Stun",
      94: "Black Hole Damage",
      95: "Extra Black Hole",
      96: "Black Hole Coin Bonus",
      97: "Spotlight Coin Bonus",
      98: "Spotlight Missiles",
      99: "Black Hole ignore Protector",
      100: "Recovery Package Max",
      101: "Recovery Package Chance",
      102: "Flame Bot - Cooldown",
      103: "Thunder Bot - Cooldown",
      104: "Gold Bot - Cooldown",
      105: "Amp Bot - Cooldown",
      106: "Flame Bot - Burn Stack",
      107: "Thunder Bot - Linger Time",
      108: "Gold Bot - Duration",
      109: "Amp Bot - Duration",
      110: "Common Enemy Health",
      111: "Common Enemy Attack",
      112: "Fast Enemy Health",
      113: "Fast Enemy Attack",
      114: "Fast Enemy Speed",
      115: "Tank Enemy Health",
      116: "Tank Enemy Attack",
      117: "Ranged Enemy Health",
      118: "Ranged Enemy Attack",
      119: "Boss Health",
      120: "Boss Attack",
      121: "Protector Health",
      122: "Protector Radius",
      123: "Protector Damage Reduction",
      124: "Enemy Attack Level Skip",
      125: "Enemy Health Level Skip",
      126: "Wall Health",
      127: "Wall Rebuild",
      128: "Wall Regen",
      129: "Wall Thorns",
      130: "Wall Invincibility",
      131: "Max Rend Armor Multiplier",
      132: "Light Speed Shots",
      133: "Black Hole Disable Ranged Enemies",
      134: "Common Drop Chance",
      139: "Reroll Shards",
      140: "Daily Mission Shards",
      141: "Module Shards Cost",
      142: "Module Coin Cost",
      143: "Rare Drop Chance",
      144: "Wall Fortification",
      145: "Recharge Second Wind",
      146: "Recharge Demon Mode",
      147: "Recharge Missile Barrage",
      148: "Reroll Daily Mission",
      149: "Recharge Nuke",
      150: "Workshop Enhancements",
      151: "Unmerge Module",
      152: "Shatter Shards",
      153: "Auto Pick Ranking",
      154: "Enhancement Attack - Coin Discount",
      155: "Enhancement Defense - Coin Discount",
      156: "Swamp Rend",
      157: "Swamp Rend - Additional Enemies",
      158: "Chain Thunder",
      159: "Lightning Amplifier - Scatter",
      160: "Damage Mastery",
      161: "Attack Speed Mastery",
      162: "Health Mastery",
      163: "Health Regen Mastery",
      164: "Range Mastery",
      165: "Cash Mastery",
      166: "Coins Mastery",
      167: "Slow Aura Mastery",
      168: "Critical Chance Mastery",
      169: "Enemy Balance Mastery",
      170: "Extra Defense Mastery",
      171: "Fortress Mastery",
      172: "Free Upgrades Mastery",
      173: "Extra Orb Mastery",
      174: "Plasma Cannon Mastery",
      175: "Critical Coin Mastery",
      176: "Wave Skip Mastery",
      177: "Intro Sprint Mastery",
      178: "Land Mine Stun Mastery",
      179: "Recovery Package Chance Mastery",
      180: "Death Ray Mastery",
      181: "Energy Net Mastery",
      182: "Super Tower Mastery",
      183: "Second Wind Mastery",
      184: "Demon Mode Mastery",
      185: "Energy Shield Mastery",
      186: "Wave Accelerator Mastery",
      187: "Berserker Mastery",
      188: "Ultimate Crit Mastery",
      189: "Nuke Mastery",
      190: "Death Wave Cells Bonus",
      191: "Death Wave Damage Amplifier",
      192: "Death Wave Armor Stripping",
      193: "Garlic Thorns",
      194: "Cannon Effect Bans",
      195: "Armor Effect Bans",
      196: "Generator Effect Bans",
      197: "Core Effect Bans",
      198: "Inner Land Mine - Chrono Jump",
      199: "Battle Condition Reduction",
      200: "Area of Effect Mastery",
      201: "Knockback Resistance",
      202: "Thorns Resistance",
      203: "Orb Resistance",
      204: "Plasma Cannon Resistance",
      205: "Death Ray Resistance",
      206: "Ultimate Weapon Durations",
      207: "Death Defy Down",
      208: "Energy Shields Down",
      209: "Enemy Level Skip Reduction",
      210: "Fast's Ultimate",
      211: "Ranged Ultimate",
      212: "Boss's Ultimate",
      213: "Basic's Ultimate",
      214: "Tank's Ultimate",
      215: "Protector's Ultimate",
      216: "Armored Enemies",
      217: "Enemy Speed",
      218: "More Enemies",
      219: "Enemy Attack Speed",
      220: "Ray Enemy Attack",
      221: "Ray Enemy Health",
      222: "Vampire Enemy Attack",
      223: "Vampire Enemy Health",
      224: "Scatter Enemy Attack",
      225: "Scatter Enemy Health",
      226: "Ranged Enemy Range",
      227: "Enhancement Utility - Coin Discount",
      228: "Bot Bot - Cooldown",
      229: "Bot Bot - Duration",
      230: "Assist Module Substats - Cannon",
      231: "Assist Module Substats - Armor",
      232: "Assist Module Substats - Generator",
      233: "Assist Module Substats - Core",
      234: "Assist Module Bonus - Cannon",
      235: "Assist Module Bonus - Armor",
      236: "Assist Module Bonus - Generator",
      237: "Assist Module Bonus - Core",
      238: "Dissonant Echo - Utility",
      239: "Dissonant Echo - Attack",
      240: "Dissonant Echo - Defense",
      241: "Dissonant Echo - Ultimate Weapons",
      242: "Overcharge Enemy Health",
      243: "Overcharge Enemy Damage",
      244: "Commander Enemy Health",
      245: "Saboteur Enemy Health",
      252: "Global Presets",
    };
    // Missing labs:
    // Cells Mastery
    // Overcharge Exponent Reducer
    // Commander Radius
    // Saboteur Attack Speed
    
    const labLevels = data.researchLevel || [];
    var oldLabLevels = {};
    var labOrder = [];
    labLevels.forEach(function (labLevel, index) {
      var labName = labNamesByIndex[index];
      if (!labName && !labLevel) {
        return;
      }
      if (!labName) {
        console.log(`No lab name found for index ${index}. lab level ${labLevel}`);
        labName = `Unknown Lab ${index}`;
      }
      oldLabLevels[labName] = [labLevel, null];
      labOrder[index] = labName;
    });

    return {
      oldLabLevels: oldLabLevels,
      labOrder: labOrder,
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
    };
  },

  // #endregion
  // #region Compatibility Check
  isCompatibleVersion: function (oldVersion) {
    console.log("Called: lab.isCompatibleVersion");
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
