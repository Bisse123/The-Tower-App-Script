const lab = {
  // #region Export Functions
  exportData: function (versionDifference) {
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

      var oldDataResult = getVersionFunction();
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
  importData: function (data) {
    try {
      console.log("Called: lab.importData");
      var newSpreadsheet = spreadsheets("Laboratory newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
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

      // Add import status update to batch only if there were data updates
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }

      // Set sheet IDs and IDS Master ID (moved from copyFileTemplate for optimization)
      batchUpdate = shared.addIDUpdatesToBatch(
        batchUpdate,
        "Laboratory",
        newSheetID,
        idsData,
        data.idMasterID,
      );

      // Apply all updates (always includes ID setting, conditionally includes import status)
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

          var oldLabLevel = oldLabLevels[cellValue];
          if (oldLabLevel && oldLabLevel.length >= 2) {
            newLabLevels.push([oldLabLevel[0] || 0, oldLabLevel[1] || ""]);
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
  version1_0: function () {
    try {
      console.log("Called: lab.version1_0");
      var oldSpreadsheet = spreadsheets("Laboratory oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old lab spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet™ not found in old lab spreadsheet™",
        };
      }

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
    const labNameIndices = [
      "Damage",                                // 0
      "Attack Speed",                          // 1
      "Critical Factor",                       // 2
      "Range",                                 // 3
      "Damage / Meter",                        // 4
      "Super Crit Chance",                     // 5
      "Super Crit Multi",                      // 6
      null,                                    // 7
      null,                                    // 8
      null,                                    // 9
      "Health",                                // 10
      "Health Regen",                          // 11
      "Defense Absolute",                      // 12
      "Defense %",                             // 13
      "Orbs Speed",                            // 14
      "Land Mine Damage",                      // 15
      "Land Mine Decay",                       // 16
      "Shockwave Size",                        // 17
      "Orb Boss Hit",                          // 18
      "Recovery Package Amount",               // 19
      "Cash Bonus",                            // 20
      "Cash / Wave",                           // 21
      "Coins / Kill Bonus",                    // 22
      "Coins / Wave",                          // 23
      "Interest",                              // 24
      "Max Interest",                          // 25
      "Package After Boss",                    // 26
      null,                                    // 27
      null,                                    // 28
      null,                                    // 29
      "Game Speed",                            // 30
      "Starting Cash",                         // 31
      "Workshop Attack Discount",              // 32
      "Workshop Defense Discount",             // 33
      "Workshop Utility Discount",             // 34
      "Labs Coin Discount",                    // 35
      "Labs Speed",                            // 36
      "Buy Multiplier",                        // 37
      "More Round Stats",                      // 38
      "Target Priority",                       // 39
      "Card Presets",                          // 40
      "Workshop Respec",                       // 41
      null,                                    // 42
      null,                                    // 43
      null,                                    // 44
      null,                                    // 45
      null,                                    // 46
      null,                                    // 47
      null,                                    // 48
      null,                                    // 49
      "Missile Despawn Time",                  // 50
      "Missiles Explosion",                    // 51
      "Missile Radius",                        // 52
      "Chrono Field Duration",                 // 53
      "Chrono Field Damage Reduction",         // 54
      "Chrono Field Reduction %",              // 55
      "Swamp Radius",                          // 56
      "Swamp Stun",                            // 57
      "Swamp Stun Chance",                     // 58
      "Swamp Stun Time",                       // 59
      "Golden Tower Bonus",                    // 60
      "Golden Tower Duration",                 // 61
      "Chain Lightning Shock",                 // 62
      "Shock Chance",                          // 63
      "Shock Multiplier",                      // 64
      "Death Wave Health",                     // 65
      "Death Wave Coin Bonus",                 // 66
      "Inner Mine Blast Radius",               // 67
      "Inner Mine Rotation Speed",             // 68
      "Chrono Field Range",                    // 69
      "Second Wind Blast",                     // 70
      "Double Death Ray",                      // 71
      "Extra Orb Adjuster",                    // 72
      "Extra Extra Orbs",                      // 73
      "Energy Shield Extra Hit",               // 74
      "Super Tower Bonus",                     // 75
      null,                                    // 76
      null,                                    // 77
      null,                                    // 78
      null,                                    // 79
      "Unlock Perks",                          // 80
      "Waves Required",                        // 81
      "Auto Pick Perks",                       // 82
      "Standard Perks Bonus",                  // 83
      "Perk Option Quantity",                  // 84
      "First Perk Choice",                     // 85
      "First Trade-off Choice",                // 86
      "Ban Perks",                             // 87
      "Improve Trade-off Perks",               // 88
      null,                                    // 89
      "Missile Amplifier",                     // 90
      "Missile Barrage",                       // 91
      "Missile Barrage Quantity",              // 92
      "Inner Mine Stun",                       // 93
      "Black Hole Damage",                     // 94
      "Extra Black Hole",                      // 95
      "Black Hole Coin Bonus",                 // 96
      "Spotlight Coin Bonus",                  // 97
      "Spotlight Missiles",                    // 98
      "Black Hole ignore Protector",           // 99
      "Recovery Package Max",                  // 100
      "Recovery Package Chance",               // 101
      "Flame Bot - Cooldown",                  // 102
      "Thunder Bot - Cooldown",                // 103
      "Gold Bot - Cooldown",                   // 104
      "Amp Bot - Cooldown",                    // 105
      "Flame Bot - Burn Stack",                // 106
      "Thunder Bot - Linger Time",             // 107
      "Gold Bot - Duration",                   // 108
      "Amp Bot - Duration",                    // 109
      "Common Enemy Health",                   // 110
      "Common Enemy Attack",                   // 111
      "Fast Enemy Health",                     // 112
      "Fast Enemy Attack",                     // 113
      "Fast Enemy Speed",                      // 114
      "Tank Enemy Health",                     // 115
      "Tank Enemy Attack",                     // 116
      "Ranged Enemy Health",                   // 117
      "Ranged Enemy Attack",                   // 118
      "Boss Health",                           // 119
      "Boss Attack",                           // 120
      "Protector Health",                      // 121
      "Protector Radius",                      // 122
      "Protector Damage Reduction",            // 123
      "Enemy Attack Level Skip",               // 124
      "Enemy Health Level Skip",               // 125
      "Wall Health",                           // 126
      "Wall Rebuild",                          // 127
      "Wall Regen",                            // 128
      "Wall Thorns",                           // 129
      "Wall Invincibility",                    // 130
      "Max Rend Armor Multiplier",             // 131
      "Light Speed Shots",                     // 132
      "Black Hole Disable Ranged Enemies",     // 133
      "Common Drop Chance",                    // 134
      null,                                    // 135
      null,                                    // 136
      null,                                    // 137
      null,                                    // 138
      "Reroll Shards",                         // 139
      "Daily Mission Shards",                  // 140
      "Module Shards Cost",                    // 141
      "Module Coin Cost",                      // 142
      "Rare Drop Chance",                      // 143
      "Wall Fortification",                    // 144
      "Recharge Second Wind",                  // 145
      "Recharge Demon Mode",                   // 146
      "Recharge Missile Barrage",              // 147
      "Reroll Daily Mission",                  // 148
      "Recharge Nuke",                         // 149
      "Workshop Enhancements",                 // 150
      "Unmerge Module",                        // 151
      "Shatter Shards",                        // 152
      "Auto Pick Ranking",                     // 153
      "Enhancement Attack - Coin Discount",    // 154
      "Enhancement Defense - Coin Discount",   // 155
      "Swamp Rend",                            // 156
      "Swamp Rend - Additional Enemies",       // 157
      "Chain Thunder",                         // 158
      "Lightning Amplifier - Scatter",         // 159
      "Damage Mastery",                        // 160
      "Attack Speed Mastery",                  // 161
      "Health Mastery",                        // 162
      "Health Regen Mastery",                  // 163
      "Range Mastery",                         // 164
      "Cash Mastery",                          // 165
      "Coins Mastery",                         // 166
      "Slow Aura Mastery",                     // 167
      "Critical Chance Mastery",               // 168
      "Enemy Balance Mastery",                 // 169
      "Extra Defense Mastery",                 // 170
      "Fortress Mastery",                      // 171
      "Free Upgrades Mastery",                 // 172
      "Extra Orb Mastery",                     // 173
      "Plasma Cannon Mastery",                 // 174
      "Critical Coin Mastery",                 // 175
      "Wave Skip Mastery",                     // 176
      "Intro Sprint Mastery",                  // 177
      "Land Mine Stun Mastery",                // 178
      "Recovery Package Chance Mastery",       // 179
      "Death Ray Mastery",                     // 180
      "Energy Net Mastery",                    // 181
      "Super Tower Mastery",                   // 182
      "Second Wind Mastery",                   // 183
      "Demon Mode Mastery",                    // 184
      "Energy Shield Mastery",                 // 185
      "Wave Accelerator Mastery",              // 186
      "Berserker Mastery",                     // 187
      "Ultimate Crit Mastery",                 // 188
      "Nuke Mastery",                          // 189
      "Death Wave Cells Bonus",                // 190
      "Death Wave Damage Amplifier",           // 191
      "Death Wave Armor Stripping",            // 192
      "Garlic Thorns",                         // 193
      "Cannon Effect Bans",                    // 194
      "Armor Effect Bans",                     // 195
      "Generator Effect Bans",                 // 196
      "Core Effect Bans",                      // 197
      "Inner Land Mine - Chrono Jump",         // 198
      "Battle Condition Reduction",            // 199
      "Area of Effect Mastery",                // 200
      "Knockback Resistance",                  // 201
      "Thorns Resistance",                     // 202
      "Orb Resistance",                        // 203
      "Plasma Cannon Resistance",              // 204
      "Death Ray Resistance",                  // 205
      "Ultimate Weapon Durations",             // 206
      "Death Defy Down",                       // 207
      "Energy Shields Down",                   // 208
      "Enemy Level Skip Reduction",            // 209
      "Fast's Ultimate",                       // 210
      "Ranged Ultimate",                       // 211
      "Boss's Ultimate",                       // 212
      "Basic's Ultimate",                      // 213
      "Tank's Ultimate",                       // 214
      "Protector's Ultimate",                  // 215
      "Armored Enemies",                       // 216
      "Enemy Speed",                           // 217
      "More Enemies",                          // 218
      "Enemy Attack Speed",                    // 219
      "Ray Enemy Attack",                      // 220
      "Ray Enemy Health",                      // 221
      "Vampire Enemy Attack",                  // 222
      "Vampire Enemy Health",                  // 223
      "Scatter Enemy Attack",                  // 224
      "Scatter Enemy Health",                  // 225
      "Ranged Enemy Range",                    // 226
      "Enhancement Utility - Coin Discount",   // 227
      "Bot Bot - Cooldown",                    // 228
      "Bot Bot - Duration",                    // 229
      "Assist Module Substats - Cannon",       // 230
      "Assist Module Substats - Armor",        // 231
      "Assist Module Substats - Generator",    // 232
      "Assist Module Substats - Core",         // 233
      "Assist Module Bonus - Cannon",          // 234
      "Assist Module Bonus - Armor",           // 235
      "Assist Module Bonus - Generator",       // 236
      "Assist Module Bonus - Core",            // 237
      "Dissonant Echo - Utility",              // 238
      "Dissonant Echo - Attack",               // 239
      "Dissonant Echo - Defense",              // 240
      "Dissonant Echo - Ultimate Weapons",     // 241
      "Overcharge Enemy Health",               // 242
      "Overcharge Enemy Damage",               // 243
      "Commander Enemy Health",                // 244
      "Saboteur Enemy Health",                 // 245
      null,                                    // 246
      null,                                    // 247
      null,                                    // 248
      null,                                    // 249
    ];

    const labLevels = data.researchLevel || [];
    var oldLabLevels = {};
    labNameIndices.forEach(function (labName, index) {
      if (labName) {
        oldLabLevels[labName] = [labLevels[index] || 0, null];
      }
    });
    
    return {
      oldLabLevels: oldLabLevels,
      labNameIndices: labNameIndices,
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
