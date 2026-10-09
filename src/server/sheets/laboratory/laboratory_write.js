const labWriter = {

  /**
   * Builds the batch update that writes LabLevels into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldLabLevels
   * @param {Object} masterSheetData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateLabLevels: function (sheetName, oldLabLevels, masterSheetData) {
    try {
      console.log("Called: labWriter.updateLabLevels");
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
          var startCol = sheetRefs.columnToLetter(col + 1);
          var endCol = sheetRefs.columnToLetter(col + 2);
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
      var errorReport = errors.report("labWriter.updateLabLevels", error, {
        sheetName: sheetName,
        oldLabLevels: oldLabLevels,
        masterSheetData: masterSheetData,
        batchUpdate: batchUpdate,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes LabPlanner into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldLabPlanner
   * @param {Object} labPlannerData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateLabPlanner: function (sheetName, oldLabPlanner, labPlannerData) {
    try {
      console.log("Called: labWriter.updateLabPlanner");
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
            var boostRange = `${sheetName}!${sheetRefs.columnToLetter(
              firstColIndex + 3,
            )}${rowIndex + 1}`;
            batchUpdate.push({
              range: boostRange,
              values: [[oldBoost]],
            });

            var oldLabData = oldLabPlanner[labHeader]["Labs"];
            if (oldLabData && oldLabData.length !== 0) {
              var startCol = sheetRefs.columnToLetter(firstColIndex);
              var endCol = sheetRefs.columnToLetter(firstColIndex + 2);
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
              var startCol = sheetRefs.columnToLetter(colIndex + 3);
              var endCol = sheetRefs.columnToLetter(colIndex + 4);
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
              var col = sheetRefs.columnToLetter(miscColIndex + 1);
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
              var showLabCol = sheetRefs.columnToLetter(showLabColIndex + 1);
              var optionCol = sheetRefs.columnToLetter(optionColIndex + 1);
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

      var errorReport = errors.report("labWriter.updateLabPlanner", error, {
        sheetName: sheetName,
        oldLabPlanner: oldLabPlanner,
        labPlannerData: labPlannerData,
        rowIndex: rowIndex,
        batchUpdate: batchUpdate,
      });
      return errors.fail(errorReport);
    }
  },
};
