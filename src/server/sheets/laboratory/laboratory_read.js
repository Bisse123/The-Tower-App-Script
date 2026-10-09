const labReader = {

  /**
   * Reads Laboratory data from a v1.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: labReader.version1_0");
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

      var labLevelsResult = this.getVersion1_0LabLevels(oldLabLevelsValues);
      if (!labLevelsResult || !labLevelsResult.success) {
        return labLevelsResult;
      }

      var oldLabLevels = labLevelsResult.oldLabLevels;
      var oldLabMax = labLevelsResult.oldLabMax;

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
      var errorReport = errors.report("labReader.version1_0", error, {
        oldSheetID: oldSheetID,
        oldLabLevelsValues: oldLabLevelsValues,
        oldLabPlannerValues: oldLabPlannerValues,
        oldLabPlannerFormulas: oldLabPlannerFormulas,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts LabLevels from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldLabLevelsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0LabLevels: function (oldLabLevelsValues) {
    try {
      console.log("Called: labReader.getVersion1_0LabLevels");
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
      var errorReport = errors.report("labReader.getVersion1_0LabLevels", error, {
        oldLabLevelsValues: oldLabLevelsValues,
        oldLabLevels: oldLabLevels,
        oldLabMax: oldLabMax,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts LabPlanner from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldLabPlannerValues
   * @param {Array<Array<string>>} oldLabPlannerFormulas
   * @param {Array<Array<*>>} oldLabLevels
   * @param {Array<Array<*>>} oldLabMax
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0LabPlanner: function (
    oldLabPlannerValues,
    oldLabPlannerFormulas,
    oldLabLevels,
    oldLabMax,
  ) {
    try {
      console.log("Called: labReader.getVersion1_0LabPlanner");
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
      var errorReport = errors.report("labReader.getVersion1_0LabPlanner", error, {
        oldLabPlannerFormulas: oldLabPlannerFormulas,
        oldLabPlannerValues: oldLabPlannerValues,
        rowIndex: rowIndex,
        oldLabPlanner: oldLabPlanner,
      });
      return errors.fail(errorReport);
    }
  },
};
