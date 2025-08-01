const lab = {
  importData: function (versionDifference) {
    try {
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

      var getVersionFunction = this.convertVersionFunctions[versionDifference];
      if (!getVersionFunction) {
        console.log(`Unsupported version difference: ${versionDifference}`);
        return {
          success: false,
          message: `Unsupported version difference: ${versionDifference}`,
        };
      }
      var oldDataResult = getVersionFunction();
      if (!oldDataResult || !oldDataResult.success) {
        console.log(`${oldDataResult.message}`);
        return oldDataResult;
      }

      var oldLabLevels = oldDataResult.oldLabLevels || [];
      var requiredRanges = ["Master Sheet"];
      var labPlannerSheetName = "";

      if (oldDataResult.oldLabPlanner) {
        var labPlannerSheet = SheetsAPI.getSheetBySubstring(newSpreadsheet, "Lab Planner");
        if (labPlannerSheet) {
          labPlannerSheetName = labPlannerSheet.title;
          requiredRanges.push(labPlannerSheetName);
        }
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
      var labPlannerData = batchResults[1].values;

      var labResult = this.updateLabLevels(
        "Master Sheet",
        oldLabLevels,
        masterSheetData
      );
      if (!labResult || !labResult.success) {
        console.log(`Error updating lab levels: ${labResult.message}`);
        return labResult;
      }

      var batchUpdate = labResult.batchUpdate || [];

      if (oldDataResult.oldLabPlanner && Object.keys(oldDataResult.oldLabPlanner).length !== 0 && labPlannerSheetName) {
        console.log(oldDataResult.oldLabPlanner);
        var labPlannerResult = this.updateLabPlanner(
          labPlannerSheetName,
          oldDataResult.oldLabPlanner,
          labPlannerData
        );
        if (!labPlannerResult || !labPlannerResult.success) {
          console.log(`Error updating lab planner: ${labPlannerResult.message}`);
          return labPlannerResult;
        }
        batchUpdate = batchUpdate.concat(labPlannerResult.batchUpdate);
      }
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
        return {
          success: true,
          message: labResult.message,
        };
      }

      return {
        success: true,
        message: `No updates needed for Laboratory`,
      };
    } catch (error) {
      console.log(`Error in importLabData: ${error.toString()}`);
      return {
        success: false,
        message: "Error importing lab data: " + error.message,
      };
    }
  },

  updateLabLevels: function (sheetName, oldLabLevels, masterSheetData) {
    try {
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

      // Find columns where header is in headerValues
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

      // oldLabLevels is now a dictionary, no need for separate updateMap
      var batchUpdate = [];
      // Iterate each "Labs" column
      columnsToCheck.forEach(function (col) {
        var newLabLevels = [];
        // Find labNames in each column (skip header row, ignore last row with sums)
        var numRows = lastRow - 2;

        for (var row = 1; row < numRows + 1; row++) {
          if (row >= masterSheetData.length) break;

          var cellValue = masterSheetData[row][col - 1];
          if (!cellValue || cellValue.trim() === "") break;

          var oldLabLevel = oldLabLevels[cellValue];
          if (oldLabLevel) {
            newLabLevels.push([oldLabLevel[0] || 0, oldLabLevel[1] || ""]);
          } else {
            // Keep existing values
            var currentLevel = masterSheetData[row][col] || 0;
            var currentTarget = masterSheetData[row][col + 1] || "";
            newLabLevels.push([currentLevel, currentTarget]);
          }
        }
        // Add batch update for this column's Level and Target columns
        if (newLabLevels.length > 0) {
          var startCol = shared.columnToLetter(col + 1); // Level column
          var endCol = shared.columnToLetter(col + 2); // Target column
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

      var batchUpdate = [];
      var labHeaders = Object.keys(oldLabPlanner).filter(function (header) {
        return header && header.trim() !== "" && header.toLowerCase().includes("lab") && !header.toLowerCase().includes("reminder")
      });
      var reminderHeaders = Object.keys(oldLabPlanner).filter(function (header) {
        return header && header.trim() !== "" && header.toLowerCase().includes("lab") && header.toLowerCase().includes("reminder")
      });
      var miscHeaders = Object.keys(oldLabPlanner).filter(function (header) {
        return header && header.trim() !== "" && !header.toLowerCase().includes("lab")
      });

      for (var rowIndex = 0; rowIndex < labPlannerData.length; rowIndex++) {
        var row = labPlannerData[rowIndex];
        if (labHeaders.length === 0 && reminderHeaders.length === 0 && miscHeaders.length === 0) {
          break;
        }
        labHeaders = labHeaders.filter(function(labHeader) {
          var colIndex = row.findIndex(function (cellValue) {
            return cellValue && 
                   typeof cellValue === 'string' &&
                   cellValue.startsWith('=') &&
                   cellValue.includes(labHeader);
          });
          
          if (colIndex !== -1) {
            var firstColIndex = colIndex + row[colIndex].split(",").length;
            var oldBoost = oldLabPlanner[labHeader]["Boost"];
            var boostRange = `${sheetName}!${shared.columnToLetter(firstColIndex + 3)}${rowIndex + 1}`;
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
        reminderHeaders = reminderHeaders.filter(function(reminderHeader) {
          var colIndex = row.findIndex(function (cellValue) {
            return cellValue && 
                   typeof cellValue === 'string' &&
                   cellValue.trim().toLowerCase() === reminderHeader.toLowerCase();
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
        miscHeaders = miscHeaders.filter(function(miscHeader) {
          var miscColIndex = row.findIndex(function (cellValue) {
            return cellValue && 
                   typeof cellValue === 'string' &&
                   cellValue.trim().toLowerCase() === miscHeader.toLowerCase();
          });
          if (miscColIndex !== -1) {
            var miscData = oldLabPlanner[miscHeader];
            if (miscHeader === "Estimated Daily Coins required to Sustain:" && miscData && miscData.length !== 0) {
              var col = shared.columnToLetter(miscColIndex + 1);
              var startCell = `${col}${rowIndex + 2}`;
              var endCell = `${col}${rowIndex + 6}`;
              var range = `${sheetName}!${startCell}:${endCell}`;
              batchUpdate.push({
                range: range,
                values: miscData,
              });
            } else if (miscHeader === "OPTIONS" && miscData && Object.keys(miscData).length !== 0) {
              var plannerType = row[miscColIndex + 1] !== "" ? 1 : 2;
              var plannerRows = (labPlannerData[rowIndex + 1][plannerType] !== "" ? 1 : 2) * 4;
              var showLabColIndex = miscColIndex + (4 * plannerType) - 2;
              var optionColIndex = miscColIndex + (5 * plannerType) - 2;
              var showLabCol = shared.columnToLetter(showLabColIndex + 1);
              var optionCol = shared.columnToLetter(optionColIndex + 1);
              var startCell = `${showLabCol}${rowIndex + 1}`;
              var endCell = `${optionCol}${rowIndex + plannerRows}`;
              var range = `${sheetName}!${startCell}:${endCell}`;
              var values = [];
              for (var i = 0; i < plannerRows; i++) {
                var currentRowIndex = rowIndex + i;
                var optionKey = labPlannerData[currentRowIndex][miscColIndex + plannerType] || "";
                if (optionKey.startsWith("=")) {
                  optionKey = optionKey.split(",").pop().trim().replace(/['"]/g, "").replace(/[')]/g, "");
                }
                if (optionKey && miscData[optionKey]) {
                  if (plannerType === 1) {
                    values.push([miscData[optionKey][0] || "", miscData[optionKey][1] || ""]);
                  } else {
                    values.push([miscData[optionKey][0] || "", "", miscData[optionKey][1] || ""]);
                  }
                } else {
                  if (plannerType === 1) {
                    values.push([labPlannerData[currentRowIndex][showLabColIndex] || "", labPlannerData[currentRowIndex][optionColIndex] || ""]);
                  } else {
                    values.push([labPlannerData[currentRowIndex][showLabColIndex] || "", "", labPlannerData[currentRowIndex][optionColIndex] || ""]);
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

  version10: function () {
    try {
      var oldSpreadsheet = spreadsheets("oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old lab spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet™ not found in old lab spreadsheet™",
        };
      }
      
      var labLevelsRange = "EXPORT!B5:E"

      var labBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        labLevelsRange,
      ]);
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

      var oldLabLevels = {};
      var oldLabMax = {};
      oldLabLevelsValues.forEach(function (row) {
        var hasData = row.some(function (cell) {
          return cell !== null &&
                 cell !== undefined &&
                 String(cell || "").trim() !== "";
        });
        
        if (hasData && row[0]) {
          oldLabLevels[row[0]] = [row[1] || 0, row[2] || ""];
          oldLabMax[row[0]] = row[3] || "";
        }
      });

      var oldLabPlannerSheet = SheetsAPI.getSheetBySubstring(oldSpreadsheet, "Lab Planner");
      if (!oldLabPlannerSheet) {
        console.log(`No sheet containing "Lab Planner" found in old spreadsheet`);
        return {
          success: true,
          message: "No sheet containing 'Lab Planner' found in old spreadsheet",
          oldLabLevels: oldLabLevels,
        };
      }
      var oldLabPlannerSheetName = oldLabPlannerSheet.title;

      var oldLabPlannerData = SheetsAPI.batchGetFormulas(oldSheetID, [
        oldLabPlannerSheetName
      ]);
      if (!oldLabPlannerData || oldLabPlannerData.length === 0 || !oldLabPlannerData[0].values) {
        console.log(`Could not read old lab planner data`);
        return {
          success: true,
          message: "Could not read old lab planner data",
          oldLabLevels: oldLabLevels,
        };
      }

      var oldLabPlannerValues = oldLabPlannerData[0].values;
      var labHeaders = ["Lab One", "Lab Two", "Lab Three", "Lab Four", "Lab Five"];
      var reminderHeaders = ["Lab One Reminder", "Lab Two Reminder", "Lab Three Reminder", "Lab Four Reminder", "Lab Five Reminder"];
      var miscHeaders = ["OPTIONS", "Estimated Daily Coins required to Sustain:"];
      
      var oldLabPlanner = {};
      for (var rowIndex = 0; rowIndex < oldLabPlannerValues.length; rowIndex++) {
        var row = oldLabPlannerValues[rowIndex];
        if (labHeaders.length === 0 && reminderHeaders.length === 0 && miscHeaders.length === 0) {
          break;
        }
        labHeaders = labHeaders.filter(function(labHeader) {
          var colIndex = row.findIndex(function (cellValue) {
            return cellValue && 
                   typeof cellValue === 'string' && 
                   cellValue.startsWith('=') && 
                   cellValue.includes(labHeader);
          });
          if (colIndex !== -1) {
            var firstColIndex = colIndex + row[colIndex].split(",").length - 1;
            if (!oldLabPlanner[labHeader]) {
              oldLabPlanner[labHeader] = {};
            }
            if (!oldLabPlanner[labHeader]["Labs"]) {
              oldLabPlanner[labHeader]["Labs"] = [];
            }
            
            oldLabPlanner[labHeader]["Boost"] = oldLabPlannerValues[rowIndex][firstColIndex + 3] || "";

            var lastNonEmptyRow = -1;
            for (var i = rowIndex + 3; i < oldLabPlannerValues.length; i++) {
              if (oldLabPlannerValues[i][colIndex].trim() === "") {
                break;
              }
              
              var labName = oldLabPlannerValues[i][firstColIndex + 2] || "";
              if (labName.trim() === "") {
                oldLabPlanner[labHeader]["Labs"].push([
                "", "", "",
                ]);
                continue;
              }
              lastNonEmptyRow = i - (rowIndex + 3);
              var plannerLevel = oldLabPlannerValues[i][firstColIndex] || "";
              if (plannerLevel === oldLabLevels[labName][0]) {
                plannerLevel = "";
              }
              var plannerTarget = oldLabPlannerValues[i][firstColIndex + 1] || "";
              if (plannerTarget === oldLabLevels[labName][1] || plannerTarget === oldLabMax[labName]) {
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
              oldLabPlanner[labHeader]["Labs"] = oldLabPlanner[labHeader]["Labs"].slice(0, lastNonEmptyRow + 1);
            }
            return false;
          }
          return true;
        });
        reminderHeaders = reminderHeaders.filter(function(reminderHeader) {
          var colIndex = row.findIndex(function (cellValue) {
            return cellValue && 
                   typeof cellValue === 'string' && 
                   cellValue.trim().toLowerCase() === reminderHeader.toLowerCase();
          });
          if (colIndex !== -1) {
            console.log(`Found reminder header: ${reminderHeader} at column index ${colIndex}`);
            var reminderRowIndex = rowIndex;
            if (!oldLabPlanner[reminderHeader]) {
              oldLabPlanner[reminderHeader] = [];
            }
            while (oldLabPlannerValues[reminderRowIndex][colIndex] === reminderHeader) {
              var reminderData = oldLabPlannerValues[reminderRowIndex];
              oldLabPlanner[reminderHeader].push([reminderData[colIndex + 2] || "", reminderData[colIndex + 3] || ""]);
              reminderRowIndex++;
            }
            return false;
          }
          return true;
        });
        miscHeaders = miscHeaders.filter(function(miscHeader) {
          var miscColIndex = row.findIndex(function (cellValue) {
            return cellValue && 
                   typeof cellValue === 'string' && 
                   cellValue.trim().toLowerCase() === miscHeader.toLowerCase();
          });
          if (miscColIndex !== -1) {
            if (miscHeader === "Estimated Daily Coins required to Sustain:") {
              oldLabPlanner[miscHeader] = oldLabPlannerValues.slice(rowIndex + 1, rowIndex + 6).map(function (row) {
                return [row[miscColIndex] || ""];
              });
            } else if (miscHeader === "OPTIONS") {
              var plannerType = row[miscColIndex + 1] !== "" ? 1 : 2;
              var plannerRows =  (oldLabPlannerValues[rowIndex + 1][miscColIndex + plannerType] !== "" ? 1 : 2) * 4;
              var showLabColIndex = miscColIndex + (4 * plannerType - 2);
              var optionColIndex = miscColIndex + (5 * plannerType - 2);
              var optionDict = {};
              oldLabPlannerValues.slice(rowIndex, rowIndex + plannerRows).forEach(function (row) {
                if (row[miscColIndex + 1] !== "") {
                  var optionKey = row[miscColIndex + 1];
                  if (optionKey.startsWith("=")) {
                    optionKey = optionKey.split(",").pop().trim().replace(/['"]/g, "").replace(/[')]/g, "");
                  }
                  optionDict[optionKey] = [row[showLabColIndex] || "", row[optionColIndex] || ""];
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
        message: "Laboratory processed successfully",
        oldLabLevels: oldLabLevels,
        oldLabPlanner: oldLabPlanner,
      };
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
    };
  },

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
};
