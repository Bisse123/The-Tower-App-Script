const relics = {
  convertVersionFunctions: {},

  importData: function (versionDifference) {
    function importRelicsData(versionDifference) {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!newSpreadsheet) {
          console.log("New spreadsheet not found");
          return {
            success: false,
            message: "New spreadsheet not found",
          };
        }
        var newSheetID = newSpreadsheet.spreadsheetId;

        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        if (!oldSpreadsheet) {
          console.log("Old spreadsheet not found");
          return {
            success: false,
            message: "Old spreadsheet not found",
          };
        }
        var oldSheetID = oldSpreadsheet.spreadsheetId;
        
        if (versionDifference === 0) {
          console.log("Same Version");

          // Check if Relics sheet exists in old spreadsheet
          if (!SheetsAPI.getSheetByName(oldSheetID, "Relics")) {
            console.log("Relics sheet not found in old relic spreadsheet");
            return {
              success: false,
              message: "Relics sheet not found in old relic spreadsheet"
            };
          }

          // Get all data from old Relics sheet
          var oldRelicsData = SheetsAPI.getDataRange(
            oldSheetID,
            "Relics"
          );
          if (!oldRelicsData || oldRelicsData.length === 0) {
            console.log("Could not read data from old Relics sheet");
            return {
              success: false,
              message: "Could not read data from old Relics sheet"
            };
          }

          var oldRelicHeaderRow = null;
          var oldRelicNameCol = null;
          var oldRelicUnlockedCol = null;

          // Scan each row to find the header
          for (var row = 0; row < oldRelicsData.length; row++) {
            var rowValues = oldRelicsData[row];
            var relicNameIndex = rowValues.indexOf("Relic Name");
            var relicUnlockedIndex = rowValues.indexOf("Unlocked");
            if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
              oldRelicHeaderRow = row + 1; // Convert to 1-based
              oldRelicNameCol = relicNameIndex + 1; // Convert to 1-based
              oldRelicUnlockedCol = relicUnlockedIndex + 1; // Convert to 1-based
              break;
            }
          }
          if (oldRelicHeaderRow) {
            var startRow = oldRelicHeaderRow + 1;
            var oldRelicsRange =
              "Relics!" +
              shared.columnToLetter(oldRelicNameCol) +
              startRow + ":" +
              shared.columnToLetter(oldRelicUnlockedCol);

            var oldRelicsValues = SheetsAPI.getValues(
              oldSheetID,
              oldRelicsRange
            );
            if (!oldRelicsValues) {
              console.log("Could not read relic data from old spreadsheet");
              return {
                success: false,
                message: "Could not read relic data from old spreadsheet"
              };
            }

            // Filter out empty rows
            var oldRelics = oldRelicsValues.filter((row) =>
              row.some(
                (cell) =>
                  cell !== null &&
                  cell !== undefined &&
                  String(cell).trim() !== ""
              )
            );
            // Check if Relics sheet exists in new spreadsheet
            if (!SheetsAPI.getSheetByName(newSheetID, "Relics")) {
              console.log("Relics sheet not found in new relic spreadsheet");
              return {
                success: false,
                message: "Relics sheet not found in new relic spreadsheet"
              };
            }

            return updateRelics(newSheetID, oldRelics);
          } 
          // else {
          // }
        } else {
          console.log("Version mismatch - skipping relic data import");
          return {
            success: false,
            message: "Version mismatch - skipping relic data import"
          };
        }
      } catch (error) {
        console.log("Error in importRelicsData: " + error.toString());
        return {
          success: false,
          message: "Error importing relics data: " + error.message,
        };
      }
    }

    function updateRelics(newSheetID, oldRelics) {
      try {
        // Get all data from new Relics sheet
        var newRelicsData = SheetsAPI.getValues(newSheetID, "Relics");
        if (!newRelicsData || newRelicsData.length < 3) {
          console.log("Not enough data in new Relics sheet");
          return {
            success: false,
            message: "Not enough data in new Relics sheet"
          };
        }

        var newRelicHeaderRow = null;
        var newRelicNameCol = null;
        var newRelicUnlockedCol = null;

        // Scan each row to find the header
        for (var row = 0; row < newRelicsData.length; row++) {
          var rowValues = newRelicsData[row];
          var relicNameIndex = rowValues.indexOf("Relic Name");
          var relicUnlockedIndex = rowValues.indexOf("Unlocked");
          if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
            newRelicHeaderRow = row + 1; // Convert to 1-based
            newRelicNameCol = relicNameIndex + 1; // Convert to 1-based
            newRelicUnlockedCol = relicUnlockedIndex + 1; // Convert to 1-based
            break;
          }
        }

        if (!newRelicHeaderRow) {
          console.log("Could not find header row in new Relics sheet");
          return {
            success: false,
            message: "Could not find header row in new Relics sheet"
          };
        }

        // Extract unlocked relic names from old data (skip header row)
        var oldRelicsNames = [];
        oldRelics.forEach(function (relic) {
          if (relic[relic.length - 1] === true || relic[relic.length - 1] === "TRUE" || relic[relic.length - 1] === "true") {
            oldRelicsNames.push(relic[0]);
          }
        });

        // Get new relic names starting from the first data row (skip header)
        var startRow = newRelicHeaderRow + 1;
        var endRow = newRelicsData.length;
        var newRelicNamesRange =
          "Relics!" +
          shared.columnToLetter(newRelicNameCol) +
          startRow + ":" +
          shared.columnToLetter(newRelicNameCol) +
          endRow;

        var newRelicNamesValues = SheetsAPI.getValues(
          newSheetID,
          newRelicNamesRange
        );
        if (!newRelicNamesValues) {
          console.log("Could not read new relic names");
          return {
            success: false,
            message: "Could not read new relic names"
          };
        }

        // Filter out empty rows and flatten the array
        var newRelicsNames = newRelicNamesValues
          .filter((row) =>
            row.some(
              (cell) =>
                cell !== null &&
                cell !== undefined &&
                String(cell || '').trim() !== ""
            )
          )
          .map((row) => row[0]);

        // Create unlocked status array
        var newRelicsUnlocked = [];
        newRelicsNames.forEach(function (relicName) {
          if (oldRelicsNames.includes(relicName)) {
            newRelicsUnlocked.push([true]);
          } else {
            newRelicsUnlocked.push([false]);
          }
        });
        // Update the unlocked column
        if (newRelicsUnlocked.length > 0) {
          var unlockedRange =
            "Relics!" +
            shared.columnToLetter(newRelicUnlockedCol) +
            startRow + ":" +
            shared.columnToLetter(newRelicUnlockedCol) +
            endRow;

          SheetsAPI.setValues(newSheetID, unlockedRange, newRelicsUnlocked);
          console.log(
            "Relics updated successfully: " +
              newRelicsUnlocked.length +
              " relics processed"
          );
          return {
            success: true,
            message: "Relics updated successfully: " + newRelicsUnlocked.length + " relics processed"
          };
        }
        console.log("No updates needed for relics");
        return {
          success: true,
          message: "No updates needed for relics"
        };
      } catch (error) {
        console.log("Error in updateRelics: " + error.toString());
        return {
          success: false,
          message: "Error updating relics: " + error.message,
        };
      }
    }
    return importRelicsData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
