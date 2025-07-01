const relics = {
  convertVersionFunctions: {},

  importData: function (sheetType, newRelicSpreadsheetId) {
    function importRelicsData(sheetType, newRelicSpreadsheetId) {
      try {
        var idType = sheetType + " ID";

        // Check if required sheets exist
        if (!SheetsAPI.hasSheet(newRelicSpreadsheetId, "IDS")) {
          console.log("IDS sheet not found in new relic spreadsheet");
          return {
            success: false,
            message: "IDS sheet not found in new relic spreadsheet"
          };
        }
        if (!SheetsAPI.hasSheet(newRelicSpreadsheetId, "STATS")) {
          console.log("STATS sheet not found in new relic spreadsheet");
          return {
            success: false,
            message: "STATS sheet not found in new relic spreadsheet"
          };
        }

        // Get version from STATS sheet
        var newRelicVersion = SheetsAPI.getValue(
          newRelicSpreadsheetId,
          "STATS!A1"
        );

        // Get ID Master spreadsheet info
        var idMasterSpreadsheetInfo = shared.findSheetTypeID(
          newRelicSpreadsheetId,
          "IDS"
        );
        var idMasterSpreadsheetId = shared.extractSheetId(
          idMasterSpreadsheetInfo.id
        );
        if (!idMasterSpreadsheetId) {
          console.log("Could not find ID Master spreadsheet");
          return {
            success: false,
            message: "Could not find ID Master spreadsheet"
          };
        }

        var oldRelicSpreadsheetInfo = shared.findSheetTypeID(
          idMasterSpreadsheetId,
          "IDS",
          idType
        );
        var oldRelicSpreadsheetId = shared.extractSheetId(
          oldRelicSpreadsheetInfo.id
        );
        if (!oldRelicSpreadsheetId) {
          console.log("Could not find old relic spreadsheet");
          return {
            success: false,
            message: "Could not find old relic spreadsheet"
          };
        }

        var oldRelicVersion = SheetsAPI.getValue(
          oldRelicSpreadsheetId,
          "STATS!A1"
        );
        var versionCheck = shared.compareVersions(
          oldRelicVersion,
          newRelicVersion
        );

        if (versionCheck === 0) {
          console.log("Same Version");

          // Check if Relics sheet exists in old spreadsheet
          if (!SheetsAPI.hasSheet(oldRelicSpreadsheetId, "Relics")) {
            console.log("Relics sheet not found in old relic spreadsheet");
            return {
              success: false,
              message: "Relics sheet not found in old relic spreadsheet"
            };
          }

          // Get all data from old Relics sheet
          var oldRelicsData = SheetsAPI.getDataRange(
            oldRelicSpreadsheetId,
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
              oldRelicSpreadsheetId,
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
            if (!SheetsAPI.hasSheet(newRelicSpreadsheetId, "Relics")) {
              console.log("Relics sheet not found in new relic spreadsheet");
              return {
                success: false,
                message: "Relics sheet not found in new relic spreadsheet"
              };
            }

            return updateRelics(newRelicSpreadsheetId, oldRelics);
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

    function updateRelics(spreadsheetId, oldRelics) {
      try {
        // Get all data from new Relics sheet
        var newRelicsData = SheetsAPI.getValues(spreadsheetId, "Relics");
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
          spreadsheetId,
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

          SheetsAPI.setValues(spreadsheetId, unlockedRange, newRelicsUnlocked);
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
        console.log("No relics to update");
        return {
          success: false,
          message: "No relics to update"
        };
      } catch (error) {
        console.log("Error in updateRelics: " + error.toString());
        return {
          success: false,
          message: "Error updating relics: " + error.message,
        };
      }
    }
    return importRelicsData(sheetType, newRelicSpreadsheetId);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
