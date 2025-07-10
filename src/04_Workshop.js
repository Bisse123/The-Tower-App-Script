const workshop = {
  importData: function (versionDifference) {
    function importWorkshopData(versionDifference) {
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
        
        var getVersionFunction = convertVersionFunctions[versionDifference];
        if (!getVersionFunction) {
          console.log(`Unsupported version difference: ${versionDifference}`);
          return {
            success: false,
            message: `Unsupported version difference: ${versionDifference}`,
          };
        }
        var result = getVersionFunction();
        if (!result || !result.success) {
          console.log(`Error processing workshop data: ${result.message}`);
          return result;
        }

        var oldWorkshopLevels = result.oldWorkshopLevels || [];
        var oldWorkshopPlusLevels = result.oldWorkshopPlusLevels || [];
        return updateWorkshopLevels(newSheetID, oldWorkshopLevels, oldWorkshopPlusLevels);
      } catch (error) {
        console.log("Error in importWorkshopData: " + error.toString());
        return {
          success: false,
          message: "Error importing workshop data: " + error.message,
        };
      }
    }

    function version10() {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
          console.log(`_IDS sheet not found in new workshop spreadsheet`);
          return {
            success: false,
            message: `_IDS sheet not found in new workshop spreadsheet`,
          };
        }

        var newSheetID = newSpreadsheet.spreadsheetId;

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
          console.log(`Master Sheet not found in new workshop spreadsheet`);
          return {
            success: false,
            message: `Master Sheet not found in new workshop spreadsheet`,
          };
        }

        // Get header row to find WS and WS+ columns
        var headerValues = SheetsAPI.getValues(newSheetID, "_IDS!1:1");
        if (!headerValues || headerValues.length === 0) {
          console.log(`Could not read header row from _IDS sheet`);
          return {
            success: false,
            message: `Could not read header row from _IDS sheet`,
          };
        }

        var headerRow = headerValues[0];
        var importWorkshopColStart = headerRow.indexOf("WS") + 1;
        var importWorkshopPlusColStart = headerRow.indexOf("WS+") + 1;

        if (importWorkshopColStart === 0) {
          console.log(`WS column not found in _IDS sheet`);
          return {
            success: false,
            message: `WS column not found in _IDS sheet`,
          };
        }
        if (importWorkshopPlusColStart === 0) {
          console.log(`WS+ column not found in _IDS sheet`);
          return {
            success: false,
            message: `WS+ column not found in _IDS sheet`,
          };
        }

        // Get workshop levels data (4 columns starting from WS column)
        var workshopLevelsRange =
          "_IDS!" +
          shared.columnToLetter(importWorkshopColStart) +
          "2:" +
          shared.columnToLetter(importWorkshopColStart + 3);

        var oldWorkshopLevelsValues = SheetsAPI.getValues(
          newSheetID,
          workshopLevelsRange
        );
        if (!oldWorkshopLevelsValues) {
          console.log(`Could not read workshop levels data`);
          return {
            success: false,
            message: `Could not read workshop levels data`,
          };
        }

        // Filter out empty rows
        var oldWorkshopLevels = oldWorkshopLevelsValues.filter((row) =>
          row.some(
            (cell) =>
              cell !== null &&
              cell !== undefined &&
              String(cell || "").trim() !== ""
          )
        );

        // Get workshop plus levels data (3 columns starting from WS+ column)
        var workshopPlusLevelsRange =
          "_IDS!" +
          shared.columnToLetter(importWorkshopPlusColStart) +
          "2:" +
          shared.columnToLetter(importWorkshopPlusColStart + 2);

        var oldWorkshopPlusLevelsValues = SheetsAPI.getValues(
          newSheetID,
          workshopPlusLevelsRange
        );
        if (!oldWorkshopPlusLevelsValues) {
          console.log(`Could not read workshop plus levels data`);
          return {
            success: false,
            message: `Could not read workshop plus levels data`,
          };
        }

        // Filter out empty rows
        var oldWorkshopPlusLevels = oldWorkshopPlusLevelsValues.filter(
          (row) =>
            row.some(
              (cell) =>
                cell !== null &&
                cell !== undefined &&
                String(cell || "").trim() !== ""
            )
        );

        return {
          success: true,
          message: "Workshop levels processed successfully",
          oldWorkshopLevels: oldWorkshopLevels,
          oldWorkshopPlusLevels: oldWorkshopPlusLevels,
        };
      } catch (error) {
        console.log(`Error processing workshop data: ${error.toString()}`);
        return {
          success: false,
          message: `Error processing workshop data: ${error.message}`,
        };
      }
    }

    function updateWorkshopLevels(
      newSheetID,
      workshopLevels,
      workshopPlusLevels
    ) {
      try {
        // Get all data from Master Sheet to determine range and find columns
        var allData = SheetsAPI.getValues(newSheetID, "Master Sheet");
        if (!allData || allData.length < 2) {
          console.log(`Not enough data in Master Sheet`);
          return {
            success: false,
            message: `Not enough data in Master Sheet`,
          };
        }

        var headerRow = allData[0];
        var lastRow = allData.length;

        // Find header row and relevant columns
        var upgradeCol = headerRow.indexOf("Workshop Upgrade") + 1;
        var enhancementCol = headerRow.indexOf("Workshop Enhancement") + 1;

        if (upgradeCol === 0) {
          console.log(`Workshop Upgrade column not found`);
          return {
            success: false,
            message: `Workshop Upgrade column not found`,
          };
        }
        if (enhancementCol === 0) {
          console.log(`Workshop Enhancement column not found`);
          return {
            success: false,
            message: `Workshop Enhancement column not found`,
          };
        }

        // Split workshopLevels into workshopUnlocked and workshopLevelsSplit
        // workshopUnlocked: first element of each sublist
        // workshopLevelsSplit: 3rd and 4th elements (indices 2 and 3) of each sublist
        var workshopUnlocked = workshopLevels.map(function (sublist) {
          return [sublist[0]];
        });
        var workshopLevelsSplit = workshopLevels.map(function (sublist) {
          return [sublist[2], sublist[3]];
        });
        var workshopPlusLevelsSplit = workshopPlusLevels.map(function (
          sublist
        ) {
          return [sublist[2]];
        });

        // Prepare batch updates
        var batchUpdate = [];

        // Write workshopUnlocked: column before "Workshop Upgrade"
        if (upgradeCol > 1 && workshopUnlocked.length) {
          var unlockedCol = shared.columnToLetter(upgradeCol - 1);
          var unlockedRange =
            "Master Sheet!" +
            unlockedCol +
            "2:" +
            unlockedCol +
            (workshopUnlocked.length + 1);
          batchUpdate.push({
            range: unlockedRange,
            values: workshopUnlocked,
          });
        }

        // Write workshopLevelsSplit: 1 column after "Workshop Upgrade"
        if (upgradeCol > 0 && workshopLevelsSplit.length) {
          var levelsStartCol = shared.columnToLetter(upgradeCol + 1);
          var levelsEndCol = shared.columnToLetter(upgradeCol + 2);
          var levelsRange =
            "Master Sheet!" +
            levelsStartCol +
            "2:" +
            levelsEndCol +
            (workshopLevelsSplit.length + 1);
          batchUpdate.push({
            range: levelsRange,
            values: workshopLevelsSplit,
          });
        }

        // Write workshopPlusLevelsSplit: 2 columns after "Workshop Enhancement"
        if (enhancementCol > 0 && workshopPlusLevelsSplit.length) {
          var plusCol = shared.columnToLetter(enhancementCol + 2);
          var plusRange =
            "Master Sheet!" +
            plusCol +
            "2:" +
            plusCol +
            (workshopPlusLevelsSplit.length + 1);
          batchUpdate.push({
            range: plusRange,
            values: workshopPlusLevelsSplit,
          });
        }

        // Execute batch updates
        if (batchUpdate.length > 0) {
          SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
          // console.log(`Workshop levels updated successfully`);
          return {
            success: true,
            message: `Workshop levels updated successfully`,
          };
        }
        return {
          success: true,
          message: `No updates needed for workshop levels`,
        };
      } catch (error) {
        console.log(`Error in updateWorkshopLevels: ${error.toString()}`);
        return {
          success: false,
          message: `Error updating workshop levels: ${error.message}`,
        };
      }
    }

    var convertVersionFunctions = {
      "v1.0": version10,
    };

    return importWorkshopData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = [
      "v1.0"
    ];
    
    var sortedThresholds = versionCompatibility.slice().sort(function(a, b) {
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