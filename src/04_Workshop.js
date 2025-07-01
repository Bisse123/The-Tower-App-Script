const workshop = {
  convertVersionFunctions: {},

  importData: function (sheetType, newWorkshopSpreadsheetId) {
    function importWorkshopData(sheetType, newWorkshopSpreadsheetId) {
      try {
        var idType = sheetType + " ID";

        // Check if required sheets exist
        if (!SheetsAPI.hasSheet(newWorkshopSpreadsheetId, "IDS")) {
          console.log("IDS sheet not found in new workshop spreadsheet");
          return {
            success: false,
            message: "IDS sheet not found in new workshop spreadsheet",
          };
        }
        if (!SheetsAPI.hasSheet(newWorkshopSpreadsheetId, "EXPORT")) {
          console.log("EXPORT sheet not found in new workshop spreadsheet");
          return {
            success: false,
            message: "EXPORT sheet not found in new workshop spreadsheet",
          };
        }

        // Get version from EXPORT sheet
        var newWorkshopVersion = SheetsAPI.getValue(
          newWorkshopSpreadsheetId,
          "EXPORT!A1"
        );
        if (!newWorkshopVersion) {
          console.log("Error getting new workshop version");
          return {
            success: false,
            message: "Error getting new workshop version",
          };
        }

        // Get ID Master spreadsheet info
        var idMasterSpreadsheetInfo = shared.findSheetTypeID(
          newWorkshopSpreadsheetId,
          "IDS"
        );
        if (!idMasterSpreadsheetInfo || !idMasterSpreadsheetInfo.id) {
          console.log("Could not find ID Master spreadsheet info");
          return {
            success: false,
            message: "Could not find ID Master spreadsheet info",
          };
        }

        var idMasterSpreadsheetId = shared.extractSheetId(
          idMasterSpreadsheetInfo.id
        );
        if (!idMasterSpreadsheetId) {
          console.log("Could not find ID Master spreadsheet");
          return {
            success: false,
            message: "Could not find ID Master spreadsheet",
          };
        }

        var oldWorkshopSpreadsheetInfo = shared.findSheetTypeID(
          idMasterSpreadsheetId,
          "IDS",
          idType
        );
        if (!oldWorkshopSpreadsheetInfo || !oldWorkshopSpreadsheetInfo.id) {
          console.log("Could not find old workshop spreadsheet info");
          return {
            success: false,
            message: "Could not find old workshop spreadsheet info",
          };
        }

        var oldWorkshopSpreadsheetId = shared.extractSheetId(
          oldWorkshopSpreadsheetInfo.id
        );
        if (!oldWorkshopSpreadsheetId) {
          console.log("Could not find old workshop spreadsheet");
          return {
            success: false,
            message: "Could not find old workshop spreadsheet",
          };
        }

        var oldWorkshopVersion = SheetsAPI.getValue(
          oldWorkshopSpreadsheetId,
          "EXPORT!A1"
        );
        if (!oldWorkshopVersion) {
          console.log("Error getting old workshop version");
          return {
            success: false,
            message: "Error getting old workshop version",
          };
        }

        var versionCheck = shared.compareVersions(
          oldWorkshopVersion,
          newWorkshopVersion
        );

        if (versionCheck === 0) {
          console.log("Same Version - proceeding with workshop data import");

          // Check if _IDS sheet exists
          if (!SheetsAPI.hasSheet(newWorkshopSpreadsheetId, "_IDS")) {
            console.log("_IDS sheet not found in new workshop spreadsheet");
            return {
              success: false,
              message: "_IDS sheet not found in new workshop spreadsheet",
              code: 'NO_IDS_SHEET',
              context: { spreadsheetId: newWorkshopSpreadsheetId }
            };
          }

          // Get header row to find WS and WS+ columns
          var headerValues = SheetsAPI.getValues(
            newWorkshopSpreadsheetId,
            "_IDS!1:1"
          );
          if (!headerValues || headerValues.length === 0) {
            console.log("Could not read header row from _IDS sheet");
            return {
              success: false,
              message: "Could not read header row from _IDS sheet",
              code: 'NO_HEADER_ROW',
              context: { spreadsheetId: newWorkshopSpreadsheetId }
            };
          }

          var headerRow = headerValues[0];
          var importWorkshopColStart = headerRow.indexOf("WS") + 1;
          var importWorkshopPlusColStart = headerRow.indexOf("WS+") + 1;

          if (importWorkshopColStart === 0) {
            console.log("WS column not found in _IDS sheet");
            return {
              success: false,
              message: "WS column not found in _IDS sheet",
              code: 'NO_WS_COLUMN',
              context: { spreadsheetId: newWorkshopSpreadsheetId }
            };
          }
          if (importWorkshopPlusColStart === 0) {
            console.log("WS+ column not found in _IDS sheet");
            return {
              success: false,
              message: "WS+ column not found in _IDS sheet",
              code: 'NO_WSPLUS_COLUMN',
              context: { spreadsheetId: newWorkshopSpreadsheetId }
            };
          }

          // Get workshop levels data (4 columns starting from WS column)
          var workshopLevelsRange =
            "_IDS!" +
            shared.columnToLetter(importWorkshopColStart) +
            "2:" +
            shared.columnToLetter(importWorkshopColStart + 3);

          var oldWorkshopLevelsValues = SheetsAPI.getValues(
            newWorkshopSpreadsheetId,
            workshopLevelsRange
          );
          if (!oldWorkshopLevelsValues) {
            console.log("Could not read workshop levels data");
            return {
              success: false,
              message: "Could not read workshop levels data",
              code: 'NO_WORKSHOP_LEVELS',
              context: { spreadsheetId: newWorkshopSpreadsheetId, range: workshopLevelsRange }
            };
          }

          // Filter out empty rows
          var oldWorkshopLevels = oldWorkshopLevelsValues.filter((row) =>
            row.some(
              (cell) =>
                cell !== null &&
                cell !== undefined &&
                String(cell || '').trim() !== ""
            )
          );

          // Get workshop plus levels data (3 columns starting from WS+ column)
          var workshopPlusLevelsRange =
            "_IDS!" +
            shared.columnToLetter(importWorkshopPlusColStart) +
            "2:" +
            shared.columnToLetter(importWorkshopPlusColStart + 2);

          var oldWorkshopPlusLevelsValues = SheetsAPI.getValues(
            newWorkshopSpreadsheetId,
            workshopPlusLevelsRange
          );
          if (!oldWorkshopPlusLevelsValues) {
            console.log("Could not read workshop plus levels data");
            return {
              success: false,
              message: "Could not read workshop plus levels data",
              code: 'NO_WORKSHOP_PLUS_LEVELS',
              context: { spreadsheetId: newWorkshopSpreadsheetId, range: workshopPlusLevelsRange }
            };
          }

          // Filter out empty rows
          var oldWorkshopPlusLevels = oldWorkshopPlusLevelsValues.filter(
            (row) =>
            row.some(
              (cell) =>
                cell !== null &&
                cell !== undefined &&
                String(cell || '').trim() !== ""
            )
          );

          // Check if Master Sheet exists
          if (!SheetsAPI.hasSheet(newWorkshopSpreadsheetId, "Master Sheet")) {
            console.log("Master Sheet not found in new workshop spreadsheet");
            return {
              success: false,
              message: "Master Sheet not found in new workshop spreadsheet"
            };
          }

          return updateWorkshopLevels(
            newWorkshopSpreadsheetId,
            oldWorkshopLevels,
            oldWorkshopPlusLevels
          );
        } else {
          console.log("Version mismatch - skipping workshop data import");
          return {
            success: false,
            message: "Workshop version mismatch",
          };
        }
      } catch (error) {
        console.log("Error in importWorkshopData: " + error.toString());
        return {
          success: false,
          message: "Error importing workshop data: " + error.message,
        };
      }
    }

    function updateWorkshopLevels(
      spreadsheetId,
      workshopLevels,
      workshopPlusLevels
    ) {
      try {
        // Get all data from Master Sheet to determine range and find columns
        var allData = SheetsAPI.getValues(spreadsheetId, "Master Sheet");
        if (!allData || allData.length < 2) {
          console.log("Not enough data in Master Sheet");
          return {
            success: false,
            message: "Not enough data in Master Sheet"
          };
        }

        var headerRow = allData[0];
        var lastRow = allData.length;

        // Find header row and relevant columns
        var upgradeCol = headerRow.indexOf("Workshop Upgrade") + 1;
        var enhancementCol = headerRow.indexOf("Workshop Enhancement") + 1;

        if (upgradeCol === 0) {
          console.log("Workshop Upgrade column not found");
          return {
            success: false,
            message: "Workshop Upgrade column not found"
          };
        }
        if (enhancementCol === 0) {
          console.log("Workshop Enhancement column not found");
          return {
            success: false,
            message: "Workshop Enhancement column not found"
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
        var batchUpdates = [];

        // Write workshopUnlocked: column before "Workshop Upgrade"
        if (upgradeCol > 1 && workshopUnlocked.length) {
          var unlockedCol = shared.columnToLetter(upgradeCol - 1);
          var unlockedRange =
            "Master Sheet!" +
            unlockedCol +
            "2:" +
            unlockedCol +
            (workshopUnlocked.length + 1);
          batchUpdates.push({
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
          batchUpdates.push({
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
          batchUpdates.push({
            range: plusRange,
            values: workshopPlusLevelsSplit,
          });
        }

        // Execute batch updates
        if (batchUpdates.length > 0) {
          batchUpdates.forEach(function (update) {
            SheetsAPI.setValues(spreadsheetId, update.range, update.values);
          });
          console.log("Workshop levels updated successfully");
          return {
            success: true,
            message: "Workshop levels updated successfully",
          };
        }
        return {
          success: true,
          message: "No updates needed for workshop levels",
        };
      } catch (error) {
        console.log("Error in updateWorkshopLevels: " + error.toString());
        return {
          success: false,
          message: "Error updating workshop levels: " + error.message,
        };
      }
    }

    return importWorkshopData(sheetType, newWorkshopSpreadsheetId);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
