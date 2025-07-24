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
        var oldDataResult = getVersionFunction();
        if (!oldDataResult || !oldDataResult.success) {
          console.log(
            `Error processing workshop data: ${oldDataResult.message}`
          );
          return oldDataResult;
        }

        var oldWorkshopLevels = oldDataResult.oldWorkshopLevels || [];
        var oldWorkshopPlusLevels = oldDataResult.oldWorkshopPlusLevels || [];

        var requiredRanges = ["Master Sheet"];
        var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
        if (!batchResults || batchResults.length === 0) {
          console.log(`Could not read required data from spreadsheet`);
          return {
            success: false,
            message: "Could not read required data from spreadsheet",
          };
        }

        var masterSheetData = batchResults[0].values;

        var workshopResult = updateWorkshopLevels(
          "Master Sheet",
          oldWorkshopLevels,
          oldWorkshopPlusLevels,
          masterSheetData
        );
        if (!workshopResult || !workshopResult.success) {
          console.log(
            `Error updating workshop levels: ${workshopResult.message}`
          );
          return workshopResult;
        }

        var batchUpdate = workshopResult.batchUpdate || [];

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
            message: workshopResult.message,
          };
        }

        return {
          success: true,
          message: `No updates needed for Workshop`,
        };
      } catch (error) {
        console.log("Error in importWorkshopData: " + error.toString());
        return {
          success: false,
          message: "Error importing workshop data: " + error.message,
        };
      }
    }

    function updateWorkshopLevels(
      sheetName,
      workshopLevels,
      workshopPlusLevels,
      masterSheetData
    ) {
      try {
        if (!masterSheetData || masterSheetData.length < 2) {
          console.log(`Not enough data in Master Sheet`);
          return {
            success: false,
            message: `Not enough data in Master Sheet`,
          };
        }

        var headerRow = masterSheetData[0];

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

        var batchUpdate = [];

        if (upgradeCol > 1 && workshopUnlocked.length) {
          var unlockedCol = shared.columnToLetter(upgradeCol - 1);
          var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
            workshopUnlocked.length + 1
          }`;
          batchUpdate.push({
            range: unlockedRange,
            values: workshopUnlocked,
          });
        }

        if (upgradeCol > 0 && workshopLevelsSplit.length) {
          var levelsStartCol = shared.columnToLetter(upgradeCol + 1);
          var levelsEndCol = shared.columnToLetter(upgradeCol + 2);
          var levelsRange = `${sheetName}!${levelsStartCol}2:${levelsEndCol}${
            workshopLevelsSplit.length + 1
          }`;
          batchUpdate.push({
            range: levelsRange,
            values: workshopLevelsSplit,
          });
        }

        if (enhancementCol > 0 && workshopPlusLevelsSplit.length) {
          var plusCol = shared.columnToLetter(enhancementCol + 2);
          var plusRange = `${sheetName}!${plusCol}2:${plusCol}${
            workshopPlusLevelsSplit.length + 1
          }`;
          batchUpdate.push({
            range: plusRange,
            values: workshopPlusLevelsSplit,
          });
        }

        if (batchUpdate.length > 0) {
          return {
            success: true,
            message: `Workshop levels updated successfully`,
            batchUpdate: batchUpdate,
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

    function version10() {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
          console.log(`_IDS sheet not found in new workshop spreadsheet`);
          return {
            success: false,
            message: `_IDS sheet™ not found in new workshop spreadsheet™`,
          };
        }

        var newSheetID = newSpreadsheet.spreadsheetId;

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
          console.log(`Master Sheet not found in new workshop spreadsheet`);
          return {
            success: false,
            message: `Master Sheet™ not found in new workshop spreadsheet™`,
          };
        }

        // Get header row to find WS and WS+ columns
        var updateHeaderBatchResult = SheetsAPI.batchGetValues(newSheetID, [
          "_IDS!1:1",
        ]);
        if (
          !updateHeaderBatchResult ||
          updateHeaderBatchResult.length === 0 ||
          !updateHeaderBatchResult[0].values
        ) {
          console.log(`Could not read header row from _IDS sheet`);
          return {
            success: false,
            message: `Could not read header row from _IDS sheet`,
          };
        }
        var headerValues = updateHeaderBatchResult[0].values;

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

        var updateWorkshopBatchResult = SheetsAPI.batchGetValues(newSheetID, [
          workshopLevelsRange,
        ]);
        if (
          !updateWorkshopBatchResult ||
          updateWorkshopBatchResult.length === 0 ||
          !updateWorkshopBatchResult[0].values
        ) {
          console.log(`Could not read workshop levels data`);
          return {
            success: false,
            message: `Could not read workshop levels data`,
          };
        }
        var oldWorkshopLevelsValues = updateWorkshopBatchResult[0].values;

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

        var updateWorkshopPlusBatchResult = SheetsAPI.batchGetValues(
          newSheetID,
          [workshopPlusLevelsRange]
        );
        if (
          !updateWorkshopPlusBatchResult ||
          updateWorkshopPlusBatchResult.length === 0 ||
          !updateWorkshopPlusBatchResult[0].values
        ) {
          console.log(`Could not read workshop plus levels data`);
          return {
            success: false,
            message: `Could not read workshop plus levels data`,
          };
        }
        var oldWorkshopPlusLevelsValues =
          updateWorkshopPlusBatchResult[0].values;

        // Filter out empty rows
        var oldWorkshopPlusLevels = oldWorkshopPlusLevelsValues.filter((row) =>
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

    function version19() {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!SheetsAPI.getSheetByName(newSpreadsheet, "_IDS")) {
          console.log(`_IDS sheet not found in new workshop spreadsheet`);
          return {
            success: false,
            message: `_IDS sheet™ not found in new workshop spreadsheet™`,
          };
        }

        var newSheetID = newSpreadsheet.spreadsheetId;

        if (!SheetsAPI.getSheetByName(newSpreadsheet, "Master Sheet")) {
          console.log(`Master Sheet not found in new workshop spreadsheet`);
          return {
            success: false,
            message: `Master Sheet™ not found in new workshop spreadsheet™`,
          };
        }

        var headerBatchResult = SheetsAPI.batchGetValues(newSheetID, [
          "_IDS!1:1",
        ]);
        if (
          !headerBatchResult ||
          headerBatchResult.length === 0 ||
          !headerBatchResult[0].values
        ) {
          console.log(`Could not read header row from _IDS sheet`);
          return {
            success: false,
            message: `Could not read header row from _IDS sheet™`,
          };
        }
        var headerValues = headerBatchResult[0].values;

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
          shared.columnToLetter(importWorkshopColStart + 11);

        var workshopBatchResult = SheetsAPI.batchGetValues(newSheetID, [
          workshopLevelsRange,
        ]);
        if (
          !workshopBatchResult ||
          workshopBatchResult.length === 0 ||
          !workshopBatchResult[0].values
        ) {
          console.log(`Could not read workshop levels data`);
          return {
            success: false,
            message: `Could not read workshop levels data`,
          };
        }
        var oldWorkshopLevelsValues = workshopBatchResult[0].values;

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
          shared.columnToLetter(importWorkshopPlusColStart + 6);

        var workshopPlusBatchResult = SheetsAPI.batchGetValues(newSheetID, [
          workshopPlusLevelsRange,
        ]);
        if (
          !workshopPlusBatchResult ||
          workshopPlusBatchResult.length === 0 ||
          !workshopPlusBatchResult[0].values
        ) {
          console.log(`Could not read workshop plus levels data`);
          return {
            success: false,
            message: `Could not read workshop plus levels data`,
          };
        }
        var oldWorkshopPlusLevelsValues = workshopPlusBatchResult[0].values;

        // Filter out empty rows
        var oldWorkshopPlusLevels = oldWorkshopPlusLevelsValues.filter((row) =>
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

    var convertVersionFunctions = {
      "v1.0": version10,
      // "v1.9": version19,
    };

    return importWorkshopData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = [
      "v1.0",
      // "v1.9",
    ];

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
