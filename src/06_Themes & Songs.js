const themes = {
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
        console.log(`Error processing themes data: ${oldDataResult.message}`);
        return oldDataResult;
      }

      var targetThemes = oldDataResult.targetThemes || [];
      var oldThemesNames = oldDataResult.oldThemesNames || {};

      // Batch get required data for update function only
      var requiredRanges = ["Themes & Songs"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var newThemesData = batchResults[0].values;

      var themesResult = this.updateThemes(
        targetThemes,
        "Themes & Songs",
        oldThemesNames,
        newThemesData
      );
      if (!themesResult || !themesResult.success) {
        console.log(`Error updating themes: ${themesResult.message}`);
        return themesResult;
      }

      var batchUpdate = themesResult.batchUpdate || [];

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
          message: themesResult.message,
        };
      }

      return {
        success: true,
        message: `No updates needed for Themes & Songs`,
      };
    } catch (error) {
      console.log(`Error importing themes data: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing themes data: ${error.message}`,
      };
    }
  },

  updateThemes: function (
    targetThemes,
    sheetName,
    oldThemesNames,
    newThemesData
  ) {
    if (!newThemesData) {
      console.log(`Error getting new themes data`);
      return { success: false, message: "Error getting new themes data" };
    }

    // For each header, store {col, startRow} for quick reference
    var headerLocations = {};

    // Pre-scan to find header columns and their start rows
    for (var i = 0; i < newThemesData.length; i++) {
      for (var j = 0; j < newThemesData[i].length; j++) {
        var newThemeUnlocked = String(newThemesData[i][j] || "").trim();
        if (targetThemes.indexOf(newThemeUnlocked) !== -1) {
          // If not already recorded for this col, store its location
          if (!headerLocations[newThemeUnlocked]) {
            headerLocations[newThemeUnlocked] = [];
          }
          headerLocations[newThemeUnlocked].push({ col: j, startRow: i + 1 }); // +1 to start below header
        }
      }
    }

    var batchUpdate = [];

    // For each header, possibly in multiple places
    for (var key in headerLocations) {
      headerLocations[key].forEach(function (loc) {
        var checkboxCol = loc.col;
        var startRow = loc.startRow;
        var checkedSet = new Set((oldThemesNames[key] || []).map(String));
        var checkboxArr = [];

        for (var row = startRow; row < newThemesData.length; row++) {
          var newThemeName = newThemesData[row][checkboxCol + 1];
          if (
            newThemeName === "" ||
            newThemeName === null ||
            typeof newThemeName === "undefined" ||
            targetThemes.indexOf(String(newThemeName || "").trim()) !== -1
          ) {
            break;
          }
          checkboxArr.push([checkedSet.has(String(newThemeName))]);
        }

        if (checkboxArr.length > 0) {
          var startCell =
            shared.columnToLetter(checkboxCol + 1) + (startRow + 1);
          var endCell =
            shared.columnToLetter(checkboxCol + 1) +
            (startRow + checkboxArr.length);
          batchUpdate.push({
            range: `${sheetName}!${startCell}:${endCell}`,
            values: checkboxArr,
          });
        }
      });
    }

    if (batchUpdate.length !== 0) {
      return {
        success: true,
        message: "Themes & Songs updated successfully",
        batchUpdate: batchUpdate,
      };
    }
    return { success: true, message: "No updates needed for Themes & Songs" };
  },

  version10: function () {
    try {
      var oldSpreadsheet = spreadsheets("oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner",
      ];

      var themesValuesRange = "Themes & Songs"
      var themesOldBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        themesValuesRange,
      ]);
      if (
        !themesOldBatchResult ||
        themesOldBatchResult.length === 0 ||
        !themesOldBatchResult[0].values
      ) {
        console.log(`Error getting old themes data`);
        return { success: false, message: "Error getting old themes data" };
      }
      var oldThemesData = themesOldBatchResult[0].values;
      var oldThemesNames = {};

      targetThemes.forEach(function (header) {
        oldThemesNames[header] = [];
      });
      var currentHeader = null;
      var headerCol = -1;
      // Loop through each column first, then rows
      for (var col = 1; col < oldThemesData[1].length; col++) {
        for (var row = 0; row < oldThemesData.length; row++) {
          var oldThemeUnlocked = oldThemesData[row][col];
          // If cell is a header
          if (
            targetThemes.indexOf(String(oldThemeUnlocked || "").trim()) !== -1
          ) {
            currentHeader = String(oldThemeUnlocked || "").trim();
            headerCol = col;
            continue;
          }
          var isThemeUnlocked =
            oldThemeUnlocked === true ||
            oldThemeUnlocked === "TRUE" ||
            oldThemeUnlocked === "true";
          if (currentHeader && col === headerCol && isThemeUnlocked) {
            var oldThemeName = oldThemesData[row][col + 1];
            oldThemesNames[currentHeader].push(oldThemeName);
          }
        }
      }

      return {
        success: true,
        targetThemes: targetThemes,
        oldThemesNames: oldThemesNames,
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
