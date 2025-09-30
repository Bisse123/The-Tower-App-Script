const themes = {
  exportData: function (versionDifference) {
    try {
      console.log("Called: themes.exportData");
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
        message: "Themes & Songs export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting themes data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      console.log("Called: themes.importData");
      var newSpreadsheet = spreadsheets("Themes & Songs newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }

      // Batch get required data for update function only
      var requiredRanges = ["Themes & Songs", "IDS"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var newThemesData = batchResults[0].values;
      var idsData = batchResults[1].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(
        newSheetID,
        "IDS",
        "IDS Master's",
        idsData
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

      var batchUpdate = [];

      // Only update themes if key exists
      if (data.hasOwnProperty("oldThemesNames")) {
        var oldThemesNames = data.oldThemesNames;
        var themesResult = this.updateThemes(
          "Themes & Songs",
          oldThemesNames,
          newThemesData
        );
        if (!themesResult || !themesResult.success) {
          console.log(`Error updating themes: ${themesResult.message}`);
          return themesResult;
        }
        batchUpdate = batchUpdate.concat(themesResult.batchUpdate || []);
      }

      // Add import status update to batch if any update was made
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });

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
          message: `Themes & Songs import completed successfully`,
        };
      }
      return {
        success: true,
        message: "No themes data to update",
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing themes data: ${error.message}`,
      };
    }
  },

  updateThemes: function (sheetName, oldThemesNames, newThemesData) {
    try {
      console.log("Called: themes.updateThemes");
      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner",
      ];
      if (!newThemesData) {
        console.log(`Error getting new themes data`);
        return { success: false, message: "Error getting new themes data" };
      }
      var autoFill = false;
      if (oldThemesNames.hasOwnProperty("autoFill")) {
        autoFill =
          oldThemesNames.autoFill === true ||
          oldThemesNames.autoFill === "TRUE" ||
          oldThemesNames.autoFill === "true";
        if (!autoFill) {
          targetThemes.push("Milestone Skin");
        }
      }

      // For each header, store {col, startRow} for quick reference
      var headerLocations = {};
      var batchUpdate = [];

      // Pre-scan to find header columns and their start rows
      for (var i = 0; i < newThemesData.length; i++) {
        for (var j = 0; j < newThemesData[i].length; j++) {
          var newThemeUnlocked = String(newThemesData[i][j] || "").trim();
          if (newThemeUnlocked === "Auto-fill from Player and Stuff") {
            batchUpdate.push({
              range: `${sheetName}!${shared.columnToLetter(j + 1) + (i + 2)}`,
              values: [[autoFill]],
            });
            break;
          }
          if (targetThemes.indexOf(newThemeUnlocked) !== -1) {
            // If not already recorded for this col, store its location
            if (!headerLocations[newThemeUnlocked]) {
              headerLocations[newThemeUnlocked] = [];
            }
            headerLocations[newThemeUnlocked].push({ col: j, startRow: i + 1 }); // +1 to start below header
          }
        }
      }

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
    } catch (error) {
      console.log("Error in updateThemes: " + error.toString());
      return {
        success: false,
        message: "Error in updateThemes: " + error.message,
      };
    }
  },

  version216: function () {
    try {
      console.log("Called: themes.version216");
      var oldSpreadsheet = spreadsheets("Themes & Songs oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var themesValuesRange = "Themes & Songs";
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

      var themesData = this.getversion216Themes(oldThemesData);
      return themesData;
    } catch (error) {
      console.log("Error in version216: " + error.toString());
      return {
        success: false,
        message: "Error in version216: " + error.message,
      };
    }
  },

  version10: function () {
    try {
      console.log("Called: themes.version10");
      var oldSpreadsheet = spreadsheets("Themes & Songs oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      var themesValuesRange = "Themes & Songs";
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

      var themesData = this.getVersion10Themes(oldThemesData);
      return themesData;
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10Themes: function (oldThemesData) {
    try {
      console.log("Called: themes.getVersion10Themes");
      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner",
      ];

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
            if (oldThemesData[row][col + 2] === "Tier Unlocked") {
              currentHeader = "Milestone Skin";
              oldThemesNames["Milestone Skin"] = [];
            }
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
        oldThemesNames: oldThemesNames,
      };
    } catch (error) {
      console.log("Error in getVersion10Themes: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10Themes: " + error.message,
      };
    }
  },

  getversion216Themes: function (oldThemesData) {
    try {
      console.log("Called: themes.getversion216Themes");
      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner",
        "Milestone Skin",
      ];

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
          if (oldThemeUnlocked === "Auto-fill from Player and Stuff") {
            oldThemesNames["autoFill"] = oldThemesData[row + 1][col];
            continue;
          } 
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
        oldThemesNames: oldThemesNames,
      };
    } catch (error) {
      console.log("Error in getversion216Themes: " + error.toString());
      return {
        success: false,
        message: "Error in getversion216Themes: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
      "v2.1.4": this.version216.bind(this),
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
