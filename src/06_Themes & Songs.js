const themes = {
  convertVersionFunctions: {
    },

  importData: function(sheetType, newThemeSpreadsheet) {
    function importThemesData(sheetType, newThemeSpreadsheetId) {
      var idType = sheetType + " ID";
      var targetThemes = [
        "Tower Skin",
        "Background Skin",
        "Songs",
        "Guardians",
        "Menu",
        "Profile Banner"
      ];
      // Get new theme version using SheetsAPI
      var newThemeVersion = SheetsAPI.getValue(newThemeSpreadsheetId, "STATS!A1");
      if (!newThemeVersion) {
        console.log("Error getting new theme version");
        return { success: false, message: "Error getting new theme version" };
      }

      // Get ID Master spreadsheet info
      var idMasterSpreadsheetInfo = shared.findSheetTypeID(
        newThemeSpreadsheetId,
        "IDS"
      );
      var idMasterSpreadsheetId = idMasterSpreadsheetInfo ? shared.extractSheetId(idMasterSpreadsheetInfo.id) : null;
      if (!idMasterSpreadsheetId) {
        console.log("Could not find ID Master spreadsheet");
        return { success: false, message: "Could not find ID Master spreadsheet" };
      }

      var oldThemeSpreadsheetInfo = shared.findSheetTypeID(
        idMasterSpreadsheetId,
        "IDS",
        idType
      );
      var oldThemeSpreadsheetId = oldThemeSpreadsheetInfo ? shared.extractSheetId(oldThemeSpreadsheetInfo.id) : null;
      if (!oldThemeSpreadsheetId) {
        console.log("Could not find old theme spreadsheet");
        return { success: false, message: "Could not find old theme spreadsheet" };
      }

      // Get old theme version using SheetsAPI
      var oldThemeVersion = SheetsAPI.getValue(oldThemeSpreadsheetId, "STATS!A1");
      if (!oldThemeVersion) {
        console.log("Error getting old theme version");
        return { success: false, message: "Error getting old theme version" };
      }

      var versionCheck = shared.compareVersions(
        oldThemeVersion,
        newThemeVersion
      );
      if (versionCheck === 0) {
        console.log("Same Version");
        // Get old themes data using SheetsAPI
        var oldThemesData = SheetsAPI.getDataRange(
          oldThemeSpreadsheetId,
          "Themes & Songs"
        );
        if (!oldThemesData) {
          console.log("Error getting old themes data");
          return { success: false, message: "Error getting old themes data" };
        }
        var oldThemesNames = getOldUnlockedThemesNames(
          targetThemes,
          oldThemesData
        );
        var updateResult = updateThemes(
          targetThemes,
          newThemeSpreadsheetId,
          "Themes & Songs",
          oldThemesNames
        );
        return updateResult;
      }
      // else {// Else do something to convert old version to new one (Future me problem)
      // }
      return { success: false, message: "Theme version mismatch or conversion not implemented." };
    }

    function updateThemes(targetThemes, spreadsheetId, sheetName, oldThemesNames) {
      // Get sheet data using SheetsAPI
      var newThemesData = SheetsAPI.getDataRange(spreadsheetId, sheetName);
      if (!newThemesData) {
        console.log("Error getting new themes data");
        return { success: false, message: "Error getting new themes data" };
      }

      // For each header, store {col, startRow} for quick reference
      var headerLocations = {};

      // Pre-scan to find header columns and their start rows
      for (var i = 0; i < newThemesData.length; i++) {
        for (var j = 0; j < newThemesData[i].length; j++) {
          var newThemeUnlocked = String(newThemesData[i][j] || '').trim();
          if (targetThemes.indexOf(newThemeUnlocked) !== -1) {
            // If not already recorded for this col, store its location
            if (!headerLocations[newThemeUnlocked]) {
              headerLocations[newThemeUnlocked] = [];
            }
            headerLocations[newThemeUnlocked].push({ col: j, startRow: i + 1 }); // +1 to start below header
          }
        }
      }

      var rangeUpdates = [];

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
              targetThemes.indexOf(String(newThemeName || '').trim()) !== -1
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
            rangeUpdates.push({
              range: sheetName + "!" + startCell + ":" + endCell,
              values: checkboxArr,
            });
          }
        });
      }

      // Apply all range updates using SheetsAPI
      rangeUpdates.forEach(function (update) {
        try {
          SheetsAPI.setValues(spreadsheetId, update.range, update.values);
        } catch (error) {
          console.log(
            "Error updating themes range " +
              update.range +
              ": " +
              error.toString()
          );
        }
      });
      return { success: true, message: "Themes updated successfully" };
    }

    function getOldUnlockedThemesNames(targetThemes, oldThemes) {
      var oldThemesNames = {}
      targetThemes.forEach(function(header) {
        oldThemesNames[header] = []
      })
      var currentHeader = null
      var headerCol = -1
      // Loop through each column first, then rows
      for (var col = 1; col < oldThemes[1].length; col++) {
        for (var row = 0; row < oldThemes.length; row++) {
          var oldThemeUnlocked = oldThemes[row][col];
          // If cell is a header
          if (targetThemes.indexOf(String(oldThemeUnlocked || '').trim()) !== -1) {
            currentHeader = String(oldThemeUnlocked || '').trim();
            headerCol = col;
            continue;
          }
          var isThemeUnlocked = oldThemeUnlocked === true || oldThemeUnlocked === "TRUE" || oldThemeUnlocked === "true";
          if (currentHeader && col === headerCol && isThemeUnlocked) {
            var oldThemeName = oldThemes[row][col+1];
            oldThemesNames[currentHeader].push(oldThemeName);
          }
        }
      }

      return oldThemesNames
    }
    return importThemesData(sheetType, newThemeSpreadsheet);
  },

  isCompatibleVersion: function(oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  }
};