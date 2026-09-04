const themes = {

  /**
   * Reads Themes_&_Songs data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
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

      var oldDataResult = getVersionFunction(oldSheetID);
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
      var errorReport = errors.report("themes.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Themes_&_Songs data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: themes.importData");

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

      var batchUpdate = [];

      if (data.hasOwnProperty("oldThemesNames")) {
        var oldThemesNames = data.oldThemesNames;
        var themesResult = this.updateThemes(
          "Themes & Songs",
          oldThemesNames,
          newThemesData,
        );
        if (!themesResult || !themesResult.success) {
          console.log(`Error updating themes: ${themesResult.message}`);
          return themesResult;
        }
        batchUpdate = batchUpdate.concat(themesResult.batchUpdate || []);
      }

      shared.addIDUpdatesToBatch(
        batchUpdate,
        "Themes & Songs",
        newSheetID,
        idsData,
        data.idMasterID,
      );

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
    } catch (error) {
      var errorReport = errors.report("themes.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes Themes into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldThemesNames
   * @param {Object} newThemesData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
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
      }
      if (!autoFill) {
        targetThemes.push("Milestone Skin");
      }

      var headerLocations = {};
      var batchUpdate = [];

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

            if (!headerLocations[newThemeUnlocked]) {
              headerLocations[newThemeUnlocked] = [];
            }
            headerLocations[newThemeUnlocked].push({ col: j, startRow: i + 1 });
          }
        }
      }

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
            checkboxArr.push([checkedSet.has(String(newThemeName).trim())]);
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
      var errorReport = errors.report("themes.updateThemes", error, {
        sheetName: sheetName,
        oldThemesNames: oldThemesNames,
        newThemesData: newThemesData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Themes_&_Songs data from a v2.1.6 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_1_6: function (oldSheetID) {
    try {
      console.log("Called: themes.version2_1_6");

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

      var themesData = this.getVersion2_1_6Themes(oldThemesData);
      return themesData;
    } catch (error) {
      var errorReport = errors.report("themes.version2_1_6", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads Themes_&_Songs data from a v1.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: themes.version1_0");

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

      var themesData = this.getVersion1_0Themes(oldThemesData);
      return themesData;
    } catch (error) {
      var errorReport = errors.report("themes.version1_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Themes from a v2.1.6 sheet's values.
   * @param {Array<Array<*>>} oldThemesData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_1_6Themes: function (oldThemesData) {
    try {
      console.log("Called: themes.getVersion2_1_6Themes");
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

      for (var col = 1; col < oldThemesData[1].length; col++) {
        for (var row = 0; row < oldThemesData.length; row++) {
          var oldThemeUnlocked = oldThemesData[row][col];
          if (oldThemeUnlocked === "Auto-fill from Player and Stuff") {
            oldThemesNames["autoFill"] = oldThemesData[row + 1][col];
            continue;
          }

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
      var errorReport = errors.report("themes.getVersion2_1_6Themes", error, {
        oldThemesData: oldThemesData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Themes from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldThemesData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0Themes: function (oldThemesData) {
    try {
      console.log("Called: themes.getVersion1_0Themes");
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

      for (var col = 1; col < oldThemesData[1].length; col++) {
        for (var row = 0; row < oldThemesData.length; row++) {
          var oldThemeUnlocked = oldThemesData[row][col];

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
      var errorReport = errors.report("themes.getVersion1_0Themes", error, {
        oldThemesData: oldThemesData,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
      "v2.1.6": this.version2_1_6.bind(this),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
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
