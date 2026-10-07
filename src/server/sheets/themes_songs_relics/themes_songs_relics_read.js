const themesAndRelicsReader = {

  /**
   * Reads Themes_Songs_Relics data from a v4.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_0: function (oldSheetID) {
    try {
      console.log("Called: themesAndRelicsReader.version4_0");

      var requiredRanges = ["Themes & Songs", "Relics"];
      var batchResults = SheetsAPI.batchGetValues(oldSheetID, requiredRanges);
      if (!batchResults || batchResults.length < requiredRanges.length) {
        console.log(`Could not read required data from old spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from old spreadsheet",
        };
      }

      var oldThemesData = batchResults[0].values;
      var oldRelicsData = batchResults[1].values;

      var themesResult = this.getVersion4_0Themes(oldThemesData);
      if (!themesResult || !themesResult.success) {
        console.log(`Error converting themes: ${themesResult.message}`);
        return themesResult;
      }

      var relicsResult = this.getVersion4_0Relics(oldRelicsData);
      if (!relicsResult || !relicsResult.success) {
        console.log(`Error converting relics: ${relicsResult.message}`);
        return relicsResult;
      }

      return {
        success: true,
        oldThemesNames: themesResult.oldThemesNames,
        oldRelics: relicsResult.oldRelics,
      };

    } catch (error) {
      var errorReport = errors.report("themesAndRelicsReader.version4_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Themes from a v4.0 sheet's values.
   * @param {Array<Array<*>>} oldThemesData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0Themes: function (oldThemesData) {
    try {
      console.log("Called: themes.getVersion4_0Themes");
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
      var errorReport = errors.report("themesAndRelicsReader.getVersion4_0Themes", error, {
        oldThemesData: oldThemesData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts Relics from a v4.0 sheet's values.
   * @param {Array<Array<*>>} oldRelicsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0Relics: function (oldRelicsData) {
    try {
      console.log("Called: relics.getVersion4_0Relics");
      var oldRelicHeaderRow = -1;
      var relicNameIndex = -1;
      var relicUnlockedIndex = -1;

      for (var row = 0; row < oldRelicsData.length; row++) {
        var rowValues = oldRelicsData[row];
        relicNameIndex = rowValues.indexOf("Relic Name");
        relicUnlockedIndex = rowValues.indexOf("Unlocked");
        if (relicNameIndex !== -1 && relicUnlockedIndex !== -1) {
          oldRelicHeaderRow = row + 1;
          break;
        }
      }

      if (oldRelicHeaderRow === -1) {
        console.log(`Could not find header row in old Relics sheet`);
        return {
          success: false,
          message: `Could not find header row in old Relics sheet`,
        };
      }

      var startRow = oldRelicHeaderRow + 1;

      var oldRelics = [];
      oldRelicsData.slice(startRow - 1).forEach(function (row) {
        var relicName = row[relicNameIndex].trim();
        if (relicName.includes("T:")) {
          relicName = relicName.replace(/T:\s*/g, "T: ");
        }
        var isUnlocked = row[relicUnlockedIndex];

        if (
          relicName &&
          (isUnlocked === true ||
            isUnlocked === "TRUE" ||
            isUnlocked === "true")
        ) {
          oldRelics.push(relicName);
        }
      });

      return {
        success: true,
        oldRelics: oldRelics,
      };
    } catch (error) {
      var errorReport = errors.report("themesAndRelicsReader.getVersion4_0Relics", error, {
        oldRelicsData: oldRelicsData,
      });
      return errors.fail(errorReport);
    }
  },
};
