const cards = {

  /**
   * Reads Cards data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: cards.exportData");
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
        message: "Cards export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("cards.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Cards data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: cards.importData");

      var requiredRanges = [
        "Master Sheet",
        "Card Preset",
        "Card and Mastery Tracker",
        "IDS",
      ];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var cardPresetsData = batchResults[1].values;
      var cardTrackerData = batchResults[2].values;
      var idsData = batchResults[3].values;

      var batchUpdate = [];

      if (
        data.hasOwnProperty("oldCardsLevel") &&
        data.hasOwnProperty("oldCardSlots")
      ) {
        var oldCardsLevel = data.oldCardsLevel;
        var oldCardSlots = data.oldCardSlots || "";
        var levelsResult = cardsWriter.updateCardsLevels(
          "Master Sheet",
          oldCardsLevel,
          oldCardSlots,
          masterSheetData,
        );
        if (!levelsResult || !levelsResult.success) {
          console.log(`Error updating cards levels: ${levelsResult.message}`);
          return levelsResult;
        }
        batchUpdate = batchUpdate.concat(levelsResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldCardsPreset")) {
        var oldCardsPreset = data.oldCardsPreset;
        var shouldRemoveUsedCards = data.hasOwnProperty("shouldRemoveUsedCards")
          ? data.shouldRemoveUsedCards
          : true;
        var presetResult = cardsWriter.updateCardsPreset(
          "Card Preset",
          oldCardsPreset,
          shouldRemoveUsedCards,
          cardPresetsData,
        );
        if (!presetResult || !presetResult.success) {
          console.log(`Error updating cards preset: ${presetResult.message}`);
          return presetResult;
        }
        batchUpdate = batchUpdate.concat(presetResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldCardsTracker")) {
        var oldCardsTracker = data.oldCardsTracker;
        var trackerResult = cardsWriter.updateCardsTracker(
          "Card and Mastery Tracker",
          oldCardsTracker,
          cardTrackerData,
        );
        if (!trackerResult || !trackerResult.success) {
          console.log(`Error updating cards tracker: ${trackerResult.message}`);
          return trackerResult;
        }
        batchUpdate = batchUpdate.concat(trackerResult.batchUpdate || []);
      }

      labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Cards",
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
        message: `Cards import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("cards.importData", error, {
        note: `Error importing cards data`,
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": cardsReader.version1_0.bind(cardsReader),
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
      return versionUtils.compareVersions(b, a) === "newer" ? 1 : -1;
    });

    for (var i = 0; i < sortedThresholds.length; i++) {
      var threshold = sortedThresholds[i];
      var compareResult = versionUtils.compareVersions(oldVersion, threshold);

      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }

    return null;
  },
};
