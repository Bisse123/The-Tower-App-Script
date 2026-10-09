const playerStuff = {

  /**
   * Reads Player_&_Stuff data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
    try {
      console.log("Called: playerStuff.exportData");
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
        message: "Player & Stuff export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      var errorReport = errors.report("playerStuff.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported Player_&_Stuff data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: playerStuff.importData");

      var requiredRanges = ["Master Sheet", "Perk Preset", "IDS"];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var perksSheetData = batchResults[1].values;
      var idsData = batchResults[2].values;

      var batchUpdate = [];

      if (
        data.hasOwnProperty("oldPlayerStuffTierData") &&
        data.hasOwnProperty("oldPlayerStuffStatsData")
      ) {
        var oldPlayerStuffTierData = data.oldPlayerStuffTierData;
        var oldPlayerStuffStatsData = data.oldPlayerStuffStatsData;
        var playerStuffResult = playerStuffWriter.updatePlayerStuffData(
          "Master Sheet",
          oldPlayerStuffTierData,
          oldPlayerStuffStatsData,
          masterSheetData,
        );
        if (!playerStuffResult || !playerStuffResult.success) {
          console.log(
            `Error updating player data: ${playerStuffResult.message}`,
          );
          return playerStuffResult;
        }
        batchUpdate = batchUpdate.concat(playerStuffResult.batchUpdate || []);
      }

      if (data.hasOwnProperty("oldPerksPreset")) {
        var oldPerksPreset = data.oldPerksPreset;
        var shouldRemoveUsedPerks = data.hasOwnProperty("shouldRemoveUsedPerks")
          ? data.shouldRemoveUsedPerks
          : true;

        var playerPerksResult = playerStuffWriter.updatePlayerPerksPreset(
          "Perk Preset",
          oldPerksPreset,
          shouldRemoveUsedPerks,
          perksSheetData,
        );
        if (!playerPerksResult || !playerPerksResult.success) {
          console.log(
            `Error updating player perks data: ${playerPerksResult.message}`,
          );
          return playerPerksResult;
        }
        batchUpdate = batchUpdate.concat(playerPerksResult.batchUpdate || []);
      }

      labelUtils.addIDUpdatesToBatch(
        batchUpdate,
        "Player & Stuff",
        newSheetID,
        idsData,
        data.idMasterID,
      );

      var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
      if (!updateResult) {
        console.log(`Error applying batch updates to new spreadsheet`);
        return {
          success: false,
          message: "Error applying batch updates to new spreadsheet",
        };
      }

      return {
        success: true,
        message: `Player & Stuff import completed successfully`,
      };
    } catch (error) {
      var errorReport = errors.report("playerStuff.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v2.0": playerStuffReader.version2_0.bind(playerStuffReader),
      "v3.2": playerStuffReader.version3_2.bind(playerStuffReader),
      "v4.0": playerStuffReader.version4_0.bind(playerStuffReader),
      "v4.2": playerStuffReader.version4_2.bind(playerStuffReader),
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
