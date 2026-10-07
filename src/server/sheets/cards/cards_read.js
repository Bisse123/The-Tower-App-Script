const cardsReader = {

  /**
   * Reads Cards data from a v1.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version1_0: function (oldSheetID) {
    try {
      console.log("Called: cardsReader.version1_0");

      var oldRanges = [
        "Card Preset",
        "Card and Mastery Tracker",
        "EXPORT!B5:D",
        "EXPORT!C2",
      ];
      var oldBatchResult = SheetsAPI.batchGetValues(oldSheetID, oldRanges);

      var oldCardsPresetData = oldBatchResult[0].values;
      var oldCardsTrackerData = oldBatchResult[1].values;
      var oldCardsLevelData = oldBatchResult[2].values;
      var oldCardSlotsData = oldBatchResult[3].values;

      var cardsPresetData = this.getVersion1_0CardsPreset(oldCardsPresetData);
      if (!cardsPresetData || !cardsPresetData.success) {
        return cardsPresetData;
      }

      var cardsTrackerData =
        this.getVersion1_0CardsTracker(oldCardsTrackerData);
      if (!cardsTrackerData || !cardsTrackerData.success) {
        return cardsTrackerData;
      }

      var cardsLevelData = this.getVersion1_0CardsLevel(
        oldCardsLevelData,
        oldCardSlotsData,
      );
      if (!cardsLevelData || !cardsLevelData.success) {
        return cardsLevelData;
      }

      return {
        success: true,
        message: "Cards processed successfully",
        oldCardsLevel: cardsLevelData.oldCardsLevel,
        oldCardSlots: cardsLevelData.oldCardSlots,
        oldCardsPreset: cardsPresetData.oldCardsPreset,
        oldCardsTracker: cardsTrackerData.oldCardsTracker,
        shouldRemoveUsedCards: cardsPresetData.shouldRemoveUsedCards,
      };
    } catch (error) {
      var errorReport = errors.report("cardsReader.version1_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts CardsTracker from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldCardsTrackerData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0CardsTracker: function (oldCardsTrackerData) {
    try {
      console.log("Called: cardsReader.getVersion1_0CardsTracker");
      var ignoreprioValues = ["*", "Purchased"];
      var oldCardsTracker = {};
      for (
        var rowIndex = 0;
        rowIndex < oldCardsTrackerData.length;
        rowIndex++
      ) {
        var row = oldCardsTrackerData[rowIndex];
        var cardNameColIndex = row.indexOf("Card");
        var progressColIndex = row.indexOf("Progress");
        var priorityColIndex = row.indexOf("Priority");
        console.log(
          "rowIndex:",
          rowIndex,
          "cardNameColIndex:",
          cardNameColIndex,
          "progressColIndex:",
          progressColIndex,
          "priorityColIndex:",
          priorityColIndex,
        );
        if (cardNameColIndex !== -1) {
          for (
            var rowIdx = rowIndex + 1;
            rowIdx < oldCardsTrackerData.length;
            rowIdx++
          ) {
            var trackerRow = oldCardsTrackerData[rowIdx];
            var cardName = trackerRow[cardNameColIndex];
            if (!cardName || String(cardName).trim() === "") {
              break;
            }
            var progress = trackerRow[progressColIndex];
            var priority = trackerRow[priorityColIndex];
            if (
              priority &&
              ignoreprioValues.includes(String(priority).trim())
            ) {
              priority = null;
            }
            oldCardsTracker[cardName] = {
              progress: progress || null,
              priority: priority || null,
            };
          }
          break;
        }
      }
      return {
        success: true,
        message: "Cards tracker processed successfully",
        oldCardsTracker: oldCardsTracker,
      };
    } catch (error) {
      var errorReport = errors.report("cardsReader.getVersion1_0CardsTracker", error, {
        oldCardsTrackerData: oldCardsTrackerData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts CardsPreset from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldCardsPresetData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0CardsPreset: function (oldCardsPresetData) {
    try {
      console.log("Called: cardsReader.getVersion1_0CardsPreset");
      var shouldRemoveUsedCards;
      var oldCardsPreset = {};

      var headerRowIndex = -1;
      for (var rowIndex = 0; rowIndex < oldCardsPresetData.length; rowIndex++) {
        var colIndex = oldCardsPresetData[rowIndex].indexOf(
          "Remove used cards from the pool",
        );
        if (colIndex !== -1) {
          shouldRemoveUsedCards =
            oldCardsPresetData[rowIndex][colIndex - 1] === "TRUE" ||
            oldCardsPresetData[rowIndex][colIndex - 1] === "true" ||
            oldCardsPresetData[rowIndex][colIndex - 1] === true;
          headerRowIndex = rowIndex + 2;
          break;
        }
      }

      if (headerRowIndex === -1 || !oldCardsPresetData[headerRowIndex]) {
        console.log(
          `Could not find the preset header row in Card Preset sheet`,
        );
        return {
          success: false,
          message: "Could not find the preset header row in Card Preset sheet",
        };
      }

      var row = oldCardsPresetData[headerRowIndex];
      var oldCardPresetNameIdxs = row
        .map(function (cell, idx) {
          return String(cell || "").trim() !== "" ? idx : -1;
        })
        .filter(function (idx) {
          return idx !== -1;
        });

      var presetOrder = presetUtils.resolvePresetOrder(
        oldCardPresetNameIdxs.map(function (colIdx) {
          return row[colIdx];
        }),
        presetUtils.templatePresetNames,
      );
      var orderBySourceIndex = {};
      presetOrder.indices.forEach(function (sourceIndex, slot) {
        orderBySourceIndex[sourceIndex] = slot + 1;
      });

      var rowType = "cards";
      for (
        var rowIdx = headerRowIndex + 1;
        rowIdx < oldCardsPresetData.length;
        rowIdx++
      ) {
        if (
          oldCardsPresetData[rowIdx].some(
            (cell) => cell === "Cards to remove from the pool",
          )
        ) {
          rowType = "remove";
          continue;
        }
        oldCardPresetNameIdxs.forEach(function (colIdx, sourceIndex) {
          if (
            oldCardsPresetData[rowIdx][colIdx + 1] &&
            oldCardsPresetData[rowIdx][colIdx + 1].trim() !== ""
          ) {
            var presetName = row[colIdx];
            if (!oldCardsPreset.hasOwnProperty(presetName)) {
              oldCardsPreset[presetName] = {
                cards: [],
                remove: [],
                order: orderBySourceIndex[sourceIndex],
              };
            }
            oldCardsPreset[presetName][rowType].push(
              oldCardsPresetData[rowIdx][colIdx + 1],
            );
          }
        });
      }

      return {
        success: true,
        message: "Cards preset processed successfully",
        oldCardsPreset: oldCardsPreset,
        shouldRemoveUsedCards: shouldRemoveUsedCards,
      };
    } catch (error) {
      var errorReport = errors.report("cardsReader.getVersion1_0CardsPreset", error, {
        oldCardsPresetData: oldCardsPresetData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts CardsLevel from a v1.0 sheet's values.
   * @param {Array<Array<*>>} oldCardsLevelData
   * @param {Array<Array<*>>} oldCardSlotsData
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion1_0CardsLevel: function (oldCardsLevelData, oldCardSlotsData) {
    try {
      console.log("Called: cardsReader.getVersion1_0CardsLevel");
      var oldCardSlots =
        oldCardSlotsData && oldCardSlotsData[0] && oldCardSlotsData[0][0]
          ? oldCardSlotsData[0][0]
          : null;

      if (!oldCardSlots) {
        console.log(`Error getting old card slots`);
        return {
          success: false,
          message: "Error getting old card slots",
        };
      }

      var oldCardsLevel = {};
      oldCardsLevelData.forEach(function (row) {
        var cardName = row[0];
        if (cardName && String(cardName).trim() !== "") {
          oldCardsLevel[cardName] = [row[1], row[2]];
        }
      });

      return {
        success: true,
        message: "Cards level processed successfully",
        oldCardsLevel: oldCardsLevel,
        oldCardSlots: oldCardSlots,
      };
    } catch (error) {
      var errorReport = errors.report("cardsReader.getVersion1_0CardsLevel", error, {
        oldCardsLevelData: oldCardsLevelData,
        oldCardSlotsData: oldCardSlotsData,
      });
      return errors.fail(errorReport);
    }
  },
};
