const cardsWriter = {

  /**
   * Builds the batch update that writes CardsLevels into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldCardsLevel
   * @param {Object} oldCardSlots
   * @param {Object} masterSheetData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateCardsLevels: function (
    sheetName,
    oldCardsLevel,
    oldCardSlots,
    masterSheetData,
  ) {
    try {
      console.log("Called: cardsWriter.updateCardsLevels");
      if (!masterSheetData) {
        console.log(`Error getting cards master sheet data`);
        return {
          success: false,
          message: "Error getting cards master sheet data",
        };
      }
      if (masterSheetData.length < 2) {
        console.log(`Master Sheet has no data or only header row`);
        return {
          success: false,
          message: "Master Sheet has no data or only header row",
        };
      }

      var headerRow = masterSheetData[0];
      var newCardNameCol = headerRow.indexOf("Card Name");
      if (newCardNameCol === -1) {
        console.log(`Card Name column not found in Master Sheet`);
        return {
          success: false,
          message: "Card Name column not found in Master Sheet",
        };
      }

      var newCards = [];
      for (var i = 1; i < masterSheetData.length; i++) {
        var newCardName = masterSheetData[i][newCardNameCol] || "";
        if (newCardName === "Card Slot (Gems)") {
          newCards.push([oldCardSlots, null]);
        } else if (oldCardsLevel.hasOwnProperty(newCardName)) {
          newCards.push(oldCardsLevel[newCardName]);
        } else {
          newCards.push([null, null]);
        }
      }

      var batchUpdate = [];
      if (newCards.length > 0) {
        var startCol = sheetRefs.columnToLetter(newCardNameCol + 2);
        var endCol = sheetRefs.columnToLetter(newCardNameCol + 3);
        var range =
          sheetName + "!" + startCol + "2:" + endCol + (1 + newCards.length);
        batchUpdate.push({
          range: range,
          values: newCards,
        });
      }
      if (batchUpdate.length !== 0) {
        return {
          success: true,
          message: `Cards levels updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for cards levels`,
      };
    } catch (error) {
      var errorReport = errors.report("cardsWriter.updateCardsLevels", error, {
        sheetName: sheetName,
        oldCardsLevel: oldCardsLevel,
        oldCardSlots: oldCardSlots,
        masterSheetData: masterSheetData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes CardsPreset into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldCardsPreset
   * @param {*} shouldRemoveUsedCards
   * @param {Object} cardPresetsData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateCardsPreset: function (
    sheetName,
    oldCardsPreset,
    shouldRemoveUsedCards,
    cardPresetsData,
  ) {
    try {
      console.log("Called: cardsWriter.updateCardsPreset");
      if (!cardPresetsData) {
        console.log(`Error getting cards preset sheet data`);
        return {
          success: false,
          message: "Error getting cards preset sheet data",
        };
      }
      if (cardPresetsData.length < 3) {
        console.log(`Card Preset sheet has no data or only header row`);
        return {
          success: false,
          message: "Card Preset sheet has no data or only header row",
        };
      }

      var headerRowIndex = -1;
      var headerColIndices = [];
      var batchUpdate = [];

      for (var row = 0; row < cardPresetsData.length; row++) {
        var removeUsedCardsIndex = cardPresetsData[row].indexOf(
          "Remove used cards from the pool",
        );
        if (removeUsedCardsIndex === -1) {
          continue;
        }
        batchUpdate.push({
          range:
            sheetName +
            "!" +
            sheetRefs.columnToLetter(removeUsedCardsIndex) +
            (row + 1),
          values: [[shouldRemoveUsedCards]],
        });
        headerRowIndex = row + 2;
        break;
      }

      if (headerRowIndex === -1 || headerRowIndex >= cardPresetsData.length) {
        console.log(`Could not find "Remove used cards from the pool"`);
        return {
          success: false,
          message: "Could not find Remove used cards from the pool",
        };
      }

      var presetHeaderRow = cardPresetsData[headerRowIndex];
      for (
        var col = 0;
        col < presetHeaderRow.length && headerColIndices.length < 5;
        col++
      ) {
        if (String(presetHeaderRow[col] || "").trim() !== "") {
          headerColIndices.push(col);
        }
      }

      if (headerColIndices.length < 5) {
        console.log(
          `Expected 5 preset columns but found ${headerColIndices.length}`,
        );
        return {
          success: false,
          message: `Expected 5 preset columns but found ${headerColIndices.length}`,
        };
      }

      Object.keys(oldCardsPreset).forEach(function (presetName) {
        var presetData = oldCardsPreset[presetName];

        if (!presetData.order) return;

        var orderIndex = presetData.order - 1;
        if (orderIndex < 0 || orderIndex >= headerColIndices.length) {
          return;
        }
        var colIndex = headerColIndices[orderIndex];

        var headerCell =
          sheetRefs.columnToLetter(colIndex + 1) + (headerRowIndex + 1);
        batchUpdate.push({
          range: sheetName + "!" + headerCell,
          values: [[presetName]],
        });

        var cardsStartRow = headerRowIndex + 1;
        var removeStartRow = -1;

        for (
          var row = headerRowIndex + 1;
          row < cardPresetsData.length;
          row++
        ) {
          if (
            cardPresetsData[row].some((cell) =>
              String(cell).includes("Cards to remove from the pool"),
            )
          ) {
            removeStartRow = row + 1;
            break;
          }
        }

        if (removeStartRow === -1) {
          removeStartRow = cardPresetsData.length;
        }

        if (presetData.cards && presetData.cards.length > 0) {
          var cardsData = presetData.cards.map(function (card) {
            return [card];
          });

          var startCell =
            sheetRefs.columnToLetter(colIndex + 2) + (cardsStartRow + 1);
          var endCell =
            sheetRefs.columnToLetter(colIndex + 2) +
            (cardsStartRow + cardsData.length);

          batchUpdate.push({
            range: sheetName + "!" + startCell + ":" + endCell,
            values: cardsData,
          });
        }

        if (
          presetData.remove &&
          presetData.remove.length > 0 &&
          removeStartRow !== -1
        ) {
          var removeData = presetData.remove.map(function (card) {
            return [card];
          });

          var startCell =
            sheetRefs.columnToLetter(colIndex + 2) + (removeStartRow + 1);
          var endCell =
            sheetRefs.columnToLetter(colIndex + 2) +
            (removeStartRow + removeData.length);

          batchUpdate.push({
            range: sheetName + "!" + startCell + ":" + endCell,
            values: removeData,
          });
        }
      });

      if (batchUpdate.length !== 0) {
        return {
          success: true,
          message: `cards preset updated successfully`,
          batchUpdate: batchUpdate,
        };
      }

      return {
        success: true,
        message: `No updates needed for cards preset`,
      };
    } catch (error) {
      var errorReport = errors.report("cardsWriter.updateCardsPreset", error, {
        sheetName: sheetName,
        oldCardsPreset: oldCardsPreset,
        shouldRemoveUsedCards: shouldRemoveUsedCards,
        cardPresetsData: cardPresetsData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes CardsTracker into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldCardsTracker
   * @param {Object} cardTrackerData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateCardsTracker: function (sheetName, oldCardsTracker, cardTrackerData) {
    try {
      console.log("Called: cardsWriter.updateCardsTracker");
      if (!cardTrackerData) {
        console.log(`Error getting cards tracker sheet data`);
        return {
          success: false,
          message: "Error getting cards tracker sheet data",
        };
      }
      if (cardTrackerData.length < 2) {
        console.log(`Card Tracker sheet has no data or only header row`);
        return {
          success: false,
          message: "Card Tracker sheet has no data or only header row",
        };
      }

      var batchUpdate = [];
      for (var i = 0; i < cardTrackerData.length; i++) {
        var row = cardTrackerData[i];
        var cardNameColIndex = row.indexOf("Card");
        var progressColIndex = row.indexOf("Progress");
        var priorityColIndex = row.indexOf("Priority");
        if (
          cardNameColIndex !== -1 &&
          progressColIndex !== -1 &&
          priorityColIndex !== -1
        ) {
          for (var rowIdx = i + 1; rowIdx < cardTrackerData.length; rowIdx++) {
            var trackerRow = cardTrackerData[rowIdx];
            var cardName = trackerRow[cardNameColIndex];
            if (!cardName || String(cardName).trim() === "") {
              break;
            }
            if (oldCardsTracker.hasOwnProperty(cardName)) {
              var oldData = oldCardsTracker[cardName];
              if (oldData.progress) {
                batchUpdate.push({
                  range:
                    sheetName +
                    "!" +
                    sheetRefs.columnToLetter(progressColIndex + 1) +
                    (rowIdx + 1),
                  values: [[oldData.progress]],
                });
              }
              if (oldData.priority) {
                batchUpdate.push({
                  range:
                    sheetName +
                    "!" +
                    sheetRefs.columnToLetter(priorityColIndex + 1) +
                    (rowIdx + 1),
                  values: [[oldData.priority]],
                });
              }
            }
          }
          break;
        }
      }
      if (batchUpdate.length !== 0) {
        return {
          success: true,
          message: `Cards tracker updated successfully`,
          batchUpdate: batchUpdate,
        };
      }

      return {
        success: true,
        message: `No updates needed for cards tracker`,
      };
    } catch (error) {
      var errorReport = errors.report("cardsWriter.updateCardsTracker", error, {
        sheetName: sheetName,
        oldCardsTracker: oldCardsTracker,
        cardTrackerData: cardTrackerData,
      });
      return errors.fail(errorReport);
    }
  },
};
