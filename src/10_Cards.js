const cards = {
  // #region Export Functions
  exportData: function (versionDifference) {
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

      var oldDataResult = getVersionFunction();
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
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting cards data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: cards.importData");
      var newSpreadsheet = spreadsheets("Cards newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }

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

      // Only update cards levels if key exists
      if (
        data.hasOwnProperty("oldCardsLevel") &&
        data.hasOwnProperty("oldCardSlots")
      ) {
        var oldCardsLevel = data.oldCardsLevel;
        var oldCardSlots = data.oldCardSlots || "";
        var levelsResult = this.updateCardsLevels(
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

      // Only update cards preset if key exists
      if (data.hasOwnProperty("oldCardsPreset")) {
        var oldCardsPreset = data.oldCardsPreset;
        var shouldRemoveUsedCards = data.hasOwnProperty("shouldRemoveUsedCards")
          ? data.shouldRemoveUsedCards
          : true;
        var presetResult = this.updateCardsPreset(
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
        var trackerResult = this.updateCardsTracker(
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

      // Always add ID updates
      shared.addIDUpdatesToBatch(
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
      console.log(`Error importing cards data: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing cards data: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Update Functions
  updateCardsLevels: function (
    sheetName,
    oldCardsLevel,
    oldCardSlots,
    masterSheetData,
  ) {
    try {
      console.log("Called: cards.updateCardsLevels");
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
        var startCol = shared.columnToLetter(newCardNameCol + 2);
        var endCol = shared.columnToLetter(newCardNameCol + 3);
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
      console.log("Error in updateCardsLevels: " + error.toString());
      return {
        success: false,
        message: "Error in updateCardsLevels: " + error.message,
      };
    }
  },

  updateCardsPreset: function (
    sheetName,
    oldCardsPreset,
    shouldRemoveUsedCards,
    cardPresetsData,
  ) {
    try {
      console.log("Called: cards.updateCardsPreset");
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
            shared.columnToLetter(removeUsedCardsIndex) +
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
          shared.columnToLetter(colIndex + 1) + (headerRowIndex + 1);
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
            shared.columnToLetter(colIndex + 2) + (cardsStartRow + 1);
          var endCell =
            shared.columnToLetter(colIndex + 2) +
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
            shared.columnToLetter(colIndex + 2) + (removeStartRow + 1);
          var endCell =
            shared.columnToLetter(colIndex + 2) +
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
      console.log("Error in updateCardsPreset: " + error.toString());
      return {
        success: false,
        message: "Error in updateCardsPreset: " + error.message,
      };
    }
  },

  updateCardsTracker: function (sheetName, oldCardsTracker, cardTrackerData) {
    try {
      console.log("Called: cards.updateCardsTracker");
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
                    shared.columnToLetter(progressColIndex + 1) +
                    (rowIdx + 1),
                  values: [[oldData.progress]],
                });
              }
              if (oldData.priority) {
                batchUpdate.push({
                  range:
                    sheetName +
                    "!" +
                    shared.columnToLetter(priorityColIndex + 1) +
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
      console.log("Error in updateCardsTracker: " + error.toString());
      return {
        success: false,
        message: "Error in updateCardsTracker: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version1_0: function () {
    try {
      console.log("Called: cards.version1_0");
      var oldSpreadsheet = spreadsheets("Cards oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

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
      console.log("Error in version1_0: " + error.toString());
      return {
        success: false,
        message: "Error in version1_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Cards Tracker
  getVersion1_0CardsTracker: function (oldCardsTrackerData) {
    try {
      console.log("Called: cards.getVersion1_0CardsTracker");
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
      console.log("Error in getVersion1_0CardsTracker: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0CardsTracker: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Cards Preset
  getVersion1_0CardsPreset: function (oldCardsPresetData) {
    try {
      console.log("Called: cards.getVersion1_0CardsPreset");
      var shouldRemoveUsedCards;
      var oldCardsPreset = {};

      // Every preset name may have been renamed, so the header row is located
      // from the "Remove used cards from the pool" toggle above it, which the
      // template always puts two rows higher.
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

      var presetOrder = shared.resolvePresetOrder(
        oldCardPresetNameIdxs.map(function (colIdx) {
          return row[colIdx];
        }),
        shared.templatePresetNames,
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
      console.log("Error in getVersion1_0CardsPreset: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0CardsPreset: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Cards Level
  getVersion1_0CardsLevel: function (oldCardsLevelData, oldCardSlotsData) {
    try {
      console.log("Called: cards.getVersion1_0CardsLevel");
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
      console.log("Error in getVersion1_0CardsLevel: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion1_0CardsLevel: " + error.message,
      };
    }
  },

  // #endregion
  // #region Parse Saved File
  parseCardsData: function (data) {
    const cardNamesByIndex = {
      0: "Damage",
      1: "Attack Speed",
      2: "Health",
      3: "Health Regen",
      4: "Range",
      5: "Cash",
      6: "Coins",
      7: "Slow Aura",
      10: "Critical Chance",
      11: "Enemy Balance",
      12: "Extra Defense",
      13: "Fortress",
      15: "Free Upgrades",
      16: "Extra Orb",
      18: "Plasma Cannon",
      19: "Critical Coin",
      20: "Wave Skip",
      21: "Intro Sprint",
      22: "Land Mine Stun",
      23: "Recovery Package Chance",
      25: "Death Ray",
      26: "Energy Net",
      27: "Super Tower",
      28: "Second Wind",
      29: "Demon Mode",
      30: "Energy Shield",
      31: "Wave Accelerator",
      32: "Berserker",
      33: "Ultimate Crit",
      34: "Nuke",
      35: "Area of Effect",
    };
    var cardNameIndices = [];
    Object.keys(cardNamesByIndex).forEach(function (index) {
      cardNameIndices[Number(index)] = cardNamesByIndex[index];
    });
    const cardLevel = data.cardLevel || [];
    const cardMasteryUnlocked = data.cardMasteryUnlocked || [];

    const presetOrder = shared.resolvePresetOrder(
      data.presetNames || [],
      shared.templatePresetNames,
    );
    const presetNames = presetOrder.order;
    const presetIndices = presetOrder.indices;

    const presetSlots = data.presetSlots || [];
    const presetCards = data.presetCards || [];
    const slotsUnlocked = data.slotsUnlocked || 0;

    var oldCardsLevel = {};
    cardNameIndices.forEach(function (cardName, i) {
      if (!cardName) return;
      oldCardsLevel[cardName] = [cardLevel[i], cardMasteryUnlocked[i]];
    });

    const numPresets = presetNames.length;
    const slotsPerPreset =
      numPresets > 0 ? Math.floor(presetSlots.length / numPresets) : 0;
    var oldCardsPreset = {};
    presetNames.forEach(function (name, slot) {
      if (!name) return;
      var sourceIndex = presetIndices[slot];
      var slotStart = sourceIndex * slotsPerPreset;
      var cards = [];
      presetSlots
        .slice(slotStart, slotStart + slotsPerPreset)
        .forEach(function (assigned, s) {
          if (assigned) {
            var resolvedName =
              cardNameIndices[presetCards[slotStart + s]] || null;
            if (resolvedName) {
              cards.push(resolvedName);
            }
          }
        });
      oldCardsPreset[name] = {
        cards: cards,
        remove: [],
        order: slot + 1,
      };
    });

    return {
      oldCardsLevel: oldCardsLevel,
      oldCardsPreset: oldCardsPreset,
      oldCardSlots: slotsUnlocked,
      cardNameIndices: cardNameIndices,
    };
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.0": this.version1_0.bind(this),
    };
  },

  // #endregion
  // #region Compatibility Check
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

  // #endregion
};
