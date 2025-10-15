const cards = {
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
          masterSheetData
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
        var shouldRemoveUsedCards = data.shouldRemoveUsedCards || true;
        var presetResult = this.updateCardsPreset(
          "Card Preset",
          oldCardsPreset,
          shouldRemoveUsedCards,
          cardPresetsData
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
          cardTrackerData
        );
        if (!trackerResult || !trackerResult.success) {
          console.log(`Error updating cards tracker: ${trackerResult.message}`);
          return trackerResult;
        }
        batchUpdate = batchUpdate.concat(trackerResult.batchUpdate || []);
      }

      // Add import status update to batch if there were data updates
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }

      // Always add ID updates
      shared.addIDUpdatesToBatch(batchUpdate, "Cards", newSheetID, idsData, data.idMasterID);

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

  updateCardsLevels: function (
    sheetName,
    oldCardsLevel,
    oldCardSlots,
    masterSheetData
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

      var newCardsLevels = [];
      for (var i = 1; i < masterSheetData.length; i++) {
        var row = masterSheetData[i];
        if (row.length > newCardNameCol + 2) {
          newCardsLevels.push([
            row[newCardNameCol] || "",
            row[newCardNameCol + 1] || "",
            row[newCardNameCol + 2] || "",
          ]);
        }
      }

      var oldCards = {};
      oldCardsLevel.forEach(function (row) {
        var oldCardName = row[0];
        var oldLevel = row[1];
        var oldMastery = row[2];
        if (oldCardName) {
          oldCards[oldCardName] = [oldLevel, oldMastery];
        }
      });

      var newCards = [];
      newCardsLevels.forEach(function (row) {
        var newCardName = row[0];
        if (newCardName === "Card Slot (Gems)") {
          newCards.push([oldCardSlots, ""]);
        } else if (oldCards.hasOwnProperty(newCardName)) {
          newCards.push(oldCards[newCardName]);
        } else {
          newCards.push([row[1], row[2]]);
        }
      });

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
    cardPresetsData
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
        if (
          cardPresetsData[row].indexOf("Remove used cards from the pool") !== -1
        ) {
          var removeUsedCardsCol = shared.columnToLetter(
            cardPresetsData[row].indexOf("Remove used cards from the pool")
          );
          batchUpdate.push({
            range: sheetName + "!" + removeUsedCardsCol + (row + 1),
            values: [[shouldRemoveUsedCards]],
          });
          continue;
        }
        var nonEmptyCells = cardPresetsData[row].filter(function (cell) {
          return (
            cell !== null && cell !== undefined && String(cell).trim() !== ""
          );
        });

        if (
          nonEmptyCells.length >= 2 &&
          nonEmptyCells[0] === "Farming" &&
          nonEmptyCells[1] === "Tourney"
        ) {
          headerRowIndex = row;

          nonEmptyCells.forEach(function (header) {
            var colIndex = cardPresetsData[row].indexOf(header);
            if (colIndex !== -1) {
              headerColIndices.push(colIndex);
            }
          });
          break;
        }
      }

      if (headerRowIndex === -1) {
        console.log(`Could not find header row with "Farming" and "Tourney"`);
        return {
          success: false,
          message: "Could not find header row with Farming and Tourney",
        };
      }

      if (headerColIndices.length < 5) {
        console.log(
          `Expected 5 preset columns but found ${headerColIndices.length}`
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
        if (orderIndex >= 0 && orderIndex < headerColIndices.length) {
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
                String(cell).includes("Cards to remove from the pool")
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

      console.log("OLD:", JSON.stringify(oldCardsTracker));
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
      console.log("BATCH:", JSON.stringify(batchUpdate));
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

  version10: function () {
    try {
      console.log("Called: cards.version10");
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

      var cardsPresetData = this.getVersion10CardsPreset(oldCardsPresetData);
      if (!cardsPresetData || !cardsPresetData.success) {
        return cardsPresetData;
      }

      var cardsTrackerData = this.getVersion10CardsTracker(oldCardsTrackerData);
      if (!cardsTrackerData || !cardsTrackerData.success) {
        return cardsTrackerData;
      }

      var cardsLevelData = this.getVersion10CardsLevel(
        oldCardsLevelData,
        oldCardSlotsData
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
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10CardsTracker: function (oldCardsTrackerData) {
    try {
      console.log("Called: cards.getVersion10CardsTracker");
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
          priorityColIndex
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
      console.log("Error in getVersion10CardsTracker: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10CardsTracker: " + error.message,
      };
    }
  },

  getVersion10CardsPreset: function (oldCardsPresetData) {
    try {
      console.log("Called: cards.getVersion10CardsPreset");
      var shouldRemoveUsedCards;
      var oldCardsPreset = {};
      for (var rowIndex = 0; rowIndex < oldCardsPresetData.length; rowIndex++) {
        var row = oldCardsPresetData[rowIndex];
        var colIndex = row.indexOf("Remove used cards from the pool");
        if (colIndex !== -1) {
          shouldRemoveUsedCards =
            row[colIndex - 1] === "TRUE" ||
            row[colIndex - 1] === "true" ||
            row[colIndex - 1] === true;
        }
        var oldCardPresetNameIdxs = row
          .map(function (cell, idx) {
            return String(cell || "").trim() !== "" ? idx : -1;
          })
          .filter(function (idx) {
            return idx !== -1;
          });
        if (
          row[oldCardPresetNameIdxs[0]] === "Farming" &&
          row[oldCardPresetNameIdxs[1]] === "Tourney"
        ) {
          var rowType = "cards";
          for (
            var rowIdx = rowIndex + 1;
            rowIdx < oldCardsPresetData.length;
            rowIdx++
          ) {
            if (
              oldCardsPresetData[rowIdx].some(
                (cell) => cell === "Cards to remove from the pool"
              )
            ) {
              rowType = "remove";
              continue;
            }
            oldCardPresetNameIdxs.forEach(function (colIdx, orderIndex) {
              if (
                oldCardsPresetData[rowIdx][colIdx + 1] &&
                oldCardsPresetData[rowIdx][colIdx + 1].trim() !== ""
              ) {
                var presetName = row[colIdx];
                if (!oldCardsPreset.hasOwnProperty(presetName)) {
                  oldCardsPreset[presetName] = {
                    cards: [],
                    remove: [],
                    order: orderIndex + 1,
                  };
                }
                oldCardsPreset[presetName][rowType].push(
                  oldCardsPresetData[rowIdx][colIdx + 1]
                );
              }
            });
          }
          break;
        }
      }
      return {
        success: true,
        message: "Cards preset processed successfully",
        oldCardsPreset: oldCardsPreset,
        shouldRemoveUsedCards: shouldRemoveUsedCards,
      };
    } catch (error) {
      console.log("Error in getVersion10CardsPreset: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10CardsPreset: " + error.message,
      };
    }
  },

  getVersion10CardsLevel: function (oldCardsLevelData, oldCardSlotsData) {
    try {
      console.log("Called: cards.getVersion10CardsLevel");
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

      var oldCardsLevel = oldCardsLevelData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      return {
        success: true,
        message: "Cards level processed successfully",
        oldCardsLevel: oldCardsLevel,
        oldCardSlots: oldCardSlots,
      };
    } catch (error) {
      console.log("Error in getVersion10CardsLevel: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10CardsLevel: " + error.message,
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
