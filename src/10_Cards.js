const cards = {
  importData: function (versionDifference) {
    function importCardsData(versionDifference) {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        if (!newSpreadsheet) {
          console.log(`New spreadsheet not found`);
          return {
            success: false,
            message: "New spreadsheet not found",
          };
        }
        var newSheetID = newSpreadsheet.spreadsheetId;

        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        if (!oldSpreadsheet) {
          console.log(`Old spreadsheet not found`);
          return {
            success: false,
            message: "Old spreadsheet not found",
          };
        }
        var oldSheetID = oldSpreadsheet.spreadsheetId;
        
        var getVersionFunction = convertVersionFunctions[versionDifference];
        if (!getVersionFunction) {
          console.log(`Unsupported version difference: ${versionDifference}`);
          return {
            success: false,
            message: `Unsupported version difference: ${versionDifference}`,
          };
        }
        var result = getVersionFunction();
        if (!result || !result.success) {
          console.log(`Error processing cards data: ${result.message}`);
          return result;
        }
        
        var oldCardsLevels = result.oldCardsLevels || [];
        var oldCardSlots = result.oldCardSlots || "";
        var oldCardsPresets = result.oldCardsPresets || {};
        var shouldRemoveUsedCards = result.shouldRemoveUsedCards || true;
        var result = updateCardsLevels(
          newSheetID,
          "Master Sheet",
          oldCardsLevels,
          oldCardSlots
        );
        if (!result || !result.success) {
          console.log(`Error updating cards levels: ${result.message}`);
          return result;
        }

        var batchUpdate = result.batchUpdate || [];

        var result = updateCardsPresets(
          newSheetID,
          "Card Preset",
          oldCardsPresets,
          shouldRemoveUsedCards
        );
        if (!result || !result.success) {
          console.log(`Error updating cards presets: ${result.message}`);
          return result;
        }
        batchUpdate = batchUpdate.concat(result.batchUpdate || []);

        var updateResult = SheetsAPI.batchUpdateValues(
          newSheetID,
          batchUpdate
        );
        if (!updateResult) {
          console.log(
            `Error applying batch updates to new spreadsheet`
          );
          return {
            success: false,
            message: "Error applying batch updates to new spreadsheet",
          };
        }
        return {
          success: true,
          message: `Cards data imported successfully`,
        };

      } catch (error) {
        console.log(`Error importing cards data: ${error.toString()}`);
        return {
          success: false,
          message: `Error importing cards data: ${error.message}`,
        };
      }
    }

    function version17(newSheetID, oldSheetID) {
      var oldSheetData = SheetsAPI.getDataRange(oldSheetID, "Card Preset");
      if (!oldSheetData) {
        console.log(`Error getting old card preset sheet data`);
        return {
          success: false,
          message: "Error getting old card preset sheet data",
        };
      }
      var headerValues = SheetsAPI.getValues(newSheetID, "_IDS!1:1");
      if (!headerValues || headerValues.length === 0) {
        console.log(`Could not read header row from _IDS sheet`);
        return {
          success: false,
          message: "Could not read header row from _IDS sheet",
        };
      }

      var headerRow = headerValues[0];
      var importCardsColStart = headerRow.indexOf("Cards");
      if (importCardsColStart === -1) {
        console.log(`Cards column not found in header`);
        return {
          success: false,
          message: "Cards column not found in header",
        };
      }

      var colStart = shared.columnToLetter(importCardsColStart + 1);
      var colEnd = shared.columnToLetter(importCardsColStart + 3);
      var oldCardsLevelsData = SheetsAPI.getValues(
        newSheetID,
        "_IDS!" + colStart + "2:" + colEnd
      );
      if (!oldCardsLevelsData) {
        console.log(`Error getting old cards levels data`);
        return {
          success: false,
          message: "Error getting old cards levels data",
        };
      }

      var oldCardsLevels = oldCardsLevelsData.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );
      
      var oldCardSlots = SheetsAPI.getValue(oldSheetID, "EXPORT!C2");
      if (!oldCardSlots) {
        console.log(`Error getting old card slots`);
        return {
          success: false,
          message: "Error getting old card slots",
        };
      }

      var shouldRemoveUsedCards
      var oldCardsPresets = {};
      for (var rowIndex = 0; rowIndex < oldSheetData.length; rowIndex++) {
        var row = oldSheetData[rowIndex];
        var colIndex = row.indexOf("Remove used cards from the pool");
        if (colIndex !== -1) {
          shouldRemoveUsedCards = row[colIndex - 1] === "TRUE" || row[colIndex - 1] === "true" || row[colIndex - 1] === true;
        }
        var oldCardPresetNameIdxs = row
          .map(function (cell, idx) {
            return String(cell || "").trim() !== "" ? idx : -1;
          })
          .filter(function (idx) {
            return idx !== -1;
          });
        if (row[oldCardPresetNameIdxs[0]] === "Farming" && row[oldCardPresetNameIdxs[1]] === "Tourney") {
          var rowType = "cards"
          for (var rowIdx = rowIndex + 1; rowIdx < oldSheetData.length; rowIdx++) {
            if (oldSheetData[rowIdx].some((cell) => cell === "Cards to remove from the pool")) {
              rowType = "remove";
              continue;
            }
            oldCardPresetNameIdxs.forEach(function (colIdx, orderIndex) {
              if (oldSheetData[rowIdx][colIdx + 1] && oldSheetData[rowIdx][colIdx + 1].trim() !== "") {
                var presetName = row[colIdx];
                if (!oldCardsPresets.hasOwnProperty(presetName)) {
                  oldCardsPresets[presetName] = {cards: [], remove: [], order: orderIndex + 1};
                }
                oldCardsPresets[presetName][rowType].push(oldSheetData[rowIdx][colIdx + 1]);
              }
            });
          }
          break;
        };
      }
      return {
        success: true,
        message: `Cards presets processed successfully`,
        oldCardsLevels: oldCardsLevels,
        oldCardSlots: oldCardSlots,
        oldCardsPresets: oldCardsPresets,
        shouldRemoveUsedCards: shouldRemoveUsedCards,
      };
    }

    function version10() {
      try {
        var newSpreadsheet = spreadsheets("newSpreadsheet");
        var newSheetID = newSpreadsheet.spreadsheetId;
        
        var oldSpreadsheet = spreadsheets("oldSpreadsheet");
        var oldSheetID = oldSpreadsheet.spreadsheetId;
        
        var headerValues = SheetsAPI.getValues(newSheetID, "_IDS!1:1");
        if (!headerValues || headerValues.length === 0) {
          console.log(`Could not read header row from _IDS sheet`);
          return {
            success: false,
            message: "Could not read header row from _IDS sheet",
          };
        }

        var headerRow = headerValues[0];
        var importCardsColStart = headerRow.indexOf("Cards");
        if (importCardsColStart === -1) {
          console.log(`Cards column not found in header`);
          return {
            success: false,
            message: "Cards column not found in header",
          };
        }

        var colStart = shared.columnToLetter(importCardsColStart + 1);
        var colEnd = shared.columnToLetter(importCardsColStart + 3);
        var oldCardsLevelsData = SheetsAPI.getValues(
          newSheetID,
          "_IDS!" + colStart + "2:" + colEnd
        );
        if (!oldCardsLevelsData) {
          console.log(`Error getting old cards levels data`);
          return {
            success: false,
            message: "Error getting old cards levels data",
          };
        }

        var oldCardsLevels = oldCardsLevelsData.filter((row) =>
          row.some(
            (cell) =>
              cell !== null &&
              cell !== undefined &&
              String(cell || "").trim() !== ""
          )
        );

        var oldCardSlots = SheetsAPI.getValue(oldSheetID, "EXPORT!C2");
        if (!oldCardSlots) {
          console.log(`Error getting old card slots`);
          return {
            success: false,
            message: "Error getting old card slots",
          };
        }

        var importCardsPresetsColStart = headerRow.indexOf("Cards Presets");
        if (importCardsPresetsColStart === -1) {
          console.log(`Cards Presets column not found in header`);
          return {
            success: false,
            message: "Cards Presets column not found in header",
          };
        }
        var colStart = shared.columnToLetter(
          importCardsPresetsColStart + 1
        );
        var colEnd = shared.columnToLetter(
          importCardsPresetsColStart + 5
        );
        var oldCardsPresetsData = SheetsAPI.getValues(
          newSheetID,
          "_IDS!" + colStart + "2:" + colEnd
        );
        var oldCardsPresets = {};
        if (oldCardsPresetsData.length > 1) {
          var headers = oldCardsPresetsData[0];
          var dataRows = oldCardsPresetsData.slice(1);
          
          headers.forEach(function(header, colIndex) {
            if (header && String(header).trim() !== "") {
              oldCardsPresets[header] = {
                cards: dataRows.map(function(row) {
                  return row[colIndex] || "";
                }).filter(function(cell) {
                  return cell !== null && cell !== undefined && String(cell).trim() !== "";
                }),
                remove: [],
                order: colIndex + 1
              };
            }
          });
        }
        return {
          success: true,
          message: `oldCardsLevels and oldCardsPresets processed successfully`,
          oldCardSlots: oldCardSlots,
          oldCardsLevels: oldCardsLevels,
          oldCardsPresets: oldCardsPresets,
        };
      } catch (error) {
        console.log(
          `Error processing cards presets: ${error.toString()}`
        );
        return {
          success: false,
          message: `Error processing cards presets: ${error.message}`,
        };
      }
    }

    function updateCardsLevels(
      newSheetID,
      sheetName,
      oldCardsLevels,
      oldCardSlots
    ) {
      // Get sheet data using Sheets API
      var sheetData = SheetsAPI.getDataRange(newSheetID, sheetName);
      if (!sheetData) {
        console.log(`Error getting cards master sheet data`);
        return {
          success: false,
          message: "Error getting cards master sheet data",
        };
      }

      if (sheetData.length < 2) {
        console.log(`Master Sheet has no data or only header row`);
        return {
          success: false,
          message: "Master Sheet has no data or only header row",
        };
      }

      var headerRow = sheetData[0];
      var newCardNameCol = headerRow.indexOf("Card Name");
      if (newCardNameCol === -1) {
        console.log(`Card Name column not found in Master Sheet`);
        return {
          success: false,
          message: "Card Name column not found in Master Sheet",
        };
      }

      var newCardsLevels = [];
      for (var i = 1; i < sheetData.length; i++) {
        var row = sheetData[i];
        if (row.length > newCardNameCol + 2) {
          newCardsLevels.push([
            row[newCardNameCol] || "",
            row[newCardNameCol + 1] || "",
            row[newCardNameCol + 2] || "",
          ]);
        }
      }

      var oldCards = {};
      oldCardsLevels.forEach(function (row) {
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
    }

    function updateCardsPresets(newSheetID, sheetName, oldCardsPresets, shouldRemoveUsedCards) {
      var sheetData = SheetsAPI.getDataRange(newSheetID, sheetName);
      if (!sheetData) {
        console.log(`Error getting cards preset sheet data`);
        return {
          success: false,
          message: "Error getting cards preset sheet data",
        };
      }

      if (sheetData.length < 3) {
        console.log(`Card Preset sheet has no data or only header row`);
        return {
          success: false,
          message: "Card Preset sheet has no data or only header row",
        };
      }

      // Find the header row dynamically by looking for "Farming" and "Tourney"
      var headerRowIndex = -1;
      var headerColIndices = [];
      var batchUpdate = [];
      
      for (var row = 0; row < sheetData.length; row++) {
        if (sheetData[row].indexOf("Remove used cards from the pool") !== -1) {
          var removeUsedCardsCol = shared.columnToLetter(sheetData[row].indexOf("Remove used cards from the pool"));
          batchUpdate.push({
            range: sheetName + "!" + removeUsedCardsCol + (row + 1),
            values: [[shouldRemoveUsedCards]],
          });
          continue;
        }
        // Filter out empty cells from the row
        var nonEmptyCells = sheetData[row].filter(function(cell) {
          return cell !== null && cell !== undefined && String(cell).trim() !== "";
        });
        
        // Check if first 2 non-empty values are "Farming" and "Tourney"
        if (nonEmptyCells.length >= 2 && nonEmptyCells[0] === "Farming" && nonEmptyCells[1] === "Tourney") {
          headerRowIndex = row;
          
          // Use forEach to get the actual column indices for each non-empty header
          nonEmptyCells.forEach(function(header) {
            var colIndex = sheetData[row].indexOf(header);
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
        console.log(`Expected 5 preset columns but found ${headerColIndices.length}`);
        return {
          success: false,
          message: `Expected 5 preset columns but found ${headerColIndices.length}`,
        };
      }

      // Process each preset from oldCardsPresets object
      Object.keys(oldCardsPresets).forEach(function(presetName) {
        var presetData = oldCardsPresets[presetName];
        
        // Skip if this is not a preset (e.g., "order" property)
        if (!presetData.order) return;
        
        // Use the order to get the actual column index (order 1-5 maps to headerColIndices[0-4])
        var orderIndex = presetData.order - 1;
        if (orderIndex >= 0 && orderIndex < headerColIndices.length) {
          var colIndex = headerColIndices[orderIndex];
          
          // Update preset header
          var headerCell = shared.columnToLetter(colIndex + 1) + (headerRowIndex + 1);
          batchUpdate.push({
            range: sheetName + "!" + headerCell,
            values: [[presetName]],
          });

          // Find where "cards" section starts and "remove" section starts
          var cardsStartRow = headerRowIndex + 1; // Start after preset names row
          var removeStartRow = -1;
          
          for (var row = headerRowIndex + 1; row < sheetData.length; row++) {
            if (sheetData[row].some(cell => String(cell).includes("Cards to remove from the pool"))) {
              removeStartRow = row + 1; // Start after the "Cards to remove" header
              break;
            }
          }
          
          // If no "remove" section found, assume all data is cards
          if (removeStartRow === -1) {
            removeStartRow = sheetData.length; // No remove section
          }

          // Update cards section
          if (presetData.cards && presetData.cards.length > 0) {
            var cardsData = presetData.cards.map(function(card) {
              return [card];
            });
            
            var startCell = shared.columnToLetter(colIndex + 2) + (cardsStartRow + 1);
            var endCell = shared.columnToLetter(colIndex + 2) + (cardsStartRow + cardsData.length);
            
            batchUpdate.push({
              range: sheetName + "!" + startCell + ":" + endCell,
              values: cardsData,
            });
          }

          // Update remove section
          if (presetData.remove && presetData.remove.length > 0 && removeStartRow !== -1) {
            var removeData = presetData.remove.map(function(card) {
              return [card];
            });
            
            var startCell = shared.columnToLetter(colIndex + 2) + (removeStartRow + 1);
            var endCell = shared.columnToLetter(colIndex + 2) + (removeStartRow + removeData.length);
            
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
          message: `Cards presets updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      
      return {
        success: true,
        message: `No updates needed for cards presets`,
      };
    }

    var convertVersionFunctions = {
      "v1.7": version17,
      "v1.0": version10,
    };

    return importCardsData(versionDifference);
  },

  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = [
      "v1.7",
      "v1.0"
    ];
    
    var sortedThresholds = versionCompatibility.slice().sort(function(a, b) {
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