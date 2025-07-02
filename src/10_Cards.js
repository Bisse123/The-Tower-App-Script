const cards = {
  convertVersionFunctions: {},

  importData: function (newSheetID, oldSheetID) {
    function importCardsData(newSheetID, oldSheetID) {
      try {
        // Get new cards version using SheetsAPI
        var newCardsVersion = SheetsAPI.getValue(newSheetID, "EXPORT!A1");
        if (!newCardsVersion) {
          console.log("Error getting new cards version");
          return {
            success: false,
            message: "Error getting new cards version"
          };
        }

        // Get old cards version using SheetsAPI
        var oldCardsVersion = SheetsAPI.getValue(oldSheetID, "EXPORT!A1");
        if (!oldCardsVersion) {
        console.log("Error getting old cards version");
        return {
          success: false,
          message: "Error getting old cards version"
        };
        }

        var versionCheck = shared.compareVersions(
          oldCardsVersion,
          newCardsVersion
        );
        if (versionCheck === 0) {
          console.log("Same Version");

          // Get header row from _IDS sheet using Sheets API
          var headerRowData;
          try {
          var headerRowData = SheetsAPI.getValues(newSheetID, "_IDS!1:1")[0] || [];
          } catch (error) {
            console.log("Error getting header row: " + error.toString());
            return {
              success: false,
              message: "Error getting header row: " + error.message
            };
          }

          var importCardsColStart = headerRowData.indexOf("Cards");
          if (importCardsColStart === -1) {
          console.log("Cards column not found in header");
          return {
            success: false,
            message: "Cards column not found in header"
          };
          }

          var colStart = shared.columnToLetter(importCardsColStart + 1);
          var colEnd = shared.columnToLetter(importCardsColStart + 3);
          var oldCardsLevelsData = SheetsAPI.getValues(
            newSheetID,
            "_IDS!" + colStart + "2:" + colEnd
          );
          if (!oldCardsLevelsData) {
            console.log("Error getting old cards levels data");
            return {
              success: false,
              message: "Error getting old cards levels data"
            };
          }

          var oldCardsLevels = oldCardsLevelsData.filter((row) =>
            row.some(
              (cell) =>
                cell !== null && cell !== undefined && String(cell || '').trim() !== ""
            )
          );

          var oldCardSlots = SheetsAPI.getValue(oldSheetID, "EXPORT!C2");
          if (!oldCardSlots) {
            console.log("Error getting old card slots");
            return {
              success: false,
              message: "Error getting old card slots"
            };
          }

          var importCardsPresetsColStart = headerRowData.indexOf("Cards Presets");
          if (importCardsPresetsColStart !== -1) {
            try {
              var colStart = shared.columnToLetter(importCardsPresetsColStart + 1);
              var colEnd = shared.columnToLetter(importCardsPresetsColStart + 5);
              var oldCardsPresetsData = SheetsAPI.getValues(
                newSheetID,
                "_IDS!" + colStart + "2:" + colEnd
              );
              var oldCardsPresets = oldCardsPresetsData.filter((row) =>
                row.some(
                  (cell) =>
                    cell !== null &&
                    cell !== undefined &&
                    String(cell || '').trim() !== ""
                )
              );

              var result = updateCardsLevels(
                newSheetID,
                "Master Sheet",
                oldCardsLevels,
                oldCardSlots
              );
              if (!result || !result.success) {
                console.log("Error updating cards levels: " + result.message);
                return result;
              }
              var result = updateCardsPresets(
                newSheetID,
                "Card Preset",
                oldCardsPresets
              );
              if (!result || !result.success) {
                console.log("Error updating cards presets: " + result.message);
                return result;
              }
              console.log("Cards data imported successfully");
              return {
                success: true,
                message: "Cards data imported successfully"
              };
            } catch (error) {
              console.log("Error processing cards presets: " + error.toString());
              return {
                success: false,
                message: "Error processing cards presets: " + error.message
              };
            }
          } else {
            var result = updateCardsLevels(
              newSheetID,
              "Master Sheet",
              oldCardsLevels,
              oldCardSlots
            );
            if (!result || !result.success) {
              console.log("Error updating cards levels: " + result.message);
              return result;
            }
            console.log("Cards levels updated successfully");
            return {
              success: true,
              message: "Cards levels updated successfully"
            };
          }
        }
        // else {// Else do something to convert old version to new one (Future me problem)
        // }
      } catch (error) {
        console.log("Error importing cards data: " + error.toString());
        return { success: false, message: "Error importing cards data" };
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
      console.log("Error getting cards master sheet data");
      return {
        success: false,
        message: "Error getting cards master sheet data"
      };
      }

      if (sheetData.length < 2) {
        console.log("Master Sheet has no data or only header row");
        return {
          success: false,
          message: "Master Sheet has no data or only header row",
        };
      }

      var headerRow = sheetData[0];
      var newCardNameCol = headerRow.indexOf("Card Name");
      if (newCardNameCol === -1) {
      console.log("Card Name column not found in Master Sheet");
      return {
        success: false,
        message: "Card Name column not found in Master Sheet"
      };
      }

      // Get card data starting from row 2
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
        SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        return {
          success: true,
          message: "Cards levels updated successfully"
        };
      }
      console.log("No updates needed for cards levels");
      return {
        success: true,
        message: "No updates needed for cards levels"
      };
    }

    function updateCardsPresets(newSheetID, sheetName, oldCardsPresets) {
      // Get sheet data using Sheets API
      var sheetData = SheetsAPI.getDataRange(newSheetID, sheetName);
      if (!sheetData) {
      console.log("Error getting cards preset sheet data");
      return {
        success: false,
        message: "Error getting cards preset sheet data"
      };
      }

      if (sheetData.length < 3) {
      console.log("Master Sheet has no data or only header row");
      return {
        success: false,
        message: "Master Sheet has no data or only header row",
      };
    }

      var headerRow = sheetData[1]; // Row 2 contains the headers
      var newCardPresetNameIdxs = headerRow
        .map(function (cell, idx) {
          return String(cell || '').trim() !== "" ? idx : -1;
        })
        .filter(function (idx) {
          return idx !== -1;
        });

      var oldCardsPresetsHeaders = oldCardsPresets[0];
      var oldCardsPresetsCards = oldCardsPresets.slice(1);

      var batchUpdate = [];
      oldCardsPresetsHeaders.forEach(function (header, headerIdx) {
        var colIdx = newCardPresetNameIdxs[headerIdx];
        if (colIdx !== undefined && colIdx >= 0) {
          // Update header
          var headerCell = shared.columnToLetter(colIdx + 1) + "2";
          batchUpdate.push({
            range: sheetName + "!" + headerCell,
            value: header,
          });

          // Update preset cards
          var newCardsPresetsCards = oldCardsPresetsCards.map(function (row) {
            return [row[headerIdx]];
          });

          if (newCardsPresetsCards.length > 0) {
            var startCell = shared.columnToLetter(colIdx + 2) + "3";
            var endCell =
              shared.columnToLetter(colIdx + 2) +
              (2 + newCardsPresetsCards.length);
            batchUpdate.push({
              range: sheetName + "!" + startCell + ":" + endCell,
              values: newCardsPresetsCards,
            });
          }
        }
      });

    if (batchUpdate.length !== 0) {
        SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        console.log("Cards presets updated successfully");
        return {
          success: true,
          message: "Cards presets updated successfully"
        };
      }
      console.log("No updates needed for cards presets");
      return {
        success: true,
        message: "No updates needed for cards presets"
      };
    }
    return importCardsData(newSheetID, oldSheetID);
  },

  isCompatibleVersion: function (oldVersion) {
    return this.convertVersionFunctions[oldVersion];
  },
};
