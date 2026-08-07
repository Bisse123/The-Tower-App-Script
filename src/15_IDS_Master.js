const master = {
  // #region Export Functions
  exportData: function (versionDifference) {
    try {
      console.log("Called: master.exportData");
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
        message: "IDS Master export completed successfully",
        data: oldDataResult,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting IDS Master data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: master.importData");
      var newSpreadsheet = spreadsheets("IDS Master newSpreadsheet");
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }
      var newSheetID = newSpreadsheet.spreadsheetId;

      var batchUpdate = [];
      var failedUpdates = [];

      // Get IDS sheet data for finding import status range
      var requiredRanges = ["IDS", "Presets Presets"];
      var idsData = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!idsData || !idsData[0] || !idsData[0].values) {
        console.log(`Could not read IDS data from new spreadsheet`);
        return {
          success: false,
          message: "Could not read IDS data from new spreadsheet",
        };
      }

      var idsValues = idsData[0].values;
      var presetsValues = idsData[1].values;

      // Update IDS sheet with exported data if available
      if (data.hasOwnProperty("oldIdsData")) {
        try {
          var idsUpdateResult = this.updateIDSData(data.oldIdsData, idsValues);
          if (
            idsUpdateResult.success &&
            idsUpdateResult.batchUpdate.length > 0
          ) {
            batchUpdate = batchUpdate.concat(idsUpdateResult.batchUpdate);
          } else if (!idsUpdateResult.success) {
            failedUpdates.push({
              sheetType: "IDS",
              message: idsUpdateResult.message,
            });
          }
        } catch (error) {
          console.log(`Error updating IDS data: ${error.toString()}`);
          failedUpdates.push({
            sheetType: "IDS",
            message: "Error updating IDS data: " + error.message,
          });
        }
      }

      if (data.hasOwnProperty("oldPresetsData")) {
        try {
          var presetsUpdateResult = this.updatePresetsData(
            data.oldPresetsData,
            presetsValues,
          );
          if (
            presetsUpdateResult.success &&
            presetsUpdateResult.batchUpdate.length > 0
          ) {
            batchUpdate = batchUpdate.concat(presetsUpdateResult.batchUpdate);
          } else if (!presetsUpdateResult.success) {
            failedUpdates.push({
              sheetType: "Presets Presets",
              message: presetsUpdateResult.message,
            });
          }
        } catch (error) {
          console.log(`Error updating Presets data: ${error.toString()}`);
          failedUpdates.push({
            sheetType: "Presets Presets",
            message: "Error updating Presets data: " + error.message,
          });
        }
      }

      // The new IDS Master owns the IDs from here on, so it points at itself and
      // is marked imported without a follow-up updateIdsMaster call.
      var thisSheetInfo = shared.findSheetTypeID(
        newSheetID,
        "IDS",
        "This Sheet ID",
        idsValues,
      );
      if (thisSheetInfo && thisSheetInfo.cell && thisSheetInfo.cell.range) {
        batchUpdate.push({
          range: thisSheetInfo.cell.range,
          values: [[newSheetID]],
        });
      }

      // Execute all updates
      if (batchUpdate.length > 0) {
        var updateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        if (!updateResult) {
          console.log(`Failed to update IDS Master data`);
          return {
            success: false,
            message: "Failed to update IDS Master data",
          };
        }
      }

      var successMessage = "IDS Master data imported successfully";
      if (failedUpdates.length > 0) {
        successMessage += ` (${failedUpdates.length} sections failed)`;
      }

      return {
        success: true,
        message: successMessage,
        failedUpdates: failedUpdates,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: "Error importing IDS Master data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Update Functions
  updateIDSData: function (oldIDSValues, newIDSData) {
    try {
      console.log("Called: master.updateIDSData");
      var batchUpdate = [];

      // Update each sheet reference from old IDS data
      Object.keys(oldIDSValues).forEach(function (sheetType) {
        var sheetID = oldIDSValues[sheetType];

        // Find the sheet type in new IDS values
        var sheetInfo = shared.findSheetTypeID(
          null,
          "IDS",
          sheetType,
          newIDSData,
        );

        if (sheetInfo) {
          // Update ID if provided
          if (sheetID && sheetInfo.cell && sheetInfo.cell.range) {
            batchUpdate.push({
              range: sheetInfo.cell.range,
              values: [[sheetID]],
            });
          }
        }
      });

      return {
        success: true,
        batchUpdate: batchUpdate,
        message: `Updated ${batchUpdate.length} IDS references`,
      };
    } catch (error) {
      console.log("Error in updateIDSData: " + error.toString());
      return {
        success: false,
        message: "Error in updateIDSData: " + error.message,
      };
    }
  },

  updatePresetsData: function (oldPresetsValues, newPresetsData) {
    try {
      console.log("Called: master.updatePresetsData");

      const sliders = ["Range", "Shockwave Size"];
      const oldPresetNames =
        oldPresetsValues.presetNames || Object.keys(oldPresetsValues.data);

      var presetIndex = 0;
      var batchUpdate = [];

      for (var row = 0; row < newPresetsData.length; row++) {
        var rowData = newPresetsData[row];
        if (
          !rowData ||
          rowData.length === 0 ||
          !rowData.some((cell) => cell === "Sliders")
        ) {
          continue;
        }
        var finalRow = row;
        for (var col = 0; col < rowData.length; col++) {
          var cellValue = rowData[col];
          if (!cellValue || cellValue !== "Sliders") {
            continue;
          }
          var presetName = oldPresetNames[presetIndex++];
          if (!presetName || !oldPresetsValues.data.hasOwnProperty(presetName)) {
            continue;
          }

          batchUpdate.push({
            range: `Presets Presets!${shared.columnToLetter(col)}${row - 1}`,
            values: [[presetName]],
          });
          
          var presetData = oldPresetsValues.data[presetName];
          for (
            var nextRow = row + 1;
            nextRow < newPresetsData.length;
            nextRow++
          ) {
            var nextRowData = newPresetsData[nextRow];
            if (!nextRowData || nextRowData.length <= col) {
              continue;
            }
            var key = nextRowData[col];

            const isSlider = sliders.includes(key);
            nextRow = isSlider ? nextRow : nextRow + 1;
            var colIndex = isSlider ? col + 2 : col + 1;
            var levelValue = newPresetsData[nextRow][colIndex];
            if (!key && !levelValue) {
              break;
            }
            if (!presetData.hasOwnProperty(key)) {
              continue;
            }
            var newLevelValue = presetData[key];
            if (newLevelValue !== levelValue) {
              batchUpdate.push({
                range: `Presets Presets!${shared.columnToLetter(
                  colIndex + 1,
                )}${nextRow + 1}`,
                values: [[newLevelValue]],
              });
            }
          }
        }
      }

      return {
        success: true,
        batchUpdate: batchUpdate,
        message: `Updated ${batchUpdate.length} Presets entries`,
      };
    } catch (error) {
      console.log("Error in updatePresetsData: " + error.toString());
      return {
        success: false,
        message: "Error in updatePresetsData: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version4_0: function () {
    try {
      console.log("Called: master.version4_0");
      var oldSpreadsheet = spreadsheets("IDS Master oldSpreadsheet");
      if (!oldSpreadsheet) {
        console.log(`Old spreadsheet not found`);
        return {
          success: false,
          message: "Old spreadsheet not found",
        };
      }
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Get IDS sheet data
      var requiredRanges = ["IDS", "Presets Presets"];
      var idsResult = SheetsAPI.batchGetValues(oldSheetID, requiredRanges);
      if (!idsResult || !idsResult[0] || !idsResult[0].values) {
        console.log(`Could not read IDS data from old spreadsheet`);
        return {
          success: false,
          message: "Could not read IDS data from old spreadsheet",
        };
      }

      var idsValues = idsResult[0].values;
      var presetsValues = idsResult[1].values;

      var idsData = this.getVersion4_0IDSData(idsValues);

      var presetsData = this.getVersion4_0PresetsData(presetsValues);

      var success = idsData.success && presetsData.success;
      var message = `IDS Master export completed successfully. IDS: ${idsData.message}, Presets: ${presetsData.message}`;

      return {
        success: success,
        message: message,
        oldIdsData: idsData.oldIdsData || {},
        oldPresetsData: presetsData.oldPresetsData || {},
      };
    } catch (error) {
      console.log("Error in version4_0: " + error.toString());
      return {
        success: false,
        message: "Error in version4_0: " + error.message,
      };
    }
  },

  version2_0: function () {
    try {
      console.log("Called: master.version2_0");
      var oldSpreadsheet = spreadsheets("IDS Master oldSpreadsheet");
      if (!oldSpreadsheet) {
        console.log(`Old spreadsheet not found`);
        return {
          success: false,
          message: "Old spreadsheet not found",
        };
      }
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Get IDS sheet data
      var idsResult = SheetsAPI.batchGetValues(oldSheetID, ["IDS"]);
      if (!idsResult || !idsResult[0] || !idsResult[0].values) {
        console.log(`Could not read IDS data from old spreadsheet`);
        return {
          success: false,
          message: "Could not read IDS data from old spreadsheet",
        };
      }

      var idsValues = idsResult[0].values;
      var oldIdsData = this.getVersion2_0IDSData(idsValues);
      return oldIdsData;
    } catch (error) {
      console.log("Error in version2_0: " + error.toString());
      return {
        success: false,
        message: "Error in version2_0: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get IDS Data
  getVersion4_0IDSData: function (idsValues) {
    try {
      console.log("Called: master.getVersion4_0IDSData");

      // Extract all sheet references from the IDS sheet
      var sheetReferences = {};
      var sheetTypes = [
        "Laboratory",
        "Workshop",
        "Ultimate Weapon",
        "Themes, Songs & Relics",
        "Bots",
        "Vault",
        "Cards",
        "Modules",
        "Guardians",
        "Player & Stuff",
      ];

      for (var i = 0; i < sheetTypes.length; i++) {
        var sheetType = sheetTypes[i];
        var sheetInfo = shared.findSheetTypeID(
          null,
          "IDS",
          sheetType,
          idsValues,
        );
        if (sheetInfo && sheetInfo.id) {
          var sheetID = shared.extractSheetId(sheetInfo.id);
          sheetReferences[sheetType] = sheetID;
        }
      }

      return {
        success: true,
        oldIdsData: sheetReferences,
        message: "IDS Master data extracted successfully",
      };
    } catch (error) {
      console.log("Error in getVersion4_0IDSData: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion4_0IDSData: " + error.message,
      };
    }
  },

  getVersion2_0IDSData: function (idsValues) {
    try {
      console.log("Called: master.getVersion2_0IDSData");

      // Extract all sheet references from the IDS sheet
      var sheetReferences = {};
      var sheetTypes = [
        "Laboratory",
        "Workshop",
        "Ultimate Weapon",
        "Themes & Songs",
        "Bots",
        "Relics",
        "Vault",
        "Cards",
        "Modules",
        "Guardians",
        "Player & Stuff",
      ];

      for (var i = 0; i < sheetTypes.length; i++) {
        var sheetType = sheetTypes[i];
        var sheetInfo = shared.findSheetTypeID(
          null,
          "IDS",
          sheetType,
          idsValues,
        );
        if (sheetInfo && sheetInfo.id) {
          var sheetID = shared.extractSheetId(sheetInfo.id);
          sheetReferences[sheetType] = sheetID;
        }
      }

      return {
        success: true,
        oldIdsData: sheetReferences,
        message: "IDS Master data extracted successfully",
      };
    } catch (error) {
      console.log("Error in getVersion2_0IDSData: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion2_0IDSData: " + error.message,
      };
    }
  },

  // #endregion
  // #region Get Presets Data
  getVersion4_0PresetsData: function (presetsValues) {
    try {
      console.log("Called: master.getVersion4_0PresetsData");

      // Extract all preset data from the Presets Presets sheet
      const sliders = ["Range", "Shockwave Size"];
      var presetsData = {
        data: {},
      };
      for (var row = 1; row < presetsValues.length; row++) {
        var rowData = presetsValues[row];
        if (
          !rowData ||
          rowData.length === 0 ||
          !rowData.some((cell) => cell === "Sliders")
        ) {
          continue;
        }
        var finalRow = row;
        for (var col = 0; col < rowData.length; col++) {
          var cellValue = rowData[col];
          if (!cellValue || cellValue !== "Sliders") {
            continue;
          }
          var presetName = presetsValues[row - 2][col - 1];
          for (
            var nextRow = row + 1;
            nextRow < presetsValues.length;
            nextRow++
          ) {
            var nextRowData = presetsValues[nextRow];
            if (!nextRowData || nextRowData.length <= col) {
              continue;
            }
            var key = nextRowData[col];

            const isSlider = sliders.includes(key);
            nextRow = isSlider ? nextRow : nextRow + 1;
            var colIndex = isSlider ? col + 2 : col + 1;
            var levelValue = presetsValues[nextRow][colIndex];
            if (!key && !levelValue) {
              finalRow = nextRow;
              break;
            }
            if (!presetsData.data[presetName]) {
              presetsData.data[presetName] = {};
            }
            presetsData.data[presetName][key] = levelValue;
          }
        }
        row = finalRow;
      }

      presetsData.presetNames = shared.resolvePresetOrder(
        Object.keys(presetsData.data),
        shared.templatePresetNames,
      ).order;

      return {
        success: true,
        oldPresetsData: presetsData,
        message: "Presets data extracted successfully",
      };
    } catch (error) {
      console.log("Error in getVersion4_0PresetsData: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion4_0PresetsData: " + error.message,
      };
    }
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v2.0": this.version2_0.bind(this),
      "v4.0": this.version4_0.bind(this),
    };
  },

  // #endregion
  // #region Compatibility Check
  isCompatibleVersion: function (oldVersion) {
    console.log("Called: master.isCompatibleVersion");
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
