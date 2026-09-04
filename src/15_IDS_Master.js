const master = {

  /**
   * Reads IDS_Master data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
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

      var oldDataResult = getVersionFunction(oldSheetID);
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
      var errorReport = errors.report("master.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported IDS_Master data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: master.importData");

      var batchUpdate = [];
      var failedUpdates = [];

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
          var errorReport = errors.report("master.importData", error, {
            note: `Error updating IDS data`,
            data: data,
            newSheetID: newSheetID,
          });
          failedUpdates.push({
            sheetType: "IDS",
            message: errorReport.message,
            reference: errorReport.reference,
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
          var errorReport = errors.report("master.importData", error, {
            note: `Error updating Presets data`,
            data: data,
            newSheetID: newSheetID,
          });
          failedUpdates.push({
            sheetType: "Presets Presets",
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

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
      var errorReport = errors.report("master.importData", error, {
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes IDSData into the new sheet.
   * @param {Array<Array<*>>} oldIDSValues
   * @param {Object} newIDSData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateIDSData: function (oldIDSValues, newIDSData) {
    try {
      console.log("Called: master.updateIDSData");
      var batchUpdate = [];

      Object.keys(oldIDSValues).forEach(function (sheetType) {
        var sheetID = oldIDSValues[sheetType];

        var sheetInfo = shared.findSheetTypeID(
          null,
          "IDS",
          sheetType,
          newIDSData,
        );

        if (sheetInfo) {

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
      var errorReport = errors.report("master.updateIDSData", error, {
        oldIDSValues: oldIDSValues,
        newIDSData: newIDSData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes PresetsData into the new sheet.
   * @param {Array<Array<*>>} oldPresetsValues
   * @param {Object} newPresetsData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
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

          var oldPresetData = oldPresetsValues.data[presetName];
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
              finalRow = nextRow;
              break;
            }
            if (!oldPresetData.hasOwnProperty(key)) {
              continue;
            }
            var oldLevelValue = oldPresetData[key];
            if (oldLevelValue !== levelValue && oldLevelValue !== presetName) {
              batchUpdate.push({
                range: `Presets Presets!${shared.columnToLetter(
                  colIndex + 1,
                )}${nextRow + 1}`,
                values: [[oldLevelValue]],
              });
            }
          }
        }
        row = finalRow;
      }

      return {
        success: true,
        batchUpdate: batchUpdate,
        message: `Updated ${batchUpdate.length} Presets entries`,
      };
    } catch (error) {
      var errorReport = errors.report("master.updatePresetsData", error, {
        oldPresetsValues: oldPresetsValues,
        newPresetsData: newPresetsData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads IDS_Master data from a v4.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version4_0: function (oldSheetID) {
    try {
      console.log("Called: master.version4_0");

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
      var errorReport = errors.report("master.version4_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads IDS_Master data from a v2.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version2_0: function (oldSheetID) {
    try {
      console.log("Called: master.version2_0");

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
      var errorReport = errors.report("master.version2_0", error, {
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts IDSData from a v4.0 sheet's values.
   * @param {Array<Array<*>>} idsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0IDSData: function (idsValues) {
    try {
      console.log("Called: master.getVersion4_0IDSData");

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
      var errorReport = errors.report("master.getVersion4_0IDSData", error, {
        idsValues: idsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts IDSData from a v2.0 sheet's values.
   * @param {Array<Array<*>>} idsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion2_0IDSData: function (idsValues) {
    try {
      console.log("Called: master.getVersion2_0IDSData");

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
      var errorReport = errors.report("master.getVersion2_0IDSData", error, {
        idsValues: idsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts PresetsData from a v4.0 sheet's values.
   * @param {Array<Array<*>>} presetsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0PresetsData: function (presetsValues) {
    try {
      console.log("Called: master.getVersion4_0PresetsData");

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
            var presetType = nextRowData[col];

            const isSlider = sliders.includes(presetType);
            nextRow = isSlider ? nextRow : nextRow + 1;
            var colIndex = isSlider ? col + 2 : col + 1;
            var levelValue = presetsValues[nextRow][colIndex];
            if (!presetType && !levelValue) {
              finalRow = nextRow;
              break;
            }
            if (!presetsData.data[presetName]) {
              presetsData.data[presetName] = {};
            }
            presetsData.data[presetName][presetType] = levelValue;
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
      var errorReport = errors.report("master.getVersion4_0PresetsData", error, {
        presetsValues: presetsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Parses IDS_Master data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseMasterData: function (data) {
    try {
      const globalPresets = data.globalPresets || [];
      const workshopPresetNames = data.workshopPresetNames || [];
      const cardPresetNames = data.cardPresetNames || [];
      const botPresetNames = data.botPresetNames || [];
      const modulePresets = data.modulePresets || [];
      const guardianPresets = data.guardianPresets || [];
      const modulePresetNames = modulePresets.map((preset) => preset.presetName);
      const guardianPresetNames = guardianPresets.map((preset) => preset.presetName);

      const presetInfo = [
        {"type": "Workshop", "indexName": "workshopIndex", "data": workshopPresetNames},
        {"type": "Cards", "indexName": "cardsIndex", "data": cardPresetNames},
        {"type": "Bots", "indexName": "botsIndex", "data": botPresetNames},
        {"type": "Modules", "indexName": "modulesIndex", "data": modulePresetNames},
        {"type": "Guardians", "indexName": "guardiansIndex", "data": guardianPresetNames},
    ]

      var oldPresetNames = [];
      var oldPresetsData = {
        data: {},
      };
      globalPresets.forEach((preset, index) => {
        if (index == globalPresets.length - 1) return;
        const globalPresetName = preset.presetName;
        if (!globalPresetName) {
          return;
        }
        oldPresetNames.push(globalPresetName);
        if (!oldPresetsData.data.hasOwnProperty(globalPresetName)) {
          oldPresetsData.data[globalPresetName] = {};
        }
        presetInfo.forEach((info) => {
          const presetType = info.type;
          const indexName = info.indexName;
          const index = preset[indexName];
          const data = info.data;
          const presetName = data[index];
          if (presetName) {
            oldPresetsData.data[globalPresetName][presetType] = presetName;
          }
        });
      });

      oldPresetsData.presetNames = shared.resolvePresetOrder(
          oldPresetNames,
          shared.templatePresetNames,
        ).order;

      const presetTypesOrder = presetInfo.map((info) => info.type);
      return {
        success: true,
        success: true,
        oldPresetsData: oldPresetsData,
        presetTypesOrder: presetTypesOrder,
      };
    } catch (error) {
      var errorReport = errors.report("master.parseMasterData", error, {
        data: data,
        oldPresetsData: oldPresetsData,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v2.0": this.version2_0.bind(this),
      "v4.0": this.version4_0.bind(this),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
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

};
