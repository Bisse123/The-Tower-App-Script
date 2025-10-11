const master = {
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
        data: oldDataResult.data,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting IDS Master data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      console.log("Called: master.importData");
      console.log(JSON.stringify(data));
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
      var idsData = SheetsAPI.batchGetValues(newSheetID, ["IDS"]);
      if (!idsData || !idsData[0] || !idsData[0].values) {
        console.log(`Could not read IDS data from new spreadsheet`);
        return {
          success: false,
          message: "Could not read IDS data from new spreadsheet",
        };
      }

      var idsValues = idsData[0].values;

      // Update IDS sheet with exported data if available
      if (data && data.idsData) {
        try {
          var idsUpdateResult = this.updateIDSData(data.idsData, idsValues);
          if (idsUpdateResult.success && idsUpdateResult.batchUpdate.length > 0) {
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

  updateIDSData: function (oldIDSValues, newIDSData) {
    try {
      console.log("Called: master.updateIDSData");
      var batchUpdate = [];

      // Update each sheet reference from old IDS data
      Object.keys(oldIDSValues).forEach(function(sheetType) {
        var sheetID = oldIDSValues[sheetType];
        
        // Find the sheet type in new IDS values
        var sheetInfo = shared.findSheetTypeID(
          null,
          "IDS",
          sheetType,
          newIDSData
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

  version20: function () {
    try {
      console.log("Called: master.version20");
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
      var idsDataResult = this.getVersion20IDSData(idsValues);

      if (!idsDataResult.success) {
        console.log(`${idsDataResult.message}`);
        return idsDataResult;
      }

      return {
        success: true,
        data: {
          idsData: idsDataResult.data,
        },
        message: "IDS Master v2.0 data exported successfully",
      };
    } catch (error) {
      console.log("Error in version20: " + error.toString());
      return {
        success: false,
        message: "Error in version20: " + error.message,
      };
    }
  },

  getVersion20IDSData: function (idsValues) {
    try {
      console.log("Called: master.getVersion20IDSData");
      
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
          idsValues
        );
        if (sheetInfo && sheetInfo.id) {
          var sheetID = shared.extractSheetId(sheetInfo.id);
          sheetReferences[sheetType] = sheetID;
        }
      }

      return {
        success: true,
        data: sheetReferences,
        message: "IDS Master data extracted successfully",
      };
    } catch (error) {
      console.log("Error in getVersion20IDSData: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion20IDSData: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v2.0": this.version20.bind(this),
    };
  },

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
