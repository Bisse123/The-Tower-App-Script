const master = {
  importData: function (versionDifference) {
    try {
      var idMasterSpreadsheet = spreadsheets("idMasterSpreadsheet");
      if (!idMasterSpreadsheet) {
        console.log(`IDS Master spreadsheet not found`);
        return {
          success: false,
          message: "IDS Master spreadsheet™ not found",
        };
      }
      var idMasterID = idMasterSpreadsheet.spreadsheetId;

      // Get the IDS sheet from the master to read all sheet mappings
      if (!SheetsAPI.getSheetByName(idMasterSpreadsheet, "IDS")) {
        console.log(`IDS sheet not found in the IDS Master Spreadsheet`);
        return {
          success: false,
          message: "IDS sheet™ not found in the IDS Master Spreadsheet™",
        };
      }

      // Define all possible sheet types that can be imported
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
        "Guardians"
      ];

      var allResults = [];
      var successCount = 0;
      var failCount = 0;

      // Process each sheet type
      for (var i = 0; i < sheetTypes.length; i++) {
        var sheetType = sheetTypes[i];
        
        try {
          console.log(`Processing ${sheetType}...`);
          
          // Get old sheet ID from IDS Master
          var sheetInfo = shared.findSheetTypeID(idMasterID, "IDS", sheetType);
          if (!sheetInfo || !sheetInfo.id) {
            console.log(`Sheet info not found for ${sheetType}, skipping...`);
            allResults.push({
              sheetType: sheetType,
              success: false,
              message: `Sheet info not found for ${sheetType}`,
              skipped: true
            });
            continue;
          }

          var oldSheetID = sheetInfo.id;

          // Check if sheet has proper access status
          if (!sheetInfo.accessStatus || 
              !["✅", "Wrong ID or Version"].includes(sheetInfo.accessStatus.value)) {
            console.log(`${sheetType} does not have proper access status, skipping...`);
            allResults.push({
              sheetType: sheetType,
              success: false,
              message: `${sheetType} does not have proper access status`,
              skipped: true
            });
            continue;
          }

          // Create new template copy for this sheet type
          var newTemplateResult = this.createNewTemplate(sheetType);
          if (!newTemplateResult || !newTemplateResult.success) {
            console.log(`Failed to create new template for ${sheetType}: ${newTemplateResult ? newTemplateResult.message : "Unknown error"}`);
            allResults.push({
              sheetType: sheetType,
              success: false,
              message: `Failed to create new template for ${sheetType}: ${newTemplateResult ? newTemplateResult.message : "Unknown error"}`,
              skipped: true
            });
            continue;
          }

          var newSheetID = newTemplateResult.newSheetID;

          // Import data from old sheet to new sheet
          var importResult = this.importSheetData(sheetType, newSheetID, oldSheetID, versionDifference);
          
          if (importResult && importResult.success) {
            successCount++;
            console.log(`Successfully imported ${sheetType}`);
          } else {
            failCount++;
            console.log(`Failed to import ${sheetType}: ${importResult ? importResult.message : "Unknown error"}`);
          }
          
          allResults.push({
            sheetType: sheetType,
            success: importResult ? importResult.success : false,
            message: importResult ? importResult.message : "Unknown error",
            updated: importResult ? importResult.updated : false,
            newSheetID: newSheetID,
            oldSheetID: oldSheetID
          });

        } catch (sheetError) {
          failCount++;
          console.log(`Error processing ${sheetType}: ${sheetError.toString()}`);
          allResults.push({
            sheetType: sheetType,
            success: false,
            message: `Error processing ${sheetType}: ${sheetError.message}`,
            error: true
          });
        }
      }

      // Summary of results
      var summary = `IDS Master import completed: ${successCount} successful, ${failCount} failed out of ${sheetTypes.length} sheet types`;
      console.log(summary);

      return {
        success: successCount > 0,
        message: summary,
        results: allResults,
        successCount: successCount,
        failCount: failCount,
        totalCount: sheetTypes.length
      };

    } catch (error) {
      console.log(`Error in IDS Master importData: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Master importData: ${error.message}`,
      };
    }
  },

  createNewTemplate: function (sheetType) {
    try {
      console.log(`Creating new template for ${sheetType}...`);
      
      // Create a copy of the template file for this sheet type
      var templateResult = shared.copyFileTemplate(sheetType);
      if (!templateResult || !templateResult.success) {
        console.log(`Failed to copy template for ${sheetType}: ${templateResult ? templateResult.message : "Unknown error"}`);
        return {
          success: false,
          message: `Failed to copy template for ${sheetType}: ${templateResult ? templateResult.message : "Unknown error"}`
        };
      }

      var newSheetID = templateResult.newFileId;
      console.log(`Successfully created new template for ${sheetType} with ID: ${newSheetID}`);

      return {
        success: true,
        message: `Template created successfully for ${sheetType}`,
        newSheetID: newSheetID
      };

    } catch (error) {
      console.log(`Error creating template for ${sheetType}: ${error.toString()}`);
      return {
        success: false,
        message: `Error creating template for ${sheetType}: ${error.message}`
      };
    }
  },

  importSheetData: function (sheetType, newSheetID, oldSheetID, versionDifference) {
    try {
      console.log(`Importing data for ${sheetType} from ${oldSheetID} to ${newSheetID}...`);

      // Create unique spreadsheet keys for this import to avoid caching conflicts
      var newSpreadsheetKey = `newSpreadsheet_${sheetType}_${Date.now()}`;
      var oldSpreadsheetKey = `oldSpreadsheet_${sheetType}_${Date.now()}`;

      // Get spreadsheet instances with unique keys
      var newSpreadsheet = spreadsheets(newSpreadsheetKey, newSheetID);
      var oldSpreadsheet = spreadsheets(oldSpreadsheetKey, oldSheetID);
      var idMasterSpreadsheet = spreadsheets("idMasterSpreadsheet");

      if (!newSpreadsheet || !oldSpreadsheet) {
        console.log(`Could not access spreadsheets for ${sheetType}`);
        return {
          success: false,
          message: `Could not access spreadsheets for ${sheetType}`
        };
      }

      // Temporarily override the global spreadsheet functions to use our specific instances
      var originalSpreadsheets = global.spreadsheets || spreadsheets;
      global.spreadsheets = function(type) {
        switch(type) {
          case "newSpreadsheet":
            return newSpreadsheet;
          case "oldSpreadsheet":
            return oldSpreadsheet;
          case "idMasterSpreadsheet":
            return idMasterSpreadsheet;
          default:
            return originalSpreadsheets(type);
        }
      };

      try {
        // Get the sheet type function and import data
        var sheetTypeFunction = sheetVars(sheetType);
        if (!sheetTypeFunction) {
          console.log(`Sheet type function not found for ${sheetType}`);
          return {
            success: false,
            message: `Sheet type function not found for ${sheetType}`
          };
        }

        var result = sheetTypeFunction.importData(versionDifference);
        
        if (result && result.success) {
          // Update import status in the new sheet
          try {
            var newSheetInfo = shared.findSheetTypeID(newSheetID, "IDS");
            if (newSheetInfo && newSheetInfo.importStatus) {
              SheetsAPI.batchUpdateValues(newSheetID, [
                {
                  range: newSheetInfo.importStatus.range,
                  values: [["✅"]],
                },
              ]);
            }
          } catch (statusError) {
            console.log(`Error updating import status for ${sheetType}: ${statusError.toString()}`);
          }
        }

        return result;

      } finally {
        // Restore original spreadsheets function
        global.spreadsheets = originalSpreadsheets;
      }

    } catch (error) {
      console.log(`Error importing data for ${sheetType}: ${error.toString()}`);
      return {
        success: false,
        message: `Error importing data for ${sheetType}: ${error.message}`
      };
    }
  },
};
