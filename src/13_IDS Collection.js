const collection = {
  importData: function (versionDifference) {
    try {
      var newSpreadsheet = spreadsheets("IDS Collection - all IDS-Sheets on one file newSpreadsheet");
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet™ not found",
        };
      }
      var newSheetID = newSpreadsheet.spreadsheetId;

      var oldSpreadsheet = spreadsheets("IDS Collection - all IDS-Sheets on one file oldSpreadsheet");
      if (!oldSpreadsheet) {
        console.log(`Old spreadsheet not found`);
        return {
          success: false,
          message: "Old spreadsheet™ not found",
        };
      }
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      
    } catch (error) {
      console.log(`Error in importCollectionsData: ${error.message}`);
      return {
        success: false,
        message: `Error in importCollectionsData: ${error.message}`,
      };
    }
  },
};
