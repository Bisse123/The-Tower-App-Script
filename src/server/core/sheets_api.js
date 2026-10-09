const SheetsAPI = {
  /**
   * Sheet properties for a spreadsheet.
   * @param {string} spreadsheetId
   * @returns {Object|null} Null on failure; the failure is reported.
   */
  fetchSpreadsheet: function (spreadsheetId) {
    try {
      const response = Sheets.Spreadsheets.get(spreadsheetId, {
        fields: "spreadsheetId,sheets(properties(sheetId,title,hidden))",
      });
      return response;
    } catch (error) {
      errors.report("SheetsAPI.fetchSpreadsheet", error, {
        note: `Error getting spreadsheet`,
        spreadsheetId: spreadsheetId,
      });
      return null;
    }
  },

  /**
   * Finds a tab by exact title.
   * @param {Object} spreadsheet
   * @param {string} sheetName
   * @returns {Object|null} The tab's properties.
   */
  getSheetByName: function (spreadsheet, sheetName) {
    try {
      const sheet = spreadsheet.sheets.find(
        (s) => s.properties.title === sheetName,
      );
      return sheet ? sheet.properties : null;
    } catch (error) {
      errors.report("SheetsAPI.getSheetByName", error, {
        note: `Error getting sheet by name`,
        spreadsheet: spreadsheet,
        sheetName: sheetName,
      });
      return null;
    }
  },

  /**
   * Finds a tab whose title contains substring.
   * @param {Object} spreadsheet
   * @param {string} substring
   * @returns {Object|null} The tab's properties.
   */
  getSheetBySubstring: function (spreadsheet, substring) {
    try {
      const sheet = spreadsheet.sheets.find((s) =>
        s.properties.title.toLowerCase().includes(substring.toLowerCase()),
      );
      return sheet ? sheet.properties : null;
    } catch (error) {
      errors.report("SheetsAPI.getSheetBySubstring", error, {
        note: `Error getting sheet by substring`,
        spreadsheet: spreadsheet,
        substring: substring,
      });
      return null;
    }
  },

  /**
   * Builds the requests that restore tab visibility.
   * @param {Object} newSpreadsheet
   * @param {Object} sheetVisibility Title to hidden flag.
   * @returns {{success: boolean, message: string, requests?: Array<Object>}}
   */
  applySheetVisibility: function (newSpreadsheet, sheetVisibility) {
    try {
      if (!newSpreadsheet) {
        console.error("Missing spreadsheet parameter");
        return {
          success: false,
          message: "Missing spreadsheet parameter",
        };
      }

      if (!newSpreadsheet.sheets) {
        console.error("Invalid spreadsheet structure - missing sheets");
        return {
          success: false,
          message: "Invalid spreadsheet structure - missing sheets",
        };
      }

      if (!sheetVisibility || typeof sheetVisibility !== "object") {
        return {
          success: true,
          message: "No sheet visibility data provided",
          processedSheets: [],
        };
      }

      var requests = [];
      var processedSheets = [];

      for (var i = 0; i < newSpreadsheet.sheets.length; i++) {
        var newSheet = newSpreadsheet.sheets[i];
        var newSheetName = newSheet.properties.title;
        var newSheetId = newSheet.properties.sheetId;

        if (sheetVisibility.hasOwnProperty(newSheetName)) {
          var targetHidden = sheetVisibility[newSheetName];
          var currentHidden = newSheet.properties.hidden || false;

          if (targetHidden !== currentHidden) {
            requests.push({
              updateSheetProperties: {
                properties: {
                  sheetId: newSheetId,
                  hidden: targetHidden,
                },
                fields: "hidden",
              },
            });
            processedSheets.push({
              name: newSheetName,
              action: targetHidden ? "hidden" : "shown",
            });
          }
        }
      }

      if (requests.length > 0) {
        Sheets.Spreadsheets.batchUpdate(
          {
            requests: requests,
          },
          newSpreadsheet.spreadsheetId,
        );

        return {
          success: true,
          message: `Successfully updated visibility for ${requests.length} sheets`,
          processedSheets: processedSheets,
        };
      } else {
        return {
          success: true,
          message: "No sheet visibility changes needed",
          processedSheets: [],
        };
      }
    } catch (error) {
      var errorReport = errors.report("SheetsAPI.applySheetVisibility", error, {
        note: `Error applying sheet visibility`,
        newSpreadsheet: newSpreadsheet,
        sheetVisibility: sheetVisibility,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads several ranges' values in one call.
   * @param {string} spreadsheetId
   * @param {string[]} ranges
   * @param {boolean} [useCache]
   * @param {boolean} [forceRefresh]
   * @returns {Array<Object>|null} Null on failure; the failure is reported.
   */
  batchGetValues: function (
    spreadsheetId,
    ranges,
    useCache = true,
    forceRefresh = false,
  ) {
    try {
      if (useCache) {
        return CacheManager.getSheetValues(spreadsheetId, ranges, forceRefresh);
      }
      const response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
        ranges: ranges,
      });
      return response.valueRanges;
    } catch (error) {
      errors.report("SheetsAPI.batchGetValues", error, {
        spreadsheetId: spreadsheetId,
        ranges: ranges,
      });
      return null;
    }
  },

  /**
   * Reads several ranges' formulas in one call.
   * @param {string} spreadsheetId
   * @param {string[]} ranges
   * @param {boolean} [useCache]
   * @param {boolean} [forceRefresh]
   * @returns {Array<Object>|null} Null on failure; the failure is reported.
   */
  batchGetFormulas: function (
    spreadsheetId,
    ranges,
    useCache = true,
    forceRefresh = false,
  ) {
    try {
      if (useCache) {
        return CacheManager.getSheetFormulas(
          spreadsheetId,
          ranges,
          forceRefresh,
        );
      }
      const response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
        ranges: ranges,
        valueRenderOption: "FORMULA",
      });
      return response.valueRanges;
    } catch (error) {
      errors.report("SheetsAPI.batchGetFormulas", error, {
        spreadsheetId: spreadsheetId,
        ranges: ranges,
      });
      return null;
    }
  },

  /**
   * Writes a batch of range updates.
   * @param {string} spreadsheetId
   * @param {Array<{range: string, values: Array}>} updates
   * @returns {Object|null} Null on failure; the failure is reported.
   */
  batchUpdateValues: function (spreadsheetId, updates) {
    try {
      const requestBody = {
        data: updates,
        valueInputOption: "USER_ENTERED",
      };
      return Sheets.Spreadsheets.Values.batchUpdate(requestBody, spreadsheetId);
    } catch (error) {
      errors.report("SheetsAPI.batchUpdateValues", error, {
        spreadsheetId: spreadsheetId,
        updates: updates,
      });
      return null;
    }
  },
};
