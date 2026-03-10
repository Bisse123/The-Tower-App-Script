const SheetsAPI = {
  // Get spreadsheet metadata and sheets
  getSpreadsheet: function (spreadsheetId) {
    try {
      const response = Sheets.Spreadsheets.get(spreadsheetId, {
        fields: "spreadsheetId,sheets(properties(sheetId,title,hidden))",
      });
      return response;
    } catch (error) {
      console.error(`Error getting spreadsheet: ${error}`);
      console.error(`Error details: ${JSON.stringify(error)}`);
      return null;
    }
  },

  // Get sheet by name from spreadsheet
  getSheetByName: function (spreadsheet, sheetName) {
    try {
      const sheet = spreadsheet.sheets.find(
        (s) => s.properties.title === sheetName
      );
      return sheet ? sheet.properties : null;
    } catch (error) {
      console.error(`Error getting sheet by name: ${error}`);
      return null;
    }
  },

  // Find sheet by substring in name from spreadsheet
  getSheetBySubstring: function (spreadsheet, substring) {
    try {
      const sheet = spreadsheet.sheets.find((s) =>
        s.properties.title.toLowerCase().includes(substring.toLowerCase())
      );
      return sheet ? sheet.properties : null;
    } catch (error) {
      console.error(`Error getting sheet by substring: ${error}`);
      return null;
    }
  },

  // Apply sheet visibility using provided visibility data
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

      // Process each sheet in the new spreadsheet
      for (var i = 0; i < newSpreadsheet.sheets.length; i++) {
        var newSheet = newSpreadsheet.sheets[i];
        var newSheetName = newSheet.properties.title;
        var newSheetId = newSheet.properties.sheetId;

        // Check if we have visibility data for this sheet
        if (sheetVisibility.hasOwnProperty(newSheetName)) {
          var targetHidden = sheetVisibility[newSheetName];
          var currentHidden = newSheet.properties.hidden || false;

          // Only update if visibility states are different
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

      // Execute batch update if there are requests
      if (requests.length > 0) {
        Sheets.Spreadsheets.batchUpdate(
          {
            requests: requests,
          },
          newSpreadsheet.spreadsheetId
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
      console.error(`Error applying sheet visibility: ${error}`);
      return {
        success: false,
        message: `Error updating sheet visibility: ${error.toString()}`,
      };
    }
  },

  // Batch get values from multiple ranges
  batchGetValues: function (spreadsheetId, ranges) {
    try {
      const response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
        ranges: ranges,
      });
      return response.valueRanges;
    } catch (error) {
      console.error(`Error in batchGetValues: ${error}`);
      return null;
    }
  },

  // Batch get formulas from multiple ranges
  batchGetFormulas: function (spreadsheetId, ranges) {
    try {
      const response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
        ranges: ranges,
        valueRenderOption: "FORMULA",
      });
      return response.valueRanges;
    } catch (error) {
      console.error(`Error in batchGetFormulas: ${error}`);
      return null;
    }
  },

  batchUpdateValues: function (spreadsheetId, updates) {
    try {
      const requestBody = {
        data: updates,
        valueInputOption: "USER_ENTERED",
      };
      return Sheets.Spreadsheets.Values.batchUpdate(requestBody, spreadsheetId);
    } catch (error) {
      console.error(`Error in batchUpdateValues: ${error}`);
      return null;
    }
  },
};

const shared = {
  findSheetVersion: function (sheetID, sheetName, sheetType, preLoadedValues) {
    try {
      // If sheetType is 'Effective Paths', delegate to getEPathsVersion
      if (sheetType === "Effective Paths") {
        return shared.getEPathsVersion(sheetID, sheetName, preLoadedValues);
      }
      var values;
      // Use pre-loaded values if provided, otherwise fetch them
      if (preLoadedValues) {
        values = preLoadedValues;
      } else {
        var batchResult = SheetsAPI.batchGetValues(sheetID, [sheetName]);
        if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
          console.log(`No data found in sheet: ${sheetName} in spreadsheet: ${sheetID}`);
          return null;
        }
        values = batchResult[0].values;
      }
      var currentVersion = null;
      var latestVersion = null;
      for (var row = 0; row < values.length; row++) {
        var currentVersionCol = values[row].findIndex(
          (cell) =>
            typeof cell === "string" &&
            ["version change", "this version", "version check"].some((w) =>
              cell.toLowerCase().includes(w)
            )
        );
        var latestVersionCol = values[row].findIndex(
          (cell) =>
            typeof cell === "string" &&
            ["latest remote version", "latest version"].some((w) =>
              cell.toLowerCase().includes(w)
            )
        );
        if (currentVersionCol !== -1 && !currentVersion) {
          currentVersion = values[row + 1][currentVersionCol];
        }
        if (latestVersionCol !== -1 && !latestVersion) {
          latestVersion = values[row + 1][latestVersionCol];
        }
        if (currentVersion && latestVersion) {
          break;
        }
      }
      return {
        currentVersion: currentVersion,
        latestVersion: latestVersion,
      };
    } catch (error) {
      console.error(`Error finding sheet version: ${error}`);
      return null;
    }
  },

  // Get Effective Paths version information (both current and latest)
  getEPathsVersion: function (sheetID, sheetName, preLoadedValues) {
    try {
      var values;

      // Use pre-loaded values if provided, otherwise fetch them
      if (preLoadedValues) {
        values = preLoadedValues;
      } else {
        var batchResult = SheetsAPI.batchGetValues(sheetID, [sheetName]);
        if (
          !batchResult ||
          batchResult.length === 0 ||
          !batchResult[0].values
        ) {
          console.log(
            `No data found in sheet: ${sheetName} in spreadsheet: ${sheetID}`
          );
          return null;
        }
        values = batchResult[0].values;
      }

      var currentVersion = null;
      var latestVersion = null;

      // Search for version information in the sheet
      for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
          var cellValue = values[i] && values[i][j] ? values[i][j] : "";

          // Look for "Current Version:" and concatenate next 2 cells
          if (
            cellValue &&
            typeof cellValue === "string" &&
            cellValue.includes("Current Version:") &&
            !currentVersion
          ) {
            var currentPart1 =
              values[i] && values[i][j + 1] ? values[i][j + 1] : "";
            var currentPart2 =
              values[i] && values[i][j + 2] ? values[i][j + 2] : "";
            currentVersion = currentPart1 + currentPart2;
          }

          // Look for "Latest Version:" and concatenate next 2 cells
          if (
            cellValue &&
            typeof cellValue === "string" &&
            cellValue.includes("Latest Version:") &&
            !latestVersion
          ) {
            var latestPart1 =
              values[i] && values[i][j + 1] ? values[i][j + 1] : "";
            var latestPart2 =
              values[i] && values[i][j + 2] ? values[i][j + 2] : "";
            latestVersion = latestPart1 + latestPart2;
          }

          // If we found both versions, break out of loops
          if (currentVersion && latestVersion) {
            break;
          }
        }
        if (currentVersion && latestVersion) {
          break;
        }
      }

      return {
        currentVersion: currentVersion,
        latestVersion: latestVersion,
      };
    } catch (error) {
      console.error(`Error finding Effective Paths version: ${error}`);
      return null;
    }
  },

  compareVersions: function (oldVersion, newVersion) {
    function parseVersion(v) {
      v = v.replace(/^[^\d]*/, "");
      return v.split(".").map(Number);
    }

    var oldParts = parseVersion(oldVersion || "");
    var newParts = parseVersion(newVersion || "");
    var len = Math.max(oldParts.length, newParts.length);

    for (var i = 0; i < len; i++) {
      var oldNum = oldParts[i] || 0;
      var newNum = newParts[i] || 0;
      if (oldNum > newNum) return "newer";
      if (oldNum < newNum) return "older";
    }
    return "same";
  },

  findSheetTypeID: function (
    spreadsheetId,
    sheetName,
    sheetType,
    preLoadedValues
  ) {
    var sheetType = sheetType || "IDS Master's";
    var values;

    // Use pre-loaded values if provided, otherwise fetch them
    if (preLoadedValues) {
      values = preLoadedValues;
    } else {
      var batchResult = SheetsAPI.batchGetValues(spreadsheetId, [sheetName]);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log(
          `No data found in sheet: ${sheetName} in spreadsheet: ${spreadsheetId}`
        );
        return null;
      }
      values = batchResult[0].values;
    }

    var regex = new RegExp(sheetType, "i");
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (regex.test(values[i][j]) && values[i][j].indexOf("script") === -1) {
          var cellA1 = shared.columnToLetter(j + 2) + (i + 1);
          var accessA1 = shared.columnToLetter(j + 4) + (i + 1);
          var importedA1 = shared.columnToLetter(j + 4) + (i + 2);

          // Get access and import status values from the same pre-fetched data
          var accessValue = "";
          var importValue = "";

          if (values[i] && values[i][j + 3]) {
            accessValue = values[i][j + 3];
          }
          if (values[i + 1] && values[i + 1][j + 3]) {
            importValue = values[i + 1][j + 3];
          }

          return {
            id: values[i][j + 2],
            cell: {
              row: i + 1,
              col: j + 2,
              range: sheetName + "!" + cellA1,
            },
            accessStatus: {
              row: i + 1,
              col: j + 4,
              range: sheetName + "!" + accessA1,
              value: accessValue,
            },
            importStatus: {
              row: i + 2,
              col: j + 4,
              range: sheetName + "!" + importedA1,
              value: importValue,
            },
          };
        }
      }
    }
    return null;
  },

  findSheetTypeURL: function (
    spreadsheetId,
    sheetName,
    sheetType,
    preLoadedValues
  ) {
    var sheetType = sheetType || "IDS Master's";
    var values;

    // Use pre-loaded values if provided, otherwise fetch them
    if (preLoadedValues) {
      values = preLoadedValues;
    } else {
      var batchResult = SheetsAPI.batchGetValues(spreadsheetId, [sheetName]);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log(
          `No data found in sheet: ${sheetName} in spreadsheet: ${spreadsheetId}`
        );
        return null;
      }
      values = batchResult[0].values;
    }

    var regex = new RegExp(sheetType, "i");
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (
          regex.test(values[i][j]) &&
          values[i][j].indexOf("script") === -1 &&
          values[i][j].indexOf("More IDs are available") === -1
        ) {
          var versionA1 = shared.columnToLetter(j + 6) + (i + 1);
          var templateA1 = shared.columnToLetter(j + 1) + (i + 2);
          var oldVersionA1 = shared.columnToLetter(j + 7) + (i + 1);

          var versionValue = "";
          var oldVersionValue = "";
          if (values[i] && values[i][j + 5]) {
            versionValue = values[i][j + 5];
          }
          if (values[i] && values[i][j + 6]) {
            oldVersionValue = values[i][j + 6];
          }

          return {
            id: values[i][j + 2],
            template: {
              row: i + 2,
              col: j + 1,
              range: sheetName + "!" + templateA1,
            },
            version: {
              row: i + 1,
              col: j + 6,
              range: sheetName + "!" + versionA1,
              value: versionValue,
            },
            oldVersion: {
              row: i + 1,
              col: j + 7,
              range: sheetName + "!" + oldVersionA1,
              value: oldVersionValue,
            },
          };
        }
      }
    }
    return null;
  },

  extractSheetId: function (input) {
    input = input.trim();
    var idPattern = /^[a-zA-Z0-9-_]{20,}$/;
    var urlPattern = /\/spreadsheets\/d\/([a-zA-Z0-9-_]{20,})/;

    if (idPattern.test(input)) {
      return input;
    }
    var match = input.match(urlPattern);
    if (match && match[1]) {
      return match[1];
    }
    return null;
  },

  columnToLetter: function (column) {
    var temp = "";
    var letter = "";
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  },

  // Extract URL from HYPERLINK formula
  extractUrlFromHyperlink: function (formula) {
    if (!formula || typeof formula !== "string") {
      return null;
    }

    // Check if formula starts with "="
    if (!formula.startsWith("=")) {
      return null;
    }

    // Match HYPERLINK function anywhere in the formula: HYPERLINK("url", "text")
    var hyperlinkMatch = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
    if (hyperlinkMatch && hyperlinkMatch[1]) {
      return hyperlinkMatch[1];
    }

    return null;
  },

  getDVTValue: function (oldValue, dvtNamedRangesData) {
    if (!oldValue || !dvtNamedRangesData) {
      return oldValue; // Return original if no data
    }

    var oldLevel = oldValue.substring(0, 2);
    for (var i = 0; i < dvtNamedRangesData.length; i++) {
      var row = dvtNamedRangesData[i];
      var val = row[0];
      if (val && val.substring(0, 2) === oldLevel) {
        return val; // Return matched value
      }
    }
    return oldValue; // Return original if no match found
  },

  findSheetTemplateID: function (sheetID, sheetName, sheetType) {
    try {
      console.log(
        `Finding template ID for sheet: ${sheetID}, sheet name: ${sheetName}, type: ${sheetType}`
      );

      var spreadsheet = spreadsheets(`${sheetType} spreadsheet`, sheetID);
      if (!spreadsheet) {
        console.log(`Could not access spreadsheet with ID: ${sheetID}`);
        return null;
      }

      var sheet = SheetsAPI.getSheetByName(spreadsheet, sheetName);
      if (!sheet) {
        console.log(`Could not find sheet: ${sheetName}`);
        return null;
      }

      // Get all data from the sheet to search for HYPERLINK formulas
      // Use batch API calls for better performance
      var formulas = SheetsAPI.batchGetFormulas(sheetID, [sheetName]);
      var values = SheetsAPI.batchGetValues(sheetID, [sheetName]);

      if (!formulas || !formulas[0] || !formulas[0].values) {
        console.log(`Could not fetch formulas from sheet: ${sheetName}`);
        return null;
      }

      if (!values || !values[0] || !values[0].values) {
        console.log(`Could not fetch values from sheet: ${sheetName}`);
        return null;
      }

      var formulaData = formulas[0].values;
      var valueData = values[0].values;

      console.log(
        `Searching ${formulaData.length} rows for template HYPERLINK formulas`
      );

      var templateID = null;
      var version = null;

      // Search for template HYPERLINK formula
      for (var i = 0; i < formulaData.length; i++) {
        for (var j = 0; j < formulaData[i].length; j++) {
          var formula = formulaData[i][j];

          // Check if this cell contains a HYPERLINK formula with "copy" text
          if (
            formula &&
            typeof formula === "string" &&
            formula.toUpperCase().includes("HYPERLINK") &&
            formula.toLowerCase().includes("copy")
          ) {
            console.log(
              `Found potential template link in row ${i + 1}, col ${
                j + 1
              }: ${formula}`
            );

            // Extract URL from the HYPERLINK formula
            var templateUrl = shared.extractUrlFromHyperlink(formula);
            if (templateUrl) {
              // Extract sheet ID from the URL
              templateID = shared.extractSheetId(templateUrl);
              if (templateID) {
                console.log(`Found template ID: ${templateID}`);
              }
            }
          }

        }
        if (templateID) {
          break; // Break outer loop for all sheet types
        }
      }

      // If template ID was found, get version info and return result
      if (templateID) {
        var currentSheetVersionInfo = shared.findSheetVersion(
          sheetID,
          sheetName,
          sheetType,
          valueData
        );
        if (currentSheetVersionInfo && currentSheetVersionInfo.latestVersion) {
          console.log(
            `Template version (from latest): ${currentSheetVersionInfo.latestVersion}`
          );
          return {
            templateID: templateID,
            templateVersion: currentSheetVersionInfo.latestVersion,
          };
        }
      }

      console.log(
        `No template HYPERLINK with "copy" found in sheet: ${sheetName}`
      );
      return null;
    } catch (error) {
      console.error(`Error finding template ID: ${error.toString()}`);
      return null;
    }
  },

  // Utility function to extract column offset from a range string
  getColumnOffsetFromRange: function (range) {
    // Extract the range part after the sheet name (e.g., "AC1:AN35" from "eHP!AC1:AN35")
    var rangePart = range.split("!")[1];
    if (!rangePart) return 0;

    // Extract the starting column (e.g., "AC" from "AC1:AN35")
    var startCell = rangePart.split(":")[0];
    if (!startCell) return 0;

    // Extract column letters (remove numbers)
    var columnLetters = startCell.replace(/[0-9]/g, "");

    // Convert column letters to 0-based index
    var columnIndex = 0;
    for (var i = 0; i < columnLetters.length; i++) {
      columnIndex =
        columnIndex * 26 +
        (columnLetters.charCodeAt(i) - "A".charCodeAt(0) + 1);
    }

    return columnIndex - 1; // Convert to 0-based index
  },

  // Helper function to add sheet ID and IDS Master ID updates to batch
  addIDUpdatesToBatch: function (
    batchUpdate,
    sheetType,
    newSheetID,
    idsData,
    idMasterID
  ) {
    try {
      if (newSheetID && idMasterID) {
        // Find the "This Sheet ID" and "IDS Master's" entries
        var thisSheetInfo = shared.findSheetTypeID(
          newSheetID,
          "IDS",
          "This Sheet ID",
          idsData
        );
        var idMasterInfo = shared.findSheetTypeID(
          newSheetID,
          "IDS",
          "IDS Master's",
          idsData
        );

        // Add ID updates to batch if entries are found
        if (thisSheetInfo && thisSheetInfo.cell && thisSheetInfo.cell.range) {
          batchUpdate.push({
            range: thisSheetInfo.cell.range,
            values: [[newSheetID]],
          });
        }
        if (idMasterInfo && idMasterInfo.cell && idMasterInfo.cell.range) {
          batchUpdate.push({
            range: idMasterInfo.cell.range,
            values: [[idMasterID]],
          });
        }
      }
      return batchUpdate;
    } catch (error) {
      console.log(`Error adding ID updates to batch: ${error.toString()}`);
      return batchUpdate; // Return original batch if error occurs
    }
  },
};

function moveSheet(sheetType, newSheetID, oldSheetID) {
  try {
    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID
    );
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`,
      };
    }

    var newFile = Drive.Files.get(newSheetID, { fields: "id, name, parents" });
    var oldFile = Drive.Files.get(oldSheetID, { fields: "id, name, parents" });
    if (!newFile || !oldFile) {
      console.log(`Could not retrieve file information for new or old sheet.`);
      return {
        success: false,
        message: `Could not retrieve file information for new or old sheet™.`,
      };
    }

    var newVersionInfo;
    newVersionInfo = shared.findSheetVersion(newSheetID, "Home Page", sheetType);

    if (!newVersionInfo || !newVersionInfo.currentVersion) {
      console.log(`Could not find new sheet version`);
      return {
        success: false,
        message: `Could not find new sheet™ version`,
      };
    }
    var newVersion = newVersionInfo.currentVersion;
    var oldVersion = oldFile.name.match(/[vV]\d+(?:.\d+)*/g);

    console.log(JSON.stringify(oldVersion));
    var newFileName = newFile.name;
    if (oldVersion && oldVersion.length > 0 && newVersion) {
      newFileName = oldFile.name.replace(oldVersion[0], newVersion);
    } else if (newVersion) {
      newFileName = `${oldFile.name} ${newVersion}`;
    }
    if (sheetType === "IDS Collection") {
      newFileName = newFileName.replace("IDS Master", "IDS Collection");
    }
    console.log(
      `Updating file name from "${oldFile.name}" to "${newFileName}"`
    );

    parents = {};
    if (typeof(oldFile.parents) == "undefined") {
      console.log(`Could not find old file location.`);
      return {
        success: false,
        message: `Could not find old file location.`,
      };
    }

    parents["addParents"] = oldFile.parents.join(",");
    
    if (typeof(newFile.parents) != "undefined") {
      parents["removeParents"] = newFile.parents.join(",");
    }

    try {
      Drive.Files.update(
        {
          name: newFileName,
        },
        newSheetID,
        null,
        parents
      );
    } catch (error) {
      console.log(`Error moving new sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error moving new sheet™: ${error.toString()}`,
      };
    }

    try {
      Drive.Files.update({ trashed: true }, oldSheetID);
    } catch (error) {
      console.log(`Error deleting old sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error deleting old sheet™: ${error.toString()}`,
      };
    }

    return {
      success: true,
      message: "new sheet™ moved and renamed, old sheet™ deleted",
      newName: newFileName,
    };
  } catch (error) {
    console.log(`Error in moveSheet: ${error.toString()}`);
    return {
      success: false,
      message: error.toString(),
    };
  }
}

function updateIdsMaster(idMasterID, idDataEntries) {
  var idsMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
  if (!idsMasterSpreadsheet) {
    console.log(`IDS Master Spreadsheet not found with ID: ${idMasterID}`);
    return {
      success: false,
      message: `IDS Master Spreadsheet™ not found with ID: ${idMasterID}`,
    };
  }

  var idMasterIDSheet = SheetsAPI.getSheetByName(idsMasterSpreadsheet, "IDS");
  if (!idMasterIDSheet) {
    console.log(`IDS sheet not found in ID master spreadsheet`);
    return {
      success: false,
      message: `IDS sheet™ not found in ID master spreadsheet™`,
    };
  }
  try {
    var batchUpdate = [];
    var idsMasterData = SheetsAPI.batchGetValues(idMasterID, ["IDS"]);
    var idsMasterValues = idsMasterData[0].values;
    for (var i = 0; i < idDataEntries.length; i++) {
      var idData = idDataEntries[i];
      var sheetType = idData.sheetType;
      if (sheetType === "IDS Master") {
        continue; // Skip IDS Master entries
      }
      var idMasterSpreadsheetInfo = shared.findSheetTypeID(
        idMasterID,
        "IDS",
        sheetType,
        idsMasterValues
      );
      if (
        !idMasterSpreadsheetInfo ||
        !idMasterSpreadsheetInfo.cell ||
        !idMasterSpreadsheetInfo.cell.range
      ) {
        console.log(`IDS Master cell range not found`);
        return {
          success: false,
          message: `IDS Master cell range not found`,
        };
      }
      var cellRange = idMasterSpreadsheetInfo.cell.range;
      var values = [[idData.newSheetID]];
      batchUpdate.push({
        range: cellRange,
        values: values,
      });
    }

    if (batchUpdate.length > 0) {
      try {
        SheetsAPI.batchUpdateValues(idMasterID, [batchUpdate]);
      } catch (error) {
        console.log(`Error updating ID Master sheet: ${error.toString()}`);
        return {
          success: false,
          message: `Error updating ID Master sheet™: ${error.toString()}`,
        };
      }
    }
    return {
      success: true,
      message: "New IDS Master set successfully",
      gid: idMasterIDSheet.sheetId,
    };
  } catch (error) {
    console.log(`Error in setNewIdsMaster: ${error.toString()}`);
    return {
      success: false,
      message: error.toString(),
      gid: idMasterIDSheet.sheetId,
    };
  }
}

function checkImportStatusAndCompatibility(newSheetID, oldSheetID, sheetType) {
  try {
    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID
    );
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`,
        imported: false,
      };
    }

    // Determine sheet configuration based on sheet type
    var sheetName = "IDS";
    var searchName = "IDS Master's";
    var requiresIDSSheet = true;

    if (sheetType === "IDS Collection") {
      sheetName = "Home Page";
      searchName = "Load your file here";
      requiresIDSSheet = false;
    }

    // Check for required sheets based on sheet type
    if (requiresIDSSheet) {
      var newIdSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
      if (!newIdSheet) {
        console.log(`IDS sheet not found in new ${sheetType} spreadsheet.`);
        return {
          success: false,
          message: `IDS sheet™ not found in new ${sheetType} spreadsheet™.`,
          imported: false,
        };
      }
    }

    // Check for Effective Paths specific sheets
    if (sheetType === "Effective Paths") {
      var eHPSheet = SheetsAPI.getSheetByName(newSpreadsheet, "eHP");
      var eDamageSheet = SheetsAPI.getSheetByName(newSpreadsheet, "eDamage");
      var eEconSheet = SheetsAPI.getSheetByName(newSpreadsheet, "eEcon");

      if (!eHPSheet || !eDamageSheet || !eEconSheet) {
        console.log(
          `Required Effective Paths sheets not found in new spreadsheet.`
        );
        return {
          success: false,
          message: `New spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).`,
          imported: false,
        };
      }
    }

    var newHomePageSheet = SheetsAPI.getSheetByName(
      newSpreadsheet,
      "Home Page"
    );
    if (!newHomePageSheet) {
      console.log(`Home Page sheet not found in new ${sheetType} spreadsheet`);
      return {
        success: false,
        message: `Home Page sheet™ not found in new ${sheetType} spreadsheet™`,
        imported: false,
      };
    }

    // Batch fetch all required data from new sheet in one API call
    var ranges = requiresIDSSheet ? ["IDS", "Home Page"] : ["Home Page"];
    var newSheetData = SheetsAPI.batchGetValues(newSheetID, ranges);
    if (!newSheetData || newSheetData.length < ranges.length) {
      console.log(`Could not fetch required data from new sheet`);
      return {
        success: false,
        message: `Could not fetch required data from new sheet`,
        imported: false,
      };
    }

    var newIDSValues = newSheetData[0].values;
    var newHomePageValues = requiresIDSSheet
      ? newSheetData[1].values
      : newSheetData[0].values;

    // STEP 1: Check compatibility first (for early failure)
    var newVersionInfo;
    newVersionInfo = shared.findSheetVersion(
      newSheetID,
      "Home Page",
      sheetType,
      newHomePageValues
    );

    if (!newVersionInfo) {
      console.log(`Version not found in new ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `Version not found in new ${sheetType} spreadsheet™.`,
        imported: false,
      };
    }

    var newVersion = newVersionInfo.currentVersion;
    var latestVersion = newVersionInfo.latestVersion;
    if (!newVersion || !latestVersion) {
      console.log(
        `Version information is incomplete in new ${sheetType} spreadsheet.`
      );
      return {
        success: false,
        message: `Version information is incomplete in new ${sheetType} spreadsheet™.`,
        imported: false,
      };
    }

    if (newVersion !== latestVersion) {
      console.log(
        `The version of the new sheet (${newVersion}) is not the latest version (${latestVersion}). Please update before importing.`
      );
      return {
        success: false,
        message: `The version of the new sheet (${newVersion}) is not the latest version (${latestVersion}). Please update before importing.`,
        imported: false,
      };
    }

    // Check old sheet compatibility
    var oldSpreadsheet = spreadsheets(
      `${sheetType} oldSpreadsheet`,
      oldSheetID
    );
    if (!oldSpreadsheet) {
      console.log(`Old spreadsheet not found with ID: ${oldSheetID}`);
      return {
        success: false,
        message: `Old spreadsheet™ not found with ID: ${oldSheetID}`,
        imported: false,
      };
    }

    var oldHomePageSheet = SheetsAPI.getSheetByName(
      oldSpreadsheet,
      "Home Page"
    );
    if (!oldHomePageSheet) {
      console.log(`Home Page sheet not found in old ${sheetType} spreadsheet`);
      return {
        success: false,
        message: `Home Page sheet™ not found in old ${sheetType} spreadsheet™`,
        imported: false,
      };
    }

    // Fetch old sheet Home Page data
    var oldHomePageData = SheetsAPI.batchGetValues(oldSheetID, ["Home Page"]);
    if (!oldHomePageData || oldHomePageData.length === 0) {
      console.log(`Could not fetch Home Page data from old sheet`);
      return {
        success: false,
        message: `Could not fetch Home Page data from old sheet`,
        imported: false,
      };
    }

    var oldHomePageValues = oldHomePageData[0].values;
    var versionDifference = "None";
    if (oldHomePageValues[1][1] !== "IDS Master") {
      var oldVersionInfo;
      oldVersionInfo = shared.findSheetVersion(
        oldSheetID,
        "Home Page",
        sheetType,
        oldHomePageValues
      );

      if (!oldVersionInfo || !oldVersionInfo.currentVersion) {
        console.log(`Version not found in old ${sheetType} spreadsheet.`);
        return {
          success: false,
          message: `Version not found in old ${sheetType} spreadsheet™.`,
          imported: false,
        };
      }

      var oldVersion = oldVersionInfo.currentVersion;
      var compareVersions = shared.compareVersions(oldVersion, newVersion);

      if (compareVersions === "newer") {
        console.log(
          `The version of the old sheet (${oldVersion}) is newer than the new sheet (${newVersion}). Import aborted.`
        );
        return {
          success: false,
          message: `The version of the old sheet (${oldVersion}) is newer than the new sheet (${newVersion}). Import aborted.`,
          imported: false,
        };
      }

      var sheetTypeFunction = sheetVars(sheetType);
      if (sheetTypeFunction) {
        versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
        if (!versionDifference) {
          console.log(
            `Old version of ${sheetType} is incompatible for import (${oldVersion})`
          );
          return {
            success: false,
            message: `Old version of ${sheetType} is incompatible for import (${oldVersion}).`,
            imported: false,
          };
        }
      } else {
        console.log(
          `No compatibility function found for ${sheetType}. Assuming incompatible.`
        );
        return {
          success: false,
          message: `No compatibility function found for ${sheetType}. Assuming incompatible.`,
          imported: false,
        };
      }
    }

    // STEP 2: If compatible, check import status
    var newSpreadsheetInfo = shared.findSheetTypeID(
      newSheetID,
      sheetName,
      searchName,
      newIDSValues
    );
    if (!newSpreadsheetInfo || !newSpreadsheetInfo.importStatus) {
      console.log(`Can not find import cell in the new IDS sheet.`);
      return {
        success: false,
        message: `Can not find import cell in the new IDS sheet.`,
        imported: false,
      };
    }

    var importStatusValue = newSpreadsheetInfo.importStatus.value;
    var isImported = importStatusValue === "✅";

    if (isImported) {
      return {
        success: true,
        message: "Data is already imported.",
        imported: true,
        versionDifference: versionDifference,
      };
    } else {
      return {
        success: true,
        message: `Data is not imported yet. Ready for import.`,
        imported: false,
        versionDifference: versionDifference,
      };
    }
  } catch (error) {
    console.log(
      `Error checking import status and compatibility: ${error.message}`
    );
    return {
      success: false,
      message: `Error checking import status and compatibility: ${error.message}`,
      imported: false,
    };
  }
}

function getOAuthToken() {
  try {
    const token = ScriptApp.getOAuthToken();
    return {
      success: true,
      token: token,
      message: "Token retrieved successfully",
    };
  } catch (error) {
    console.error("Error getting OAuth token:", error);
    return {
      success: false,
      token: null,
      message: error.toString(),
    };
  }
}

function checkSheetAccess(sheetID, userEmail) {
  try {
    if (!sheetID || !userEmail) {
      return {
        success: false,
        message: "Missing sheetID or userEmail parameter",
        accessible: false,
        owned: false,
      };
    }

    try {
      const file = Drive.Files.get(sheetID, {
        fields: "name, owners",
      });

      const owners = file.owners || [];
      const isOwner = owners.some(
        (owner) =>
          owner.emailAddress &&
          owner.emailAddress.toLowerCase() === userEmail.toLowerCase()
      );

      return {
        success: true,
        message: `Sheet access verified`,
        accessible: true,
        owned: isOwner,
        sheetID: sheetID,
        name: file.name,
      };
    } catch (error) {
      console.log(`Sheet access denied for ${sheetID}: ${error.toString()}`);

      return {
        success: true,
        message: `Sheet access denied`,
        accessible: false,
        owned: false,
        sheetID: sheetID,
      };
    }
  } catch (error) {
    console.error(`Error checking sheet access: ${error.toString()}`);
    return {
      success: false,
      message: `Error checking sheet access: ${error.toString()}`,
      accessible: false,
      owned: false,
    };
  }
}

function getTemplateAndsheetIds(idMasterID, copyMode) {
  try {
    const sheetTypes = [
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

    copyMode = copyMode || "all";
    console.log(
      `Getting template and old sheet IDs for IDS Master: ${idMasterID}, mode: ${copyMode}`
    );

    // Fetch IDS Master data once
    var idsMasterData = fetchIdsMasterData(idMasterID);
    if (!idsMasterData.success) {
      console.log(`Error fetching IDS Master data: ${idsMasterData.message}`);
      return {
        success: false,
        message: `Error fetching IDS Master data: ${idsMasterData.message}`,
        collection: idsMasterData.collection || false,
      };
    }

    var templateInfo = [];
    var sheetIds = [idMasterID];

    // Process all sheet types using the pre-fetched data
    for (var i = 0; i < sheetTypes.length; i++) {
      var sheetType = sheetTypes[i];
      try {
        var templateResult = getTemplateInfo(
          idsMasterData,
          sheetType,
          copyMode
        );

        if (templateResult && templateResult.success) {
          // Skip if version filtering excluded this template
          if (templateResult.versionFiltered) {
            console.log(`Skipping ${sheetType} - version filtering applied`);
            continue;
          }

          templateInfo.push({
            templateID: templateResult.templateID,
            sheetType: sheetType,
            templateVersion: templateResult.templateVersion,
            oldVersion: templateResult.oldVersion,
            oldSheetID: templateResult.oldSheetID,
          });

          // Collect old sheet IDs
          if (templateResult.oldSheetID) {
            sheetIds.push(templateResult.oldSheetID);
          }
        } else {
          console.log(
            `Error getting template info for ${sheetType}: ${
              templateResult ? templateResult.message : "Unknown error"
            }`
          );
        }
      } catch (templateError) {
        console.log(
          `Error processing template for ${sheetType}: ${templateError.toString()}`
        );
      }
    }

    return {
      success: true,
      sheetIds: sheetIds,
      templateInfo: templateInfo,
      message: `Found ${templateInfo.length} templates and ${sheetIds.length} old sheets to check`,
    };
  } catch (error) {
    console.log(
      `Error getting template and old sheet IDs: ${error.toString()}`
    );
    return {
      success: false,
      message: `Error getting template and old sheet IDs: ${error.message}`,
    };
  }
}

// Helper function to get template information without checking access
function getTemplateInfo(idsMasterData, sheetType, copyMode) {
  try {
    var values = idsMasterData.values;
    var formulas = idsMasterData.formulas;
    var idMasterID = idsMasterData.idMasterID;
    copyMode = copyMode || "all";

    var spreadsheetInfo = shared.findSheetTypeURL(
      idMasterID,
      "IDS",
      sheetType,
      values
    );

    if (!spreadsheetInfo || !spreadsheetInfo.template) {
      console.log(`Could not find sheet template for ${sheetType}`);
      return {
        success: false,
        message: `Could not find sheet template for ${sheetType}`,
      };
    }

    if (!spreadsheetInfo.id) {
      console.log(
        `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`
      );
      return {
        success: false,
        message: `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
      };
    }

    var oldSheetID = shared.extractSheetId(spreadsheetInfo.id);
    if (!oldSheetID) {
      console.log(`Could not extract old sheet ID from ${spreadsheetInfo.id}`);
      return {
        success: false,
        message: `Could not extract old sheet™ ID from ${spreadsheetInfo.id}`,
      };
    }

    var templateVersion = spreadsheetInfo.version.value;
    var oldVersion = spreadsheetInfo.oldVersion.value;

    // Version filtering logic - only process if 'update' mode and template is newer
    if (copyMode === "update") {
      if (!templateVersion || !oldVersion) {
        console.log(
          `Version information missing for ${sheetType} - template: ${templateVersion}, old: ${oldVersion}`
        );
        return {
          success: true,
          versionFiltered: true,
          message: `Version information missing for ${sheetType}`,
        };
      }

      var versionComparison = shared.compareVersions(
        oldVersion,
        templateVersion
      );
      if (versionComparison !== "older") {
        console.log(
          `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}, skipping`
        );
        return {
          success: true,
          versionFiltered: true,
          message: `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}`,
        };
      }

      console.log(
        `${sheetType} template version ${templateVersion} is newer than old version ${oldVersion}, including`
      );
    }

    var templateRow = spreadsheetInfo.template.row - 1; // Convert to 0-based
    var templateCol = spreadsheetInfo.template.col - 1; // Convert to 0-based

    // Get template URL from pre-fetched formulas
    var templateUrl = "";
    if (
      formulas &&
      formulas[templateRow] &&
      formulas[templateRow][templateCol]
    ) {
      templateUrl = shared.extractUrlFromHyperlink(
        formulas[templateRow][templateCol]
      );
    }

    if (!templateUrl) {
      console.log(`Template URL not found for ${sheetType}`);
      return {
        success: false,
        message: `Template URL not found for ${sheetType}`,
      };
    }

    var templateID = shared.extractSheetId(templateUrl);
    if (!templateID) {
      console.log(`Could not extract template ID from URL: ${templateUrl}`);
      return {
        success: false,
        message: `Could not extract template ID from URL: ${templateUrl}`,
      };
    }

    return {
      success: true,
      templateID: templateID,
      oldSheetID: oldSheetID,
      templateVersion: templateVersion,
      oldVersion: oldVersion,
      sheetType: sheetType,
      message: `Successfully got template info for ${sheetType}`,
    };
  } catch (error) {
    console.error(
      `Error getting template info for ${sheetType}: ${error.toString()}`
    );
    return { success: false, message: `${error.toString()}` };
  }
}

// Simplified function to check access for a single template
function checkTemplateAccess(templateID) {
  try {
    if (!templateID) {
      return {
        success: false,
        message: "Missing templateID parameter",
        accessible: false,
      };
    }

    try {
      var file = Drive.Files.get(templateID, {
        fields: "id, name",
      });

      return {
        success: true,
        message: `Template access verified`,
        accessible: true,
        templateID: templateID,
        name: file.name,
      };
    } catch (error) {
      console.log(
        `Template access denied for ${templateID}: ${error.toString()}`
      );

      return {
        success: true,
        message: `Template access denied`,
        accessible: false,
        templateID: templateID,
      };
    }
  } catch (error) {
    console.error(`Error checking template access: ${error.toString()}`);
    return {
      success: false,
      message: `Error checking template access: ${error.toString()}`,
      accessible: false,
    };
  }
}

function findSheetIdAndType(sheetID, sheetType) {
  if (!sheetID) {
    console.log(`Missing sheetId parameter.`);
    return { error: "Missing sheetType parameter." };
  }
  sheetType = sheetType || "IDS Master's";
  var spreadsheetInfo = shared.findSheetTypeID(sheetID, "IDS", sheetType);
  if (!spreadsheetInfo || !spreadsheetInfo.id) {
    console.log(
      `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`
    );
    return {
      success: false,
      message: `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
    };
  }
  console.log(`Found sheet type ID: ${spreadsheetInfo.id}`);
  var spreadsheetId = shared.extractSheetId(spreadsheetInfo.id);
  if (!spreadsheetId) {
    console.log(`Could not extract sheet ID from ${spreadsheetInfo.id}`);
    return {
      success: false,
      message: `Could not extract sheet™ ID from ${spreadsheetInfo.id}`,
    };
  }
  if (!sheetType || sheetType === "IDS Master's") {
    var sheetTypeResult = SheetsAPI.batchGetValues(sheetID, ["Home Page!B2"]);
    sheetType = sheetTypeResult[0].values[0][0];
  }

  return {
    success: true,
    sheetID: spreadsheetId,
    sheetType: sheetType,
  };
}

// Optimized function that fetches all IDS Master data in one API call
function fetchIdsMasterData(idMasterID) {
  try {
    if (!idMasterID) {
      console.log(`Missing idMasterID parameter.`);
      return {
        success: false,
        message: "Missing idMasterID parameter.",
      };
    }

    var idMasterSheet = spreadsheets("idMasterSpreadsheet", idMasterID);
    if (!idMasterSheet) {
      console.log(`IDS Master file not found with ID: ${idMasterID}`);
      return {
        success: false,
        message: `IDS Master file not found with ID: ${idMasterID}`,
      };
    }

    var idMasterSheetInfo = SheetsAPI.getSheetByName(idMasterSheet, "IDS");
    if (!idMasterSheetInfo) {
      console.log(`IDS sheet not found in the IDS Master file.`);
      return {
        success: false,
        message: `IDS sheet not found in the IDS Master file.`,
        collection: true,
      };
    }

    // Two API calls to get both values and formulas from IDS sheet
    var idsValues = SheetsAPI.batchGetValues(idMasterID, ["IDS"]);
    var idsFormulas = SheetsAPI.batchGetFormulas(idMasterID, ["IDS"]);

    if (!idsValues || !idsValues[0] || !idsValues[0].values) {
      console.log(`Could not read IDS sheet data from IDS Master.`);
      return {
        success: false,
        message: `Could not read IDS sheet data from IDS Master.`,
      };
    }

    if (!idsFormulas || !idsFormulas[0] || !idsFormulas[0].values) {
      console.log(`Could not read IDS sheet formulas from IDS Master.`);
      return {
        success: false,
        message: `Could not read IDS sheet formulas from IDS Master.`,
      };
    }

    return {
      success: true,
      spreadsheet: idMasterSheet,
      sheetInfo: idMasterSheetInfo,
      values: idsValues[0].values,
      formulas: idsFormulas[0].values,
      idMasterID: idMasterID,
    };
  } catch (error) {
    console.error(`Error fetching IDS Master data: ${error.toString()}`);
    return { success: false, message: `${error.toString()}` };
  }
}

// Process template access for a single sheet type using pre-fetched data
function processTemplateAccess(idsMasterData, sheetType, copyMode) {
  try {
    var values = idsMasterData.values;
    var formulas = idsMasterData.formulas;
    var idMasterID = idsMasterData.idMasterID;
    copyMode = copyMode || "all";

    var spreadsheetInfo = shared.findSheetTypeURL(
      idMasterID,
      "IDS",
      sheetType,
      values
    );

    if (!spreadsheetInfo || !spreadsheetInfo.template) {
      console.log(`Could not find sheet template for ${sheetType}`);
      return {
        success: false,
        message: `Could not find sheet template for ${sheetType}`,
      };
    }

    if (!spreadsheetInfo.id) {
      console.log(
        `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`
      );
      return {
        success: false,
        message: `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
      };
    }

    var oldSheetID = shared.extractSheetId(spreadsheetInfo.id);
    if (!oldSheetID) {
      console.log(`Could not extract old sheet ID from ${spreadsheetInfo.id}`);
      return {
        success: false,
        message: `Could not extract old sheet™ ID from ${spreadsheetInfo.id}`,
      };
    }

    var templateVersion = spreadsheetInfo.version.value;
    var oldVersion = spreadsheetInfo.oldVersion.value;

    // Version filtering logic - only process if 'update' mode and template is newer
    if (copyMode === "update") {
      if (!templateVersion || !oldVersion) {
        console.log(
          `Version information missing for ${sheetType} - template: ${templateVersion}, old: ${oldVersion}`
        );
        return {
          success: true,
          versionFiltered: true,
          message: `Version information missing for ${sheetType}`,
        };
      }

      var versionComparison = shared.compareVersions(
        oldVersion,
        templateVersion
      );
      if (versionComparison !== "older") {
        console.log(
          `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}, skipping`
        );
        return {
          success: true,
          versionFiltered: true,
          message: `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}`,
        };
      }

      console.log(
        `${sheetType} template version ${templateVersion} is newer than old version ${oldVersion}, including`
      );
    }

    var templateRow = spreadsheetInfo.template.row - 1; // Convert to 0-based
    var templateCol = spreadsheetInfo.template.col - 1; // Convert to 0-based

    // Get template URL from pre-fetched formulas instead of making API call
    var templateUrl = "";
    if (
      formulas &&
      formulas[templateRow] &&
      formulas[templateRow][templateCol]
    ) {
      templateUrl = shared.extractUrlFromHyperlink(
        formulas[templateRow][templateCol]
      );
    }

    if (!templateUrl) {
      console.log(`Template URL not found for ${sheetType}`);
      return {
        success: false,
        message: `Template URL not found for ${sheetType}`,
      };
    }

    var templateID = shared.extractSheetId(templateUrl);
    if (!templateID) {
      console.log(`Could not extract template ID from URL: ${templateUrl}`);
      return {
        success: false,
        message: `Could not extract template ID from URL: ${templateUrl}`,
      };
    }

    // Check template access without creating a copy
    try {
      var file = Drive.Files.get(templateID, {
        fields: "id",
      });

      // Template is accessible, return success with template information
      return {
        success: true,
        message: `Template access verified for ${sheetType}.`,
        accessDenied: false,
        templateID: templateID,
        templateVersion: templateVersion,
        oldVersion: oldVersion,
        oldFileId: oldSheetID,
        idMasterFileId: idMasterID,
        sheetType: sheetType,
      };
    } catch (error) {
      console.log(
        `Error retrieving template file information: ${error.toString()}`
      );
      console.log(`Template ID: ${templateID}, Sheet Type: ${sheetType}`);

      // Template access denied, return information needed for granting access
      return {
        success: true,
        message: `Template access check completed. Access needed for ${sheetType} template.`,
        accessDenied: true,
        templateID: templateID,
        templateVersion: templateVersion,
        oldVersion: oldVersion,
        oldFileId: oldSheetID,
        idMasterFileId: idMasterID,
        sheetType: sheetType,
      };
    }
  } catch (error) {
    console.error(
      `Error processing template access for ${sheetType}: ${error.toString()}`
    );
    return { success: false, message: `${error.toString()}` };
  }
}

function checkFileTemplateAccess(idMasterID, sheetType) {
  // For backward compatibility, use the optimized approach for single calls
  var idsMasterData = fetchIdsMasterData(idMasterID);
  if (!idsMasterData.success) {
    return idsMasterData;
  }

  return processTemplateAccess(idsMasterData, sheetType, "all");
}

function copyFileTemplate(templateID, sheetType, templateVersion) {
  try {
    var fileName = `Copy of ${sheetType} ${templateVersion}`;

    var newFile = Drive.Files.copy({ name: fileName }, templateID, {
      fields: "id",
    });

    if (!newFile || !newFile.id) {
      console.log(`Error copying ${sheetType} template: no file returned`);
      return {
        success: false,
        message: `Error copying ${sheetType} template: no file returned`,
      };
    }

    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newFile.id
    );

    if (sheetType === "IDS Master") {
      console.log(`Copied IDS Master template, no further setup needed.`);
      return {
        success: true,
        message: `Successfully copied IDS Master template.`,
        fileId: newFile.id,
        fileName: fileName,
        gid: "",
      };
    }

    var newSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
    if (!newSheet) {
      console.log(`IDS sheet not found in Copy of ${sheetType} spreadsheet.`);
      return {
        success: true,
        message: `IDS sheet™ not found in Copy of ${sheetType} spreadsheet™.`,
        fileId: newFile.id,
        fileName: fileName,
        gid: "",
      };
    }

    // Single API call to get all IDS sheet data at once
    var idsValues = SheetsAPI.batchGetValues(newFile.id, ["IDS"]);
    if (!idsValues || !idsValues[0] || !idsValues[0].values) {
      console.log(
        `Could not read IDS sheet data from new ${sheetType} template`
      );
      return {
        success: true,
        message: `Could not read IDS sheet data from new ${sheetType} template`,
        fileId: newFile.id,
        fileName: fileName,
        gid: newSheet.sheetId,
      };
    }

    var values = idsValues[0].values;

    var thisSheetInfo = shared.findSheetTypeID(
      newFile.id,
      "IDS",
      "This Sheet ID",
      values
    );

    if (!thisSheetInfo || !thisSheetInfo.cell) {
      console.log(
        `Could not find 'This Sheet ID' entry in new ${sheetType} template`
      );
      return {
        success: true,
        message: `Could not find 'This Sheet ID' entry in new ${sheetType} template`,
        fileId: newFile.id,
        fileName: fileName,
        gid: newSheet.sheetId,
      };
    }

    // Note: ID setting is now handled in importData to reduce API calls during copying
    console.log(
      `Successfully copied ${sheetType} template, IDs will be set during import.`
    );

    return {
      success: true,
      message: `Successfully copied ${sheetType} template.`,
      fileId: newFile.id,
      fileName: fileName,
      gid: newSheet.sheetId,
    };
  } catch (error) {
    console.error(`Error copying ${sheetType} template: ${error.toString()}`);
    return { success: false, message: `${error.toString()}` };
  }
}

function checkNewSheetReference(newSheetID, sheetType) {
  try {
    if (!newSheetID) {
      console.log(`Missing newSheetID parameter.`);
      return { success: false, message: "Missing newSheetID parameter." };
    }

    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID
    );
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`,
      };
    }

    var newIdSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
    if (!newIdSheet) {
      console.log(`IDS sheet not found in the new spreadsheet.`);
      return {
        success: false,
        message: `IDS sheet™ not found in the new spreadsheet™.`,
      };
    }

    var newSpreadsheetInfo = shared.findSheetTypeID(newSheetID, "IDS");
    if (!newSpreadsheetInfo || !newSpreadsheetInfo.id) {
      console.log(`Could not find sheet type ID for ${newSheetID}`);
      return {
        success: false,
        message: `Could not find sheet type ID for ${newSheetID}`,
      };
    }
    var accessStatus = newSpreadsheetInfo.accessStatus;
    if (!accessStatus || accessStatus.value !== "✅") {
      console.log(`New sheet have not been granted access to IDS Master.`);
      return {
        success: false,
        message: `New sheet have not been granted access to IDS Master.`,
      };
    }

    return {
      success: true,
      message: `New sheet reference is valid.`,
    };
  } catch (error) {
    console.error(`Error checking new sheet reference: ${error.toString()}`);
    return { success: false, message: `${error.toString()}` };
  }
}

// Prepare data for IDS Master import to be executed in parallel on client side
function prepareImportData(
  idMasterID,
  copiedTemplateFiles,
  importedFilesFailed,
  exportedFilesFailed
) {
  try {
    // Ensure all arrays exist and combine them
    copiedTemplateFiles = copiedTemplateFiles || [];
    importedFilesFailed = importedFilesFailed || [];
    exportedFilesFailed = exportedFilesFailed || [];

    var allTemplateFiles = copiedTemplateFiles
      .concat(importedFilesFailed)
      .concat(exportedFilesFailed);

    console.log(
      `Preparing parallel IDS Master import data for ${allTemplateFiles.length} template files (${copiedTemplateFiles.length} copied + ${importedFilesFailed.length} import failed + ${exportedFilesFailed.length} export failed)`
    );

    // Get the IDS Master data once (single API call for values)
    var idsValues = SheetsAPI.batchGetValues(idMasterID, ["IDS"]);
    if (!idsValues || !idsValues[0] || !idsValues[0].values) {
      return {
        success: false,
        message: `Could not read IDS sheet data from IDS Master`,
        succeededTasks: [],
        failedTasks: [],
      };
    }

    var values = idsValues[0].values;
    var succeededTasks = [];
    var failedTasks = [];

    // Prepare all import tasks with required information
    for (var i = 0; i < allTemplateFiles.length; i++) {
      var templateFile = allTemplateFiles[i];
      var sheetType = templateFile.sheetType;
      var newSheetID = templateFile.fileId;

      // Get the sheet type info using pre-loaded values
      var sheetTypeInfo = shared.findSheetTypeURL(
        idMasterID,
        "IDS",
        sheetType,
        values
      );
      if (!sheetTypeInfo || !sheetTypeInfo.id) {
        failedTasks.push({
          sheetType: sheetType,
          success: false,
          message: `Could not find old sheet information for ${sheetType}`,
        });
        continue;
      }

      var oldSheetID = shared.extractSheetId(sheetTypeInfo.id);
      if (!oldSheetID) {
        failedTasks.push({
          sheetType: sheetType,
          success: false,
          message: `Could not extract old sheet ID for ${sheetType}`,
        });
        continue;
      }

      var oldVersion = sheetTypeInfo.oldVersion.value;
      var templateVersion = sheetTypeInfo.version.value;

      console.log(
        `oldVersion: ${oldVersion}, templateVersion: ${templateVersion}`
      );
      var versionDifference = null;
      if (oldVersion && templateVersion) {
        var versionComparison = shared.compareVersions(
          oldVersion,
          templateVersion
        );
        if (versionComparison === "newer") {
          failedTasks.push({
            sheetType: sheetType,
            success: false,
            message: `Old version (${oldVersion}) is newer than template version (${templateVersion})`,
          });
          continue;
        }

        // Get version difference from sheet type function
        var sheetTypeFunction = sheetVars(sheetType);
        if (sheetTypeFunction) {
          versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
          if (!versionDifference) {
            failedTasks.push({
              sheetType: sheetType,
              success: false,
              message: `Old version of ${sheetType} is incompatible for import`,
            });
            continue;
          }
        }
      }

      succeededTasks.push({
        sheetType: sheetType,
        newSheetID: newSheetID,
        oldSheetID: oldSheetID,
        idMasterID: idMasterID,
        versionDifference: versionDifference,
      });
    }

    return {
      success: true,
      message: `Prepared import data for ${succeededTasks.length} tasks`,
      succeededTasks: succeededTasks,
      failedTasks: failedTasks,
    };
  } catch (error) {
    console.error(`Error preparing import data:`, error);
    return {
      success: false,
      message: `Error preparing import data: ${error.toString()}`,
      succeededTasks: [],
      failedTasks: [],
    };
  }
}

function deleteOldSheet(sheetID) {
  try {
    console.log(`Attempting to delete sheet with ID: ${sheetID}`);

    // Get file information to verify it exists before deletion
    var fileInfo;
    try {
      fileInfo = Drive.Files.get(sheetID, { fields: "id, name, trashed" });
    } catch (error) {
      console.log(
        `Sheet ${sheetID} not found or already deleted: ${error.toString()}`
      );
      return {
        success: true,
        message: `Sheet was already deleted or not found: ${sheetID}`,
      };
    }

    // Check if file is already trashed
    if (fileInfo.trashed) {
      console.log(`Sheet ${sheetID} (${fileInfo.name}) is already trashed`);
      return {
        success: true,
        message: `Sheet "${fileInfo.name}" was already deleted`,
      };
    }

    // Delete the file by moving it to trash
    Drive.Files.update({ trashed: true }, sheetID);

    console.log(`Successfully deleted sheet: ${fileInfo.name} (${sheetID})`);
    return {
      success: true,
      message: `Successfully deleted sheet: "${fileInfo.name}"`,
    };
  } catch (error) {
    console.log(`Error deleting sheet ${sheetID}: ${error.toString()}`);
    return {
      success: false,
      message: `Error deleting sheet: ${error.toString()}`,
    };
  }
}

function getTemplateIdForSingleSheet(sheetID, sheetType) {
  try {
    if (!sheetID) {
      console.log(`Missing sheetID parameter.`);
      return { success: false, message: "Missing sheetID parameter." };
    }
    if (!sheetType) {
      console.log(`Missing sheetType parameter.`);
      return { success: false, message: "Missing sheetType parameter." };
    }
    var spreadsheetInfo = shared.findSheetTemplateID(
      sheetID,
      "Home Page",
      sheetType
    );
    if (
      !spreadsheetInfo ||
      !spreadsheetInfo.templateID ||
      !spreadsheetInfo.templateVersion
    ) {
      console.log(`Could not find sheet template for ${sheetType}`);
      return {
        success: false,
        message: `Could not find sheet template for ${sheetType}`,
      };
    }
    return {
      success: true,
      templateID: spreadsheetInfo.templateID,
      templateVersion: spreadsheetInfo.templateVersion,
      message: `Successfully got template ID for ${sheetType}`,
    };
  } catch (error) {
    console.error(
      `Error getting template ID for single sheet: ${error.toString()}`
    );
    return { success: false, message: `${error.toString()}` };
  }
}

function checkExportCompatibility(oldSheetID, sheetType) {
  try {
    if (!oldSheetID) {
      return { success: false, message: "Missing oldSheetID parameter." };
    }
    if (!sheetType) {
      return { success: false, message: "Missing sheetType parameter." };
    }

    var oldSpreadsheet = spreadsheets(
      `${sheetType} oldSpreadsheet`,
      oldSheetID
    );
    if (!oldSpreadsheet) {
      return {
        success: false,
        message: `Could not access old ${sheetType} spreadsheet with ID: ${oldSheetID}`,
      };
    }

    var oldHomePageSheet = SheetsAPI.getSheetByName(
      oldSpreadsheet,
      "Home Page"
    );
    if (!oldHomePageSheet) {
      return {
        success: false,
        message: `Home Page sheet not found in old ${sheetType} spreadsheet`,
      };
    }

    // Fetch old sheet Home Page data
    var oldHomePageData = SheetsAPI.batchGetValues(oldSheetID, ["Home Page"]);
    if (!oldHomePageData || oldHomePageData.length === 0) {
      return {
        success: false,
        message: `Could not read Home Page data from old ${sheetType} spreadsheet`,
      };
    }

    var oldHomePageValues = oldHomePageData[0].values;
    var oldVersionInfo;
    var oldVersion;

    oldVersionInfo = shared.findSheetVersion(
      oldSheetID,
      "Home Page",
      sheetType,
      oldHomePageValues
    );
    if (!oldVersionInfo || !oldVersionInfo.currentVersion) {
      return {
        success: false,
        message: `Current Version not found in old ${sheetType} spreadsheet.`,
      };
    }
    oldVersion = oldVersionInfo.currentVersion;

    // Check compatibility using the sheetType's compatibility function
    var sheetTypeFunction = sheetVars(sheetType);
    if (sheetTypeFunction) {
      var versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
      if (!versionDifference) {
        return {
          success: false,
          message: `Old version of ${sheetType} is incompatible for export (${oldVersion})`,
        };
      }
      return {
        success: true,
        message: `Old ${sheetType} version (${oldVersion}) is compatible for export`,
        oldVersion: oldVersion,
        versionDifference: versionDifference,
      };
    } else {
      return {
        success: false,
        message: `No compatibility function found for ${sheetType}. Cannot verify export compatibility.`,
      };
    }
  } catch (error) {
    console.error(`Error checking export compatibility: ${error.toString()}`);
    return {
      success: false,
      message: `Error checking export compatibility: ${error.toString()}`,
    };
  }
}

function updateSheetID(spreadsheetID, sheetID, sheetType) {
  try {
    if (!spreadsheetID) {
      return { success: false, message: "Missing spreadsheetID parameter." };
    }
    if (!sheetID) {
      return { success: false, message: "Missing sheetID parameter." };
    }
    if (!sheetType) {
      return { success: false, message: "Missing sheetType parameter." };
    }
    var spreadsheet = spreadsheets(`${sheetType} spreadsheet`, spreadsheetID);
    if (!spreadsheet) {
      return {
        success: false,
        message: `Could not access spreadsheet with ID: ${spreadsheetID}`,
      };
    }
    var idSheet = SheetsAPI.getSheetByName(spreadsheet, "IDS");
    if (!idSheet) {
      return {
        success: false,
        message: `IDS sheet not found in spreadsheet with ID: ${spreadsheetID}`,
      };
    }
    var idValues = SheetsAPI.batchGetValues(spreadsheetID, ["IDS"]);
    if (!idValues || idValues.length === 0) {
      return {
        success: false,
        message: `Could not read IDS sheet data from spreadsheet with ID: ${spreadsheetID}`,
      };
    }
    var values = idValues[0].values;
    var sheetTypeInfo = shared.findSheetTypeID(
      spreadsheetID,
      "IDS",
      "IDS Master",
      values
    );
    if (!sheetTypeInfo || !sheetTypeInfo.cell) {
      return {
        success: false,
        message: `Could not find IDS Master entry in IDS sheet.`,
      };
    }
    var currentSheetID = shared.extractSheetId(sheetTypeInfo.id);
    if (currentSheetID === sheetID) {
      return {
        success: true,
        message: `Sheet ID is already up to date in IDS sheet.`,
      };
    }
    SheetsAPI.batchUpdateValues(spreadsheetID, [
      {
        range: sheetTypeInfo.cell.range,
        values: [[sheetID]],
      },
    ]);
    return {
      success: true,
      message: `Successfully updated IDS Master ID in IDS sheet.`,
    };
  } catch (error) {
    console.error(`Error updating sheet ID: ${error.toString()}`);
    return {
      success: false,
      message: `Error updating sheet ID: ${error.toString()}`,
    };
  }
}
