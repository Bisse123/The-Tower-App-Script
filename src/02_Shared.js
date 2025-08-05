const SheetsAPI = {
  // Get spreadsheet metadata and sheets
  getSpreadsheet: function (spreadsheetId) {
    try {
      const response = Sheets.Spreadsheets.get(spreadsheetId, {
        fields: "spreadsheetId,sheets(properties(sheetId,title))",
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
      const sheet = spreadsheet.sheets.find(
        (s) => s.properties.title.toLowerCase().includes(substring.toLowerCase())
      );
      return sheet ? sheet.properties : null;
    } catch (error) {
      console.error(`Error getting sheet by substring: ${error}`);
      return null;
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
      const data = updates.map((update) => ({
        range: update.range,
        values: update.values,
      }));
      const requestBody = {
        data: data,
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
  findSheetVersion: function (sheetID, sheetName) {
    try {
      var batchResult = SheetsAPI.batchGetValues(sheetID, [sheetName]);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log(
          `No data found in sheet: ${sheetName} in spreadsheet: ${sheetID}`
        );
        return null;
      }
      var values = batchResult[0].values;
      for (var row = 0; row < values.length; row++) {
        var col = values[row].findIndex(
          (cell) => typeof cell === "string" && cell.includes("Version Change")
        );
        if (col !== -1) {
          var version = values[row + 1][col];
          if (version) {
            return version.trim();
          }
        }
      }
    } catch (error) {
      console.error(`Error finding sheet version: ${error}`);
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

  findSheetTypeID: function (spreadsheetId, sheetName, sheetType, preLoadedValues) {
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

  findSheetTypeURL: function (spreadsheetId, sheetName, sheetType, preLoadedValues) {
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
    if (!formula || typeof formula !== 'string') {
      return null;
    }
    
    // Match HYPERLINK formula pattern: =HYPERLINK("url", "text")
    var hyperlinkMatch = formula.match(/=HYPERLINK\s*\(\s*"([^"]+)"/i);
    if (hyperlinkMatch && hyperlinkMatch[1]) {
      return hyperlinkMatch[1];
    }
    
    return null;
  },
};

function updateSheet(sheetType, newSheetID, oldSheetID, idMasterID) {
  try {
    var newSpreadsheet = spreadsheets(`${sheetType} newSpreadsheet`, newSheetID);
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`
      };
    }

    if (!SheetsAPI.getSheetByName(newSpreadsheet, "IDS")) {
      console.log(`IDS sheet not found in new spreadsheet`);
      return {
        success: false,
        message: `IDS sheet™ not found in new spreadsheet™`
      };
    }

    var newSheetTypeInfo = shared.findSheetTypeID(newSheetID, "IDS");
    console.log(`New sheet type info: ${JSON.stringify(newSheetTypeInfo)}`);
    if (
      !newSheetTypeInfo ||
      !newSheetTypeInfo.importStatus ||
      newSheetTypeInfo.importStatus.value !== "✅"
    ) {
      console.log(`Can not update until old sheet has been Imported.`);
      return {
        success: false,
        message: `Can not update until old sheet™ has been Imported.`
      };
    }
    var idsMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
    if (!idsMasterSpreadsheet) {
      console.log(`IDS Master Spreadsheet not found with ID: ${idMasterID}`);
      return {
        success: false,
        message: `IDS Master Spreadsheet™ not found with ID: ${idMasterID}`
      };
    }
    var idMasterIDSheet = SheetsAPI.getSheetByName(idsMasterSpreadsheet, "IDS");
    if (!idMasterIDSheet) {
      console.log(`IDS sheet not found in ID master spreadsheet`);
      return {
        success: false,
        message: `IDS sheet™ not found in ID master spreadsheet™`
      };
    }

    var newFile = Drive.Files.get(newSheetID, { fields: "id, name, parents" });
    var oldFile = Drive.Files.get(oldSheetID, { fields: "id, name, parents" });
    if (!newFile || !oldFile) {
      console.log(`Could not retrieve file information for new or old sheet.`);
      return {
        success: false,
        message: `Could not retrieve file information for new or old sheet™.`
      };
    }

    var idMasterSpreadsheetInfo = shared.findSheetTypeID(
      idMasterID,
      "IDS",
      sheetType
    );
    if (!idMasterSpreadsheetInfo || !idMasterSpreadsheetInfo.cell) {
      console.log(`Could not find ID Master spreadsheet info`);
      return {
        success: false,
        message: `Could not find ID Master spreadsheet™ info`
      };
    }
    idCell = idMasterSpreadsheetInfo.cell;
    if (!idCell.range) {
      console.log(`ID Master cell range not found`);
      return {
        success: false,
        message: `ID Master cell range not found`
      };
    }
    var newVersion = shared.findSheetVersion(newSheetID, "Home Page") || "";
    var oldVersion = oldFile.name.match(/v\d+(?:\.\d+)*/g);

    var newFileName = oldFile.name;
    if (oldVersion && oldVersion.length > 0 && newVersion) {
      newFileName = oldFile.name.replace(oldVersion[0], newVersion);
    } else if (newVersion) {
      newFileName = `${oldFile.name} ${newVersion}`;
    }

    try {
      Drive.Files.update(
        {
          name: newFileName,
        },
        newSheetID,
        null,
        {
          addParents: oldFile.parents.join(","),
          removeParents: newFile.parents.join(","),
        }
      );
    } catch (error) {
      console.log(`Error renaming or moving new sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error renaming or moving new sheet™: ${error.toString()}`
      };
    }

    try {
      Drive.Files.update({ trashed: true }, oldSheetID);
    } catch (error) {
      console.log(`Error deleting old sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error deleting old sheet™: ${error.toString()}`
      };
    }

    try {
      SheetsAPI.batchUpdateValues(idMasterID, [
        {
          range: idCell.range,
          values: [[newSheetID]],
        },
      ]);
    } catch (error) {
      console.log(`Error updating ID Master sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating ID Master sheet™: ${error.toString()}`
      };
    }

    return {
      success: true,
      message: "New ID Set, new sheet™ moved and renamed, old sheet™ deleted.",
      gid: idMasterIDSheet.sheetId,
    };
  } catch (error) {
    console.log(`Error in updateSheet: ${error.toString()}`);
    return {
      success: false,
      message: error.toString(),
    };
  }
}

function checkCompatibility(newSheetID, oldSheetID, sheetType) {
  try {
    var newSpreadsheet = spreadsheets(`${sheetType} newSpreadsheet`, newSheetID);
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`,
      };
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
      };
    }

    var newVersion = shared.findSheetVersion(
      newSheetID,
      newHomePageSheet.title
    );
    if (!newVersion) {
      console.log(`Version not found in new ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `Version not found in new ${sheetType} spreadsheet™.`,
      };
    }

    var oldSpreadsheet = spreadsheets(`${sheetType} oldSpreadsheet`, oldSheetID);
    if (!oldSpreadsheet) {
      console.log(`Old spreadsheet not found with ID: ${oldSheetID}`);
      return {
        success: false,
        message: `Old spreadsheet™ not found with ID: ${oldSheetID}`,
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
      };
    }

    var oldVersion = shared.findSheetVersion(
      oldSheetID,
      oldHomePageSheet.title
    );

    if (!oldVersion) {
      console.log(`Version not found in old ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `Version not found in old ${sheetType} spreadsheet™.`,
      };
    }

    var compareVersions = shared.compareVersions(oldVersion, newVersion);

    if (compareVersions === "newer") {
      console.log(
        `The version of the old sheet (${oldVersion}) is newer than the new sheet (${newVersion}). Import aborted.`
      );
      return {
        success: false,
        message: `The version of the old sheet (${oldVersion}) is newer than the new sheet (${newVersion}). Import aborted.`,
      };
    }
    var sheetTypeFunction = sheetVars(sheetType);
    if (sheetTypeFunction) {
      var versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
      if (!versionDifference) {
        console.log(`Old version of ${sheetType} is incompatible for import.`);
        return {
          success: false,
          message: `Old version of ${sheetType} is incompatible for import.`,
        };
      }
      return {
        success: true,
        message: `The version of the old sheet (${oldVersion}) is compatible with the new sheet (${newVersion}).`,
        versionDifference: versionDifference,
      };
    }
    return {
      success: false,
      message: `Old version of ${sheetType} is incompatible import.`,
    };
  } catch (error) {
    console.log(`Error checking compatibility: ${error.message}`);
    return {
      success: false,
      message: "Error checking compatibility: " + error.message,
    };
  }
}

function checkImportStatus(newSheetID, sheetType) {
  try {
    var newSpreadsheet = spreadsheets(`${sheetType} newSpreadsheet`, newSheetID);
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`,
      };
    }

    var newIdSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
    if (!newIdSheet) {
      console.log(`IDS sheet not found in new ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `IDS sheet™ not found in new ${sheetType} spreadsheet™.`,
      };
    }

    var newSpreadsheetInfo = shared.findSheetTypeID(newSheetID, "IDS");
    if (!newSpreadsheetInfo || !newSpreadsheetInfo.importStatus) {
      console.log(`Can not find import cell in the new IDS sheet.`);
      return {
        success: false,
        message: `Can not find import cell in the new IDS sheet.`,
      };
    }

    var importStatusValue = newSpreadsheetInfo.importStatus.value;

    if (importStatusValue === "✅") {
      return {
        success: true,
        message: "Data is imported.",
        imported: true,
      };
    }
    return {
      success: true,
      message: "Data has not been imported yet.",
      imported: false,
    };
  } catch (error) {
    console.log(`Error checking import status: ${error.message}`);
    return {
      success: false,
      message: `Error checking import status: ${error.message}`,
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
        owned: false
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
        name: file.name
      };
    } catch (error) {
      console.log(`Sheet access denied for ${sheetID}: ${error.toString()}`);
      
      return {
        success: true,
        message: `Sheet access denied`,
        accessible: false,
        owned: false,
        sheetID: sheetID
      };
    }
  } catch (error) {
    console.error(`Error checking sheet access: ${error.toString()}`);
    return { 
      success: false, 
      message: `Error checking sheet access: ${error.toString()}`,
      accessible: false,
      owned: false
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
      "Guardians"
    ];

    copyMode = copyMode || 'all';
    console.log(`Getting template and old sheet IDs for IDS Master: ${idMasterID}, mode: ${copyMode}`);
    
    // Fetch IDS Master data once
    var idsMasterData = fetchIdsMasterData(idMasterID);
    if (!idsMasterData.success) {
      console.log(`Error fetching IDS Master data: ${idsMasterData.message}`);
      return {
        success: false,
        message: `Error fetching IDS Master data: ${idsMasterData.message}`
      };
    }
    
    var templateInfo = [];
    var sheetIds = [idMasterID];
    
    // Process all sheet types using the pre-fetched data
    for (var i = 0; i < sheetTypes.length; i++) {
      var sheetType = sheetTypes[i];
      try {
        var templateResult = getTemplateInfo(idsMasterData, sheetType, copyMode);
        
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
            oldSheetID: templateResult.oldSheetID
          });
          
          // Collect old sheet IDs
          if (templateResult.oldSheetID) {
            sheetIds.push(templateResult.oldSheetID);
          }
        } else {
          console.log(`Error getting template info for ${sheetType}: ${templateResult ? templateResult.message : 'Unknown error'}`);
        }
      } catch (templateError) {
        console.log(`Error processing template for ${sheetType}: ${templateError.toString()}`);
      }
    }
    
    return {
      success: true,
      sheetIds: sheetIds,
      templateInfo: templateInfo,
      message: `Found ${templateInfo.length} templates and ${sheetIds.length} old sheets to check`
    };
    
  } catch (error) {
    console.log(`Error getting template and old sheet IDs: ${error.toString()}`);
    return {
      success: false,
      message: `Error getting template and old sheet IDs: ${error.message}`
    };
  }
}

// Helper function to get template information without checking access
function getTemplateInfo(idsMasterData, sheetType, copyMode) {
  try {
    var values = idsMasterData.values;
    var formulas = idsMasterData.formulas;
    var idMasterID = idsMasterData.idMasterID;
    copyMode = copyMode || 'all';
    
    var spreadsheetInfo = shared.findSheetTypeURL(idMasterID, "IDS", sheetType, values);
    
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
    if (copyMode === 'update') {
      if (!templateVersion || !oldVersion) {
        console.log(`Version information missing for ${sheetType} - template: ${templateVersion}, old: ${oldVersion}`);
        return {
          success: true,
          versionFiltered: true,
          message: `Version information missing for ${sheetType}`,
        };
      }
      
      var versionComparison = shared.compareVersions(oldVersion, templateVersion);
      if (versionComparison !== 'older') {
        console.log(`${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}, skipping`);
        return {
          success: true,
          versionFiltered: true,
          message: `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}`,
        };
      }
      
      console.log(`${sheetType} template version ${templateVersion} is newer than old version ${oldVersion}, including`);
    }
    
    var templateRow = spreadsheetInfo.template.row - 1; // Convert to 0-based
    var templateCol = spreadsheetInfo.template.col - 1; // Convert to 0-based
    
    // Get template URL from pre-fetched formulas
    var templateUrl = "";
    if (formulas && formulas[templateRow] && formulas[templateRow][templateCol]) {
      templateUrl = shared.extractUrlFromHyperlink(formulas[templateRow][templateCol]);
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
      message: `Successfully got template info for ${sheetType}`
    };
  } catch (error) {
    console.error(`Error getting template info for ${sheetType}: ${error.toString()}`);
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
        accessible: false
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
        name: file.name
      };
    } catch (error) {
      console.log(`Template access denied for ${templateID}: ${error.toString()}`);
      
      return {
        success: true,
        message: `Template access denied`,
        accessible: false,
        templateID: templateID
      };
    }
  } catch (error) {
    console.error(`Error checking template access: ${error.toString()}`);
    return { 
      success: false, 
      message: `Error checking template access: ${error.toString()}`,
      accessible: false
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
      idMasterID: idMasterID
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
    copyMode = copyMode || 'all';
    
    var spreadsheetInfo = shared.findSheetTypeURL(idMasterID, "IDS", sheetType, values);
    
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
    if (copyMode === 'update') {
      if (!templateVersion || !oldVersion) {
        console.log(`Version information missing for ${sheetType} - template: ${templateVersion}, old: ${oldVersion}`);
        return {
          success: true,
          versionFiltered: true,
          message: `Version information missing for ${sheetType}`,
        };
      }
      
      var versionComparison = shared.compareVersions(oldVersion, templateVersion);
      if (versionComparison !== 'older') {
        console.log(`${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}, skipping`);
        return {
          success: true,
          versionFiltered: true,
          message: `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}`,
        };
      }
      
      console.log(`${sheetType} template version ${templateVersion} is newer than old version ${oldVersion}, including`);
    }
    
    var templateRow = spreadsheetInfo.template.row - 1; // Convert to 0-based
    var templateCol = spreadsheetInfo.template.col - 1; // Convert to 0-based
    
    // Get template URL from pre-fetched formulas instead of making API call
    var templateUrl = "";
    if (formulas && formulas[templateRow] && formulas[templateRow][templateCol]) {
      templateUrl = shared.extractUrlFromHyperlink(formulas[templateRow][templateCol]);
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
    console.error(`Error processing template access for ${sheetType}: ${error.toString()}`);
    return { success: false, message: `${error.toString()}` };
  }
}

function checkFileTemplateAccess(idMasterID, sheetType) {
  // For backward compatibility, use the optimized approach for single calls
  var idsMasterData = fetchIdsMasterData(idMasterID);
  if (!idsMasterData.success) {
    return idsMasterData;
  }
  
  return processTemplateAccess(idsMasterData, sheetType, 'all');
}

function copyFileTemplate(idMasterID, templateID, sheetType, templateVersion) {
  try {
    var newFile = Drive.Files.copy(
      { name: `Copy of ${sheetType} ${templateVersion}` },
      templateID,
      { fields: "id" }
    );

    if (!newFile || !newFile.id) {
      console.log(`Error copying ${sheetType} template: no file returned`);
      return {
        success: false,
        message: `Error copying ${sheetType} template: no file returned`,
      };
    }

    var newSpreadsheet = spreadsheets(`${sheetType} newSpreadsheet`, newFile.id);

    var newSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
    if (!newSheet) {
      console.log(`IDS sheet not found in Copy of ${sheetType} spreadsheet.`);
      return {
        success: true,
        message: `IDS sheet™ not found in Copy of ${sheetType} spreadsheet™.`,
        fileId: newFile.id,
        gid: "",
      };
    }
    
    // Single API call to get all IDS sheet data at once
    var idsValues = SheetsAPI.batchGetValues(newFile.id, ["IDS"]);
    if (!idsValues || !idsValues[0] || !idsValues[0].values) {
      console.log(`Could not read IDS sheet data from new ${sheetType} template`);
      return {
        success: true,
        message: `Could not read IDS sheet data from new ${sheetType} template`,
        fileId: newFile.id,
        gid: newSheet.sheetId,
      };
    }
    
    var values = idsValues[0].values;
    
    var thisSheetInfo = shared.findSheetTypeID(newFile.id, "IDS", "This Sheet ID", values);
    var idMasterInfo = shared.findSheetTypeID(newFile.id, "IDS", "IDS Master's", values);
    
    if (!thisSheetInfo || !thisSheetInfo.cell) {
      console.log(`Could not find 'This Sheet ID' entry in new ${sheetType} template`);
      return {
        success: true,
        message: `Could not find 'This Sheet ID' entry in new ${sheetType} template`,
        fileId: newFile.id,
        gid: newSheet.sheetId,
      };
    }
    
    if (!idMasterInfo || !idMasterInfo.cell) {
      console.log(`Could not find 'IDS Master's' entry in new ${sheetType} template`);
      return {
        success: true,
        message: `Could not find 'IDS Master's' entry in new ${sheetType} template`,
        fileId: newFile.id,
        gid: newSheet.sheetId,
      };
    }
    
    // Single batch update instead of separate calls
    SheetsAPI.batchUpdateValues(newFile.id, [
      {
        range: thisSheetInfo.cell.range,
        values: [[newFile.id]],
      },
      {
        range: idMasterInfo.cell.range,
        values: [[idMasterID]],
      },
    ]);

    return {
      success: true,
      message: `Successfully copied ${sheetType} template.`,
      fileId: newFile.id,
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

    var newSpreadsheet = spreadsheets(`${sheetType} newSpreadsheet`, newSheetID);
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
function prepareImportData(idMasterID, copiedTemplateFiles) {
  try {
    console.log(`Preparing parallel IDS Master import data for ${copiedTemplateFiles.length} template files`);
    
    // Get the IDS Master data once (single API call for values)
    var idsValues = SheetsAPI.batchGetValues(idMasterID, ["IDS"]);
    if (!idsValues || !idsValues[0] || !idsValues[0].values) {
      return {
        success: false,
        message: `Could not read IDS sheet data from IDS Master`,
        importTasks: [],
        failedTasks: []
      };
    }
    
    var values = idsValues[0].values;
    var importTasks = [];
    var failedTasks = [];
    
    // Prepare all import tasks with required information
    for (var i = 0; i < copiedTemplateFiles.length; i++) {
      var templateFile = copiedTemplateFiles[i];
      var sheetType = templateFile.sheetType;
      var newSheetID = templateFile.fileId;
      
      // Get the sheet type info using pre-loaded values
      var sheetTypeInfo = shared.findSheetTypeURL(idMasterID, "IDS", sheetType, values);
      if (!sheetTypeInfo || !sheetTypeInfo.id) {
        failedTasks.push({
          sheetType: sheetType,
          success: false,
          message: `Could not find old sheet information for ${sheetType}`
        });
        continue;
      }
      
      var oldSheetID = shared.extractSheetId(sheetTypeInfo.id);
      if (!oldSheetID) {
        failedTasks.push({
          sheetType: sheetType,
          success: false,
          message: `Could not extract old sheet ID for ${sheetType}`
        });
        continue;
      }
      
      var oldVersion = sheetTypeInfo.oldVersion.value;
      var templateVersion = sheetTypeInfo.version.value;
      
      console.log(`oldVersion: ${oldVersion}, templateVersion: ${templateVersion}`);
      var versionDifference = null;
      if (oldVersion && templateVersion) {
        var versionComparison = shared.compareVersions(oldVersion, templateVersion);
        if (versionComparison === "newer") {
          failedTasks.push({
            sheetType: sheetType,
            success: false,
            message: `Old version (${oldVersion}) is newer than template version (${templateVersion})`
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
              message: `Old version of ${sheetType} is incompatible for import`
            });
            continue;
          }
        }
      }
      
      importTasks.push({
        sheetType: sheetType,
        newSheetID: newSheetID,
        oldSheetID: oldSheetID,
        idMasterID: idMasterID,
        versionDifference: versionDifference
      });
    }
    
    return {
      success: true,
      message: `Prepared import data for ${importTasks.length} tasks`,
      importTasks: importTasks,
      failedTasks: failedTasks
    };
    
  } catch (error) {
    console.error(`Error preparing import data:`, error);
    return {
      success: false,
      message: `Error preparing import data: ${error.toString()}`,
      importTasks: [],
      failedTasks: []
    };
  }
}
