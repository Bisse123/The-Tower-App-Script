const SheetsAPI = {
  // Get spreadsheet metadata and sheets
  getSpreadsheet: function (spreadsheetId) {
    try {
      return Sheets.Spreadsheets.get(spreadsheetId);
    } catch (error) {
      console.error(`Error getting spreadsheet: ${error}`);
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

  // Get values from a range
  getValues: function (spreadsheetId, range) {
    try {
      const response = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
      return response.values;
    } catch (error) {
      console.error(`Error getting values: ${error}`);
      return null;
    }
  },

  // Get a single value from a cell
  getValue: function (spreadsheetId, range) {
    try {
      const values = this.getValues(spreadsheetId, range);
      return values && values.length > 0 && values[0].length > 0 ? values[0][0] : null;
    } catch (error) {
      console.error(`Error getting single value: ${error}`);
      return null;
    }
  },

  // Set multiple values in a range
  setValues: function (spreadsheetId, range, values) {
    try {
      const requestBody = {
        values: values,
      };
      return Sheets.Spreadsheets.Values.update(
        requestBody,
        spreadsheetId,
        range,
        {
          valueInputOption: "USER_ENTERED",
        }
      );
    } catch (error) {
      console.error(`Error setting values: ${error}`);
      return null;
    }
  },

  // Set a single value in a cell
  setValue: function (spreadsheetId, range, value) {
    try {
      const requestBody = {
        values: [[value]],
      };
      return Sheets.Spreadsheets.Values.update(
        requestBody,
        spreadsheetId,
        range,
        {
          valueInputOption: "USER_ENTERED",
        }
      );
    } catch (error) {
      console.error(`Error setting value: ${error}`);
      return null;
    }
  },

  // Get all data from a sheet
  getDataRange: function (spreadsheetId, sheetName) {
    try {
      return this.getValues(spreadsheetId, sheetName);
    } catch (error) {
      console.error(`Error getting data range: ${error}`);
      return null;
    }
  },

  batchUpdateValues: function (spreadsheetId, updates) {
    try {
      const data = updates.map(update => ({
        range: update.range,
        values: update.values
      }));
      const requestBody = {
        data: data,
        valueInputOption: "USER_ENTERED"
      };
      return Sheets.Spreadsheets.Values.batchUpdate(requestBody, spreadsheetId);
    } catch (error) {
      console.error(`Error in batchUpdateValues: ${error}`);
      return null;
    }
  },
};

const shared = {
  compareVersions: function (oldVersion, newVersion) {
    var oldVersionNumber = parseInt(oldVersion.replace(/\D/g, ""), 10);
    var newVersionNumber = parseInt(newVersion.replace(/\D/g, ""), 10);

    if (oldVersionNumber > newVersionNumber) return -1;
    if (oldVersionNumber < newVersionNumber) return oldVersionNumber;
    return 0;
  },

  findSheetTypeID: function (spreadsheetId, sheetName, idType) {
    var idType = idType || "IDS Master's ID";
    var values = SheetsAPI.getDataRange(spreadsheetId, sheetName);
    if (!values || values.length === 0) {
      console.log(`No data found in sheet: ${sheetName} in spreadsheet: ${spreadsheetId}`);
      return null;
    }

    var regex = new RegExp(idType, "i");
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (regex.test(values[i][j]) && values[i][j].indexOf("script") === -1) {
          var cellA1 = shared.columnToLetter(j + 2) + (i + 1);
          var importedA1 = shared.columnToLetter(j + 4) + (i + 2);
          return {
            id: values[i][j + 2],
            cell: {
              row: i + 1,
              col: j + 2,
              range: sheetName + "!" + cellA1,
            },
            isImported: {
              row: i + 2,
              col: j + 4,
              range: sheetName + "!" + importedA1,
              value:
                SheetsAPI.getValue(
                  spreadsheetId,
                  sheetName + "!" + importedA1
                ) || "",
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

  splitNameAndVersion: function (sheetName) {
    var lastSpace = sheetName.lastIndexOf(" ");
    if (lastSpace === -1) {
      return { base: sheetName, version: "" };
    }
    var base = sheetName.substring(0, lastSpace);
    var version = sheetName.substring(lastSpace + 1);
    return { base: base, version: version };
  },

  columnToLetter: function(column) {
    var temp = "";
    var letter = "";
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  },
};

function updateSheet(sheetType, newSheetID, oldSheetID, idMasterID) {
  try {
    var newSpreadsheet = spreadsheets("newSpreadsheet", newSheetID);
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet not found with ID: ${newSheetID}`,
        updated: false,
      };
    }
    
    if (!SheetsAPI.getSheetByName(newSpreadsheet, "IDS")) {
      console.log(`IDS sheet not found in new spreadsheet`);
      return {
        success: false,
        message: `IDS sheet not found in new spreadsheet`,
        updated: false,
      };
    }

    var newSheetTypeInfo = shared.findSheetTypeID(newSheetID, "IDS");
    if (
      !newSheetTypeInfo ||
      !newSheetTypeInfo.isImported ||
      newSheetTypeInfo.isImported.value !== "✅"
    ) {
      console.log(`Can not update until old sheet has been Imported.`);
      return {
        success: false,
        message: `Can not update until old sheet has been Imported.`,
        updated: false,
      };
    }
    var idsMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
    if (!idsMasterSpreadsheet) {
      console.log(`IDS Master Spreadsheet not found with ID: ${idMasterID}`);
      return {
        success: false,
        message: `IDS Master Spreadsheet not found with ID: ${idMasterID}`,
        updated: false,
      };
    }
    var idMasterIDSheet = SheetsAPI.getSheetByName(idsMasterSpreadsheet, "IDS");
    if (!idMasterIDSheet) {
      console.log(`IDS sheet not found in ID master spreadsheet`);
      return {
        success: false,
        message: `IDS sheet not found in ID master spreadsheet`,
        updated: false,
      };
    }

    var newFile = Drive.Files.get(newSheetID, { fields: "id, name, parents" });
    var oldFile = Drive.Files.get(oldSheetID, { fields: "id, name, parents" });
    if (!newFile || !oldFile) {
      console.log(`Could not retrieve file information for new or old sheet.`);
      return {
        success: false,
        message: `Could not retrieve file information for new or old sheet.`,
        updated: false,
      };
    }

    var newNameParts = shared.splitNameAndVersion(newFile.name);
    var oldNameParts = shared.splitNameAndVersion(oldFile.name);
    var baseName = oldNameParts.base;
    var newVersion = newNameParts.version;
    var finalName = baseName + (newVersion ? " " + newVersion : "");
    
    var idMasterSpreadsheetInfo = shared.findSheetTypeID(
      idMasterID,
      "IDS",
      sheetType + " ID"
    );
    if (!idMasterSpreadsheetInfo || !idMasterSpreadsheetInfo.cell) {
      console.log(`Could not find ID Master spreadsheet info`);
      return {
        success: false,
        message: `Could not find ID Master spreadsheet info`,
        updated: false,
      };
    }
    idCell = idMasterSpreadsheetInfo.cell;
    if (!idCell.range) {
      console.log(`ID Master cell range not found`);
      return {
        success: false,
        message: `ID Master cell range not found`,
        updated: false,
      };
    }

    try {
      Drive.Files.update(
        {
          name: finalName,
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
        message: `Error renaming or moving new sheet: ${error.toString()}`,
        updated: false,
      };
    }

    try {
      Drive.Files.update({ trashed: true }, oldSheetID);
    } catch (error) {
      console.log(`Error deleting old sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error deleting old sheet: ${error.toString()}`,
        updated: false,
      };
    }

    try {
      SheetsAPI.setValue(idMasterID, idCell.range, newSheetID);
    } catch (error) {
      console.log(`Error updating ID Master sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating ID Master sheet: ${error.toString()}`,
        updated: false,
      };
    }

    return {
      success: true,
      message: "New ID Set, new sheet moved and renamed, old sheet deleted.",
      updated: true,
      gid: idMasterIDSheet.sheetId,
    };
  } catch (error) {
    console.log(`Error in updateSheet: ${error.toString()}`);
    return {
      success: false,
      message: error.toString(),
      updated: false,
    };
  }
}

function checkCompatibility(newSheetID, oldSheetID, idMasterID, sheetType) {
  // console.log(`Checking compatibility for sheet type: ${sheetType}, newSheetID: ${newSheetID}, oldSheetID: ${oldSheetID}, idMasterID: ${idMasterID}`);
  try {
    var newSpreadsheet = spreadsheets("newSpreadsheet", newSheetID);
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet not found with ID: ${newSheetID}`,
      };
    }
    var newExportSheet = SheetsAPI.getSheetByName(newSpreadsheet, "EXPORT");
    if (!newExportSheet) {
      newExportSheet = SheetsAPI.getSheetByName(newSpreadsheet, "STATS");
    }
    if (!newExportSheet) {
      console.log(`Export sheet not found in new ${sheetType} spreadsheet`);
      return {
        success: false,
        message:
          `Export sheet not found in new ${sheetType} spreadsheet`,
      };
    }

    var newVersion = SheetsAPI.getValue(
      newSheetID,
      `${newExportSheet.title}!A1`
    );
    if (!newVersion) {
      console.log(`Export version not found in new ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `Export version not found in new ${sheetType} spreadsheet.`,
      };
    }

    var oldSpreadsheet = spreadsheets("oldSpreadsheet", oldSheetID);
    if (!oldSpreadsheet) {
      console.log(`Old spreadsheet not found with ID: ${oldSheetID}`);
      return {
        success: false,
        message: `Old spreadsheet not found with ID: ${oldSheetID}`,
      };
    }
    var oldExportSheet = SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT");
    if (!oldExportSheet) {
      oldExportSheet = SheetsAPI.getSheetByName(oldSpreadsheet, "STATS");
    }
    if (!oldExportSheet) {
      console.log(`Export sheet not found in old ${sheetType} spreadsheet`);
      return {
        success: false,
        message: `Export sheet not found in old ${sheetType} spreadsheet`,
      };
    }

    var oldVersion = SheetsAPI.getValue(
      oldSheetID,
      `${oldExportSheet.title}!A1`
    );
    if (!oldVersion) {
      console.log(`Export version not found in old ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `Export version not found in old ${sheetType} spreadsheet.`,
      };
    }

    var compareVersions = shared.compareVersions(oldVersion, newVersion);

    if (compareVersions === -1) {
      console.log(`The version of the old sheet (${oldVersion}) is newer than the new sheet (${newVersion}). Import aborted.`);
      return {
        success: false,
        message: `The version of the old sheet (${oldVersion}) is newer than the new sheet (${newVersion}). Import aborted.`,
      };
    } else if (compareVersions > 0) {
      var sheetTypeFunction = sheetVars(sheetType);
      if (
        sheetTypeFunction &&
        sheetTypeFunction.isCompatibleVersion(compareVersions)
      ) {
        return {
          success: true,
          message: `The version of the old sheet (${oldVersion}) is compatible with the new sheet (${newVersion}).`,
          versionDifference: oldVersion,
        };
      }
      return {
        success: false,
        message: `Old version of ${sheetType} is incompatible import.`,
      };
    }
    return {
      success: true,
      message: `The version of the old sheet (${oldVersion}) is the same as the new sheet (${newVersion}).`,
      versionDifference: 0,
    };
  } catch (error) {
    console.log(`Error checking compatibility: ${error.message}`);
    return {
      success: false,
      message: "Error checking compatibility: " + error.message,
    };
  }
}

function checkImportStatus(newSheetID) {
  try {
    var newSpreadsheet = spreadsheets("newSpreadsheet", newSheetID);
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet not found with ID: ${newSheetID}`,
      };
    }

    var newIdSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
    if (!newIdSheet) {
      console.log(`IDS sheet not found in new ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `IDS sheet not found in new ${sheetType} spreadsheet.`,
      };
    }

    var newSpreadsheetInfo = shared.findSheetTypeID(newSheetID, "IDS");
    if (!newSpreadsheetInfo || !newSpreadsheetInfo.isImported) {
      console.log(`Can not find import cell in the new IDS sheet.`);
      return {
        success: false,
        message: `Can not find import cell in the new IDS sheet.`,
      };
    }

    var isImportedValue = newSpreadsheetInfo.isImported.value;

    if (isImportedValue === "✅") {
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

function checkSheetAccess(fileIds, userEmail) {
  try {
    // const fileIds = [newSheetID, oldsheetID, idMasterID];
    const accessibleFiles = [];
    const inaccessibleFiles = [];
    const notOwnedFiles = [];
    
    // console.log("Checking access to", fileIds.length, "predefined sheets...");

    fileIds.forEach((fileId) => {
      try {
        const file = Drive.Files.get(fileId, {
          fields: "name, owners",
        });
        
        // console.log("Access confirmed for:", file.name, "(" + fileId + ")");

        const owners = file.owners || [];
        const isOwner = owners.some(owner => owner.emailAddress && owner.emailAddress.toLowerCase() === userEmail.toLowerCase());
        
        if (isOwner) {
          accessibleFiles.push({id: fileId, name: file.name});
        } else {
          notOwnedFiles.push({id: fileId, name: file.name});
          inaccessibleFiles.push({id: fileId});
        }

      } catch (error) {
        // console.log("No access to file:", fileId, "Error:", error.toString());
        inaccessibleFiles.push({id: fileId});
      }
    });

    // console.log(
    //   "Access check complete. Accessible:",
    //   accessibleFiles.length,
    //   "Inaccessible:",
    //   inaccessibleFiles.length
    // );

    return {
      success: true,
      accessibleFiles: accessibleFiles,
      inaccessibleFiles: inaccessibleFiles,
      notOwnedFiles: notOwnedFiles,
      message: `Access check complete. ${accessibleFiles.length} of ${fileIds.length} sheets are accessible.`,
    };
  } catch (error) {
    console.error("Error checking sheet access:", error);
    return {
      success: false,
      accessibleFiles: [],
      inaccessibleFiles: [],
      notOwnedFiles: [],
      message: error.toString(),
    };
  }
}

function findSheetIdAndType(sheetID, sheetType) {
  if (!sheetID) {
    console.log(`Missing sheetId parameter.`);
    return { error: "Missing sheetType parameter." };
  }
  var idType = sheetType ? sheetType + " ID" : "IDS Master's ID";
  var spreadsheetInfo = shared.findSheetTypeID(sheetID, "IDS", idType);
  if (!spreadsheetInfo || !spreadsheetInfo.id) {
    console.log(`Could not find sheet type ID for ${sheetID}`);
    return {error: `Could not find sheet type ID for ${sheetID}`};
  }
  
  var spreadsheetId = shared.extractSheetId(
    spreadsheetInfo.id
  );
  if (!spreadsheetId) {
    console.log(`Could not extract sheet ID from ${spreadsheetInfo.id}`);
    return {error: `Could not extract sheet ID from ${spreadsheetInfo.id}`};
  }
  if (!sheetType) {
    sheetType = SheetsAPI.getValue(sheetID, "Home Page!B2");
  }

  return {
    sheetID: spreadsheetId,
    sheetType: sheetType,
  };
}