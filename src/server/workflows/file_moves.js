/**
 * Moves the new sheet into the old one's folder and trashes the old.
 * @param {string} sheetType
 * @param {string} newSheetID
 * @param {string} oldSheetID
 * @param {string[]} [mergedOldSheetIDs] Extra sheets to trash.
 * @returns {{success: boolean, message: string}} A failure envelope on error.
 */
function moveSheet(sheetType, newSheetID, oldSheetID, mergedOldSheetIDs) {
  try {
    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID,
    );
    if (!newSpreadsheet) {
      return errors.reject(
        "moveSheet",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `New spreadsheet™ not found with ID: ${newSheetID}` },
      );
    }

    var newFile = CacheManager.getFile(newSheetID);
    var oldFile = CacheManager.getFile(oldSheetID);
    if (!newFile || !oldFile) {
      return errors.reject(
        "moveSheet",
        errors.CODES.INTERNAL,
        `Could not retrieve file information for new or old sheet™.`,
      );
    }

    var newVersionInfo;
    newVersionInfo = versionUtils.findSheetVersion(
      newSheetID,
      "Home Page",
      sheetType,
    );

    if (!newVersionInfo || !newVersionInfo.currentVersion) {
      return errors.reject(
        "moveSheet",
        errors.CODES.SHEET_STRUCTURE,
        `Could not find new sheet™ version`,
      );
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

    if (
      sheetType === "Themes, Songs & Relics" &&
      newFileName.indexOf(sheetType) === -1
    ) {
      if (newFileName.indexOf("Themes & Songs") !== -1) {
        newFileName = newFileName.replace("Themes & Songs", sheetType);
      } else if (newFileName.indexOf("Relics") !== -1) {
        newFileName = newFileName.replace("Relics", sheetType);
      }
    }
    console.log(
      `Updating file name from "${oldFile.name}" to "${newFileName}"`,
    );

    parents = {};
    if (typeof oldFile.parents == "undefined") {
      return errors.reject(
        "moveSheet",
        errors.CODES.SHEET_STRUCTURE,
        `Could not find old file location.`,
      );
    }

    parents.addParents = oldFile.parents.join(",");

    if (typeof newFile.parents != "undefined") {
      parents.removeParents = newFile.parents.join(",");
    }

    try {
      Drive.Files.update(
        {
          name: newFileName,
        },
        newSheetID,
        null,
        parents,
      );
    } catch (error) {
      var errorReport = errors.report("moveSheet", error, {
        note: `Error moving new sheet`,
        sheetType: sheetType,
        newSheetID: newSheetID,
        oldSheetID: oldSheetID,
        mergedOldSheetIDs: mergedOldSheetIDs,
      });
      return errors.fail(errorReport);
    }

    try {
      Drive.Files.update({ trashed: true }, oldSheetID);
    } catch (error) {
      var errorReport = errors.report("moveSheet", error, {
        note: `Error deleting old sheet`,
        sheetType: sheetType,
        newSheetID: newSheetID,
        oldSheetID: oldSheetID,
        mergedOldSheetIDs: mergedOldSheetIDs,
      });
      return errors.fail(errorReport);
    }

    var extraOldSheetIDs = (mergedOldSheetIDs || []).filter(function (sheetID) {
      return sheetID && sheetID !== oldSheetID;
    });
    for (var i = 0; i < extraOldSheetIDs.length; i++) {
      try {
        Drive.Files.update({ trashed: true }, extraOldSheetIDs[i]);
        CacheManager.RemoveFile(extraOldSheetIDs[i]);
        console.log(`Deleted merged old sheet: ${extraOldSheetIDs[i]}`);
      } catch (error) {
        errors.report("deleteOldSheet", error, {
          note: `Error deleting merged old sheet`,
          sheetID: extraOldSheetIDs[i],
        }, errors.CODES.RECOVERED);
      }
    }

    CacheManager.RemoveSpreadsheet(`${sheetType} newSpreadsheet`);
    CacheManager.RemoveSpreadsheet(`${sheetType} oldSpreadsheet`);
    CacheManager.RemoveFile(newSheetID);
    CacheManager.RemoveFile(oldSheetID);

    return {
      success: true,
      message: "new sheet™ moved and renamed, old sheet™ deleted",
      newName: newFileName,
    };
  } catch (error) {
    var errorReport = errors.report("moveSheet", error, {
      sheetType: sheetType,
      newSheetID: newSheetID,
      oldSheetID: oldSheetID,
      mergedOldSheetIDs: mergedOldSheetIDs,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Moves a sheet converted out of an IDS Collection into place.
 * @param {string} sheetType
 * @param {string} newSheetID
 * @param {string} oldCollectionID
 * @returns {{success: boolean, message: string}} A failure envelope on error.
 */
function moveConvertedSheet(sheetType, newSheetID, oldCollectionID) {
  try {
    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID,
    );
    if (!newSpreadsheet) {
      console.log(`New spreadsheet not found with ID: ${newSheetID}`);
      return {
        success: false,
        message: `New spreadsheet™ not found with ID: ${newSheetID}`,
      };
    }

    var newFile = CacheManager.getFile(newSheetID);
    var collectionFile = CacheManager.getFile(oldCollectionID);
    if (!newFile || !collectionFile) {
      console.log(`Could not retrieve file information for new sheet.`);
      return {
        success: false,
        message: `Could not retrieve file information for new sheet™.`,
      };
    }

    var newVersionInfo = versionUtils.findSheetVersion(
      newSheetID,
      "Home Page",
      sheetType,
    );
    var newVersion =
      newVersionInfo && newVersionInfo.currentVersion
        ? newVersionInfo.currentVersion
        : "";

    var collectionName = collectionFile.name || "";
    var newFileName = collectionName.replace("IDS Collection", sheetType);

    if (newVersion) {
      var existingVersion = newFileName.match(/[vV]\d+(?:.\d+)*/g);
      if (existingVersion && existingVersion.length > 0) {
        newFileName = newFileName.replace(existingVersion[0], newVersion);
      } else {
        newFileName = `${newFileName} ${newVersion}`;
      }
    }

    if (typeof collectionFile.parents == "undefined") {
      console.log(`Could not find old collection file location.`);
      return {
        success: false,
        message: `Could not find old collection file location.`,
      };
    }

    var parents = {
      addParents: collectionFile.parents.join(","),
    };
    if (typeof newFile.parents != "undefined") {
      parents.removeParents = newFile.parents.join(",");
    }

    try {
      Drive.Files.update(
        {
          name: newFileName,
        },
        newSheetID,
        null,
        parents,
      );
    } catch (error) {
      var errorReport = errors.report("moveConvertedSheet", error, {
        note: `Error moving new sheet`,
        sheetType: sheetType,
        newSheetID: newSheetID,
        oldCollectionID: oldCollectionID,
      });
      return errors.fail(errorReport);
    }

    CacheManager.RemoveSpreadsheet(`${sheetType} newSpreadsheet`);
    CacheManager.RemoveFile(newSheetID);

    return {
      success: true,
      message: "new sheet™ moved and renamed",
      newName: newFileName,
    };
  } catch (error) {
    var errorReport = errors.report("moveConvertedSheet", error, {
      sheetType: sheetType,
      newSheetID: newSheetID,
      oldCollectionID: oldCollectionID,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Client-callable. Copies a template into the destination folder.
 * @param {string} templateID
 * @param {string} sheetType
 * @param {string} version
 * @param {string} parentFolderID
 * @returns {{success: boolean, fileId: string}} A failure envelope on error.
 */
function copyFileTemplate(
  templateID,
  sheetType,
  templateVersion,
  parentFolderID,
) {
  try {
    var resolvedTemplateVersion =
      templateVersion && templateVersion !== "undefined"
        ? String(templateVersion).trim()
        : "";
    var fileName = `Copy of ${sheetType} ${resolvedTemplateVersion}`.trim();
    var templateCopyUrl = `https://docs.google.com/spreadsheets/d/${templateID}/copy`;
    var copyRequest = { name: fileName };

    if (parentFolderID != null && parentFolderID !== "") {
      copyRequest.parents = [parentFolderID];
    }

    var newFile = Drive.Files.copy(copyRequest, templateID, {
      fields: "id",
    });

    if (!newFile || !newFile.id) {
      return errors.reject(
        "copyFileTemplate",
        errors.CODES.SHEET_STRUCTURE,
        `Error copying ${sheetType} template: no file returned`,
        {
          copyUrl: templateCopyUrl,
        },
      );
    }

    var fileUrl = `https://docs.google.com/spreadsheets/d/${newFile.id}/edit`;

    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newFile.id,
    );

    idSheet = "IDS";
    if (sheetType === "IDS Collection") {
      idSheet = "Home Page";
    }

    var newSheet = SheetsAPI.getSheetByName(newSpreadsheet, idSheet);
    if (!newSheet) {
      console.log(`${idSheet} sheet not found in ${fileName}`);
      return {
        success: true,
        message: `${idSheet} sheet™ not found in ${fileName}`,
        fileId: newFile.id,
        fileName: fileName,
        fileUrl: fileUrl,
        copyUrl: templateCopyUrl,
        gid: "",
      };
    }
    return {
      success: true,
      message: `Successfully copied ${sheetType} template.`,
      fileId: newFile.id,
      fileName: fileName,
      fileUrl: fileUrl,
      copyUrl: templateCopyUrl,
      gid: newSheet.sheetId,
    };
  } catch (error) {
    var errorReport = errors.report("copyFileTemplate", error, {
      note: `Error copying ${sheetType} template`,
      templateID: templateID,
      sheetType: sheetType,
      templateVersion: templateVersion,
      parentFolderID: parentFolderID,
    });
    return errors.fail(errorReport, null, {
      copyUrl: templateID
        ? "https://docs.google.com/spreadsheets/d/" + templateID + "/copy"
        : "",
    });
  }
}

/**
 * Trashes a sheet, treating already-gone as success.
 * @param {string} sheetID
 * @returns {{success: boolean, message: string}} A failure envelope on error.
 */
function deleteOldSheet(sheetID) {
  try {
    console.log(`Attempting to delete sheet with ID: ${sheetID}`);

    var fileInfo;
    try {
      fileInfo = CacheManager.getFile(sheetID);
    } catch (error) {

      errors.report(
        "deleteOldSheet",
        error,
        { note: `Sheet ${sheetID} not found or already deleted`, sheetID: sheetID }, errors.CODES.NOT_FOUND);
      return {
        success: true,
        message: `Sheet was already deleted or not found: ${sheetID}`,
      };
    }

    if (fileInfo.trashed) {
      console.log(`Sheet ${sheetID} (${fileInfo.name}) is already trashed`);
      return {
        success: true,
        message: `Sheet "${fileInfo.name}" was already deleted`,
      };
    }

    Drive.Files.update({ trashed: true }, sheetID);

    console.log(`Successfully deleted sheet: ${fileInfo.name} (${sheetID})`);
    return {
      success: true,
      message: `Successfully deleted sheet: "${fileInfo.name}"`,
    };
  } catch (error) {
    var errorReport = errors.report("deleteOldSheet", error, {
      note: `Error deleting sheet ${sheetID}`,
      sheetID: sheetID,
    });
    return errors.fail(errorReport);
  }
}
