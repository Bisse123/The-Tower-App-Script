/**
 * Client-callable. Moves an existing file into the Get Started folder.
 * @param {string} fileId
 * @param {string} parentFolderID
 * @returns {{success: boolean, fileId: string}} A failure envelope on error.
 */
function moveGetStartedFileToFolder(fileId, parentFolderID) {
  try {
    if (!fileId) {
      return errors.reject(
        "moveGetStartedFileToFolder",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing fileId parameter." },
      );
    }
    if (!parentFolderID) {
      return errors.reject(
        "moveGetStartedFileToFolder",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing parentFolderID parameter." },
      );
    }

    var file = CacheManager.getFile(fileId);
    if (!file) {
      return errors.reject(
        "moveGetStartedFileToFolder",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `File not found for ID: ${fileId}` },
      );
    }

    var versionInfo = versionUtils.findSheetVersion(
      fileId,
      "Home Page",
      "Effective Paths",
    );
    var versionLabel =
      versionInfo && versionInfo.currentVersion
        ? String(versionInfo.currentVersion).trim()
        : "";
    var newFileName = "Effective Paths";
    if (versionLabel) {
      newFileName = `${newFileName} ${versionLabel}`.trim();
    }

    var removeParents = "";
    if (file.parents && file.parents.length > 0) {
      removeParents = file.parents.join(",");
    }

    var parents = {
      addParents: parentFolderID,
    };
    if (removeParents) {
      parents.removeParents = removeParents;
    }

    Drive.Files.update({ name: newFileName }, fileId, null, parents);

    return {
      success: true,
      message: "File moved to Get Started folder.",
      fileId: fileId,
      fileName: newFileName,
      fileUrl: `https://docs.google.com/spreadsheets/d/${fileId}/edit`,
    };
  } catch (error) {
    var errorReport = errors.report("moveGetStartedFileToFolder", error, {
      note: `Error moving Get Started file (${fileId})`,
      fileId: fileId,
      parentFolderID: parentFolderID,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Client-callable. The Tower Drive folder, creating it if needed.
 * @returns {{success: boolean, id: string, name: string, url: string}} A failure envelope on error.
 */
function getOrCreateGetStartedFolder() {
  try {
    var query =
      'name="The Tower" and mimeType="application/vnd.google-apps.folder" and trashed=false';
    var folderList = Drive.Files.list({
      q: query,
      spaces: "drive",
      fields: "files(id, name)",
      pageSize: 1,
    });

    if (folderList.files && folderList.files.length > 0) {
      var folder = folderList.files[0];
      console.log(`Found existing "The Tower" folder: ${folder.id}`);
      return {
        success: true,
        id: folder.id,
        name: folder.name,
        url: `https://drive.google.com/drive/folders/${folder.id}`,
      };
    }

    var fileMetadata = {
      name: "The Tower",
      mimeType: "application/vnd.google-apps.folder",
    };

    var newFolder = Drive.Files.create(fileMetadata, null, {
      fields: "id, name",
    });

    Drive.Permissions.create(
      {
        role: "reader",
        type: "anyone",
      },
      newFolder.id,
    );

    console.log(`Created new "The Tower" folder: ${newFolder.id}`);
    return {
      success: true,
      id: newFolder.id,
      name: newFolder.name,
      url: `https://drive.google.com/drive/folders/${newFolder.id}`,
    };
  } catch (error) {
    var errorReport = errors.report("getOrCreateGetStartedFolder", error, { note: `Error getting or creating The Tower folder` });
    return errors.fail(errorReport, null, {
      id: "",
      name: "",
      url: "",
      error: `Error locating or creating The Tower folder: ${error.toString()}`,
    });
  }
}

/**
 * Client-callable. Links newly created sheets to each other and renames
 * them with their version.
 * @param {string} sheetID
 * @param {string} sheetType
 * @param {Array<{sheetType: string, sheetID: string}>} [relatedSheetIDs]
 * @returns {{success: boolean, message: string, updatedCount: number}} A failure envelope on error.
 */
function updateGetStartedSheetIdsAndReferences(
  sheetID,
  sheetType,
  relatedSheetIDs,
) {
  try {
    console.log(
      `Updating IDs for ${sheetType} (${sheetID}) with related IDs:`,
      relatedSheetIDs,
    );

    var fileName = null;
    var updatedCount = 0;

    if (sheetType === "IDS Collection") {
      var idsResult = SheetsAPI.batchGetValues(sheetID, ["Home Page"]);
      if (!idsResult || !idsResult[0] || !idsResult[0].values) {
        return errors.reject(
          "updateGetStartedSheetIdsAndReferences",
          errors.CODES.SHEET_STRUCTURE,
          "Could not fetch Home Page",
        );
      }

      var idsData = idsResult[0].values;
      var ownSheetInfo = labelUtils.findSheetTypeID(
        sheetID,
        "Home Page",
        "Your ID:",
        idsData,
      );

      if (!ownSheetInfo || !ownSheetInfo.cell || !ownSheetInfo.cell.range) {
        return errors.reject(
          "updateGetStartedSheetIdsAndReferences",
          errors.CODES.SHEET_STRUCTURE,
          "Could not find 'Your ID:' in IDS Collection",
        );
      }

      var batchUpdate = [
        {
          range: ownSheetInfo.cell.range,
          values: [[sheetID]],
        },
      ];

      SheetsAPI.batchUpdateValues(sheetID, batchUpdate);
      updatedCount = 1;

      try {
        var sheetInfo = versionUtils.findSheetVersion(
          sheetID,
          "Home Page",
          sheetType,
          idsData,
        );

        if (sheetInfo && sheetInfo.currentVersion) {
          fileName = `${sheetType} ${sheetInfo.currentVersion}`;
          Drive.Files.update(
            {
              name: fileName,
            },
            sheetID,
          );
        }
      } catch (error) {
        errors.report("updateGetStartedSheetIdsAndReferences", error, {
          note: `Could not update file name with version info`,
          sheetID: sheetID,
          sheetType: sheetType,
          relatedSheetIDs: relatedSheetIDs,
        }, errors.CODES.RECOVERED);
      }

      return {
        success: true,
        message: "Updated IDS Collection ID",
        updatedCount: updatedCount,
        fileName: fileName,
      };
    } else if (sheetType === "IDS Master") {
      var idDataEntries = [];
      relatedSheetIDs.forEach((entry) => {
        idDataEntries.push({
          sheetType: entry.sheetType,
          newSheetID: entry.sheetID,
        });
      });

      var idsResult = SheetsAPI.batchGetValues(sheetID, ["Home Page", "IDS"]);
      var homePageData = idsResult[0].values;
      var masterResult = updateIdsMaster(sheetID, idDataEntries);
      updatedCount = idDataEntries.length;

      try {
        var sheetInfo = versionUtils.findSheetVersion(
          sheetID,
          "Home Page",
          sheetType,
          homePageData,
        );

        if (sheetInfo && sheetInfo.currentVersion) {
          fileName = `${sheetType} ${sheetInfo.currentVersion}`;
          Drive.Files.update(
            {
              name: fileName,
            },
            sheetID,
          );
        }
      } catch (error) {
        errors.report("updateGetStartedSheetIdsAndReferences", error, {
          note: `Could not update file name with version info`,
          sheetID: sheetID,
          sheetType: sheetType,
          relatedSheetIDs: relatedSheetIDs,
        }, errors.CODES.RECOVERED);
      }

      return {
        success: masterResult.success,
        message: masterResult.message,
        updatedCount: updatedCount,
        fileName: fileName,
      };
    }

    var idsResult = SheetsAPI.batchGetValues(sheetID, ["Home Page", "IDS"]);
    if (!idsResult || !idsResult[0] || !idsResult[0].values) {
      return errors.reject(
        "updateGetStartedSheetIdsAndReferences",
        errors.CODES.SHEET_STRUCTURE,
        "Could not fetch IDS sheet",
      );
    }

    var homePageData = idsResult[0].values;
    var idsData = idsResult[1].values;
    var batchUpdate = [];

    var idsMasterId = relatedSheetIDs
      .filter((entry) => entry.sheetType === "IDS Master")
      .map((entry) => entry.sheetID)[0];

    batchUpdate = labelUtils.addIDUpdatesToBatch(
      batchUpdate,
      sheetType,
      sheetID,
      idsData,
      idsMasterId,
    );

    if (batchUpdate.length > 0) {
      SheetsAPI.batchUpdateValues(sheetID, batchUpdate);
      updatedCount = batchUpdate.length;
    }

    try {
      var sheetInfo = versionUtils.findSheetVersion(
        sheetID,
        "Home Page",
        sheetType,
        homePageData,
      );

      if (sheetInfo && sheetInfo.currentVersion) {
        fileName = `${sheetType} ${sheetInfo.currentVersion}`;
        Drive.Files.update(
          {
            name: fileName,
          },
          sheetID,
        );
      }
    } catch (error) {
      errors.report("updateGetStartedSheetIdsAndReferences", error, {
        note: `Could not update file name with version info`,
        sheetID: sheetID,
        sheetType: sheetType,
        relatedSheetIDs: relatedSheetIDs,
      }, errors.CODES.RECOVERED);
    }

    return {
      success: true,
      message:
        updatedCount > 0
          ? `Updated ${updatedCount} ID(s)`
          : "No IDs needed updating",
      updatedCount: updatedCount,
      fileName: fileName,
    };
  } catch (error) {
    var errorReport = errors.report("updateGetStartedSheetIdsAndReferences", error, {
      note: `Error updating sheet IDs`,
      sheetID: sheetID,
      sheetType: sheetType,
      relatedSheetIDs: relatedSheetIDs,
    });
    return errors.fail(errorReport);
  }
}
