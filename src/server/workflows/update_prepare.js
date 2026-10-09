/**
 * Client-callable. Resolves the old sheet and version for an import.
 * @param {string} newSheetID
 * @param {string} sheetType
 * @param {string} idMasterID
 * @returns {{success: boolean, oldSheetID: string, versionDifference: string}} A failure envelope on error.
 */
function prepareImportData(
  idMasterID,
  copiedTemplateFiles,
  importedFilesFailed,
  exportedFilesFailed,
) {
  try {
    copiedTemplateFiles = copiedTemplateFiles || [];
    importedFilesFailed = importedFilesFailed || [];
    exportedFilesFailed = exportedFilesFailed || [];

    var allTemplateFiles = copiedTemplateFiles
      .concat(importedFilesFailed)
      .concat(exportedFilesFailed);

    console.log(
      `Preparing parallel IDS Master import data for ${allTemplateFiles.length} template files (${copiedTemplateFiles.length} copied + ${importedFilesFailed.length} import failed + ${exportedFilesFailed.length} export failed)`,
    );

    var idsValues = SheetsAPI.batchGetValues(idMasterID, ["IDS"]);
    if (!idsValues || !idsValues[0] || !idsValues[0].values) {
      return errors.reject(
        "prepareImportData",
        errors.CODES.SHEET_STRUCTURE,
        `Could not read IDS sheet data from IDS Master`,
        {
          succeededTasks: [],
          failedTasks: [],
        },
      );
    }

    var values = idsValues[0].values;
    var succeededTasks = [];
    var failedTasks = [];

    for (var i = 0; i < allTemplateFiles.length; i++) {
      var templateFile = allTemplateFiles[i];
      var sheetType = templateFile.sheetType;
      var newSheetID = templateFile.fileId;

      var sheetTypeInfo = labelUtils.findSheetTypeURL(
        idMasterID,
        "IDS",
        sheetType,
        values,
      );
      if (!sheetTypeInfo || !sheetTypeInfo.id) {
        failedTasks.push(
          errors.reject(
            "prepareImportData",
            errors.CODES.SHEET_STRUCTURE,
            `Could not find ${sheetType} in your IDS Master's IDS tab.`,
            { sheetType: sheetType },
          ),
        );
        continue;
      }

      var oldSheetID = sheetRefs.extractSheetId(sheetTypeInfo.id);
      if (!oldSheetID) {
        failedTasks.push(
          errors.reject(
            "prepareImportData",
            errors.CODES.SHEET_STRUCTURE,
            `The ${sheetType} entry in your IDS Master's IDS tab is not a usable sheet link.`,
            { sheetType: sheetType },
            { id: sheetTypeInfo.id },
          ),
        );
        continue;
      }

      var oldVersion = versionUtils.readVersion(sheetTypeInfo.oldVersion.value);
      var templateVersion = versionUtils.readVersion(sheetTypeInfo.version.value);

      console.log(
        `oldVersion: ${oldVersion}, templateVersion: ${templateVersion}`,
      );
      var versionDifference = null;
      if (oldVersion && templateVersion) {
        var versionComparison = versionUtils.compareVersions(
          oldVersion,
          templateVersion,
        );
        if (versionComparison === "newer") {
          failedTasks.push(
            errors.reject(
              "prepareImportData",
              errors.CODES.VERSION_OUTDATED,
              `Your ${sheetType} (${oldVersion}) is newer than the template it would be imported into (${templateVersion}).`,
              { sheetType: sheetType },
            ),
          );
          continue;
        }

        var sheetTypeFunction = sheetVars(sheetType);
        if (sheetTypeFunction) {
          versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
          if (!versionDifference) {
            failedTasks.push(
              errors.reject(
                "prepareImportData",
                errors.CODES.VERSION_OUTDATED,
                `Your ${sheetType} is version ${oldVersion}, which this script cannot convert to the new template.`,
                { sheetType: sheetType },
              ),
            );
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
    var errorReport = errors.report("prepareImportData", error, {
      note: `Error preparing import data`,
      idMasterID: idMasterID,
      copiedTemplateFiles: copiedTemplateFiles,
      importedFilesFailed: importedFilesFailed,
      exportedFilesFailed: exportedFilesFailed,
    });
    return errors.fail(errorReport, null, {
      succeededTasks: [],
      failedTasks: [],
    });
  }
}
