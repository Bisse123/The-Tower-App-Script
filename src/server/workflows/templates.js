/**
 * Client-callable. Template and sheet IDs for every type in an IDS Master.
 * @param {string} idMasterID
 * @param {string} [copyMode]
 * @returns {{success: boolean, sheetIds: Object, templateInfo: Object}} A failure envelope on error.
 */
function getTemplateAndsheetIds(idMasterID, copyMode) {
  try {

    const sheetTypes = [
      "Laboratory",
      "Workshop",
      "Ultimate Weapon",
      "Themes, Songs & Relics",
      "Themes & Songs",
      "Bots",
      "Relics",
      "Vault",
      "Cards",
      "Modules",
      "Guardians",
      "Player & Stuff",
    ];
    const legacyThemesSheetTypes = ["Themes & Songs", "Relics"];
    var foundMergedThemes = false;

    copyMode = copyMode || "all";
    console.log(
      `Getting template and old sheet IDs for IDS Master: ${idMasterID}, mode: ${copyMode}`,
    );

    var idsMasterData, templateInfo, skippedTemplates, sheetIds, versionLoading;
    var attempt = 0;
    var maxAttempts = 2;
    do {
      if (attempt > 0) Utilities.sleep(700);
      attempt++;

      idsMasterData = fetchIdsMasterData(idMasterID, attempt > 1);
      if (!idsMasterData.success) {
        return errors.propagate(
          "getTemplateAndsheetIds",
          idsMasterData,
          null,
          {
            collection: idsMasterData.collection || false,
          },
        );
      }

      templateInfo = [];
      skippedTemplates = [];
      sheetIds = [idMasterID];
      versionLoading = false;
      foundMergedThemes = false;

      for (var i = 0; i < sheetTypes.length; i++) {
        var sheetType = sheetTypes[i];
        if (
          foundMergedThemes &&
          legacyThemesSheetTypes.indexOf(sheetType) !== -1
        ) {
          continue;
        }
        try {
          var templateResult = getTemplateInfo(
            idsMasterData,
            sheetType,
            copyMode,
          );

          if (templateResult && templateResult.versionLoading) {
            versionLoading = true;
          }

          if (templateResult && templateResult.success) {
            if (sheetType === "Themes, Songs & Relics") {
              foundMergedThemes = true;
            }
            if (templateResult.versionFiltered) {
              console.log(`Skipping ${sheetType} - ${templateResult.message}`);
              skippedTemplates.push({
                sheetType: sheetType,
                reason: templateResult.skipReason || "filtered",
                label: templateResult.skipLabel || "",
                blocked: templateResult.blocked === true,
                templateVersion: templateResult.templateVersion || "",
                oldVersion: templateResult.oldVersion || "",
                message: templateResult.message,
              });
              continue;
            }

            templateInfo.push({
              templateID: templateResult.templateID,
              sheetType: sheetType,
              templateVersion: templateResult.templateVersion,
              oldVersion: templateResult.oldVersion,
              oldSheetID: templateResult.oldSheetID,
            });

            if (templateResult.oldSheetID) {
              sheetIds.push(templateResult.oldSheetID);
            }
          } else {
            console.log(
              `Error getting template info for ${sheetType}: ${
                templateResult ? templateResult.message : "Unknown error"
              }`,
            );
          }
        } catch (templateError) {
          errors.report("getTemplateAndsheetIds", templateError, {
            note: `Error processing template for ${sheetType}`,
            idMasterID: idMasterID,
            copyMode: copyMode,
          }, errors.CODES.RECOVERED);
        }
      }
    } while (versionLoading && attempt < maxAttempts);

    return {
      success: true,
      sheetIds: sheetIds,
      templateInfo: templateInfo,
      skippedTemplates: skippedTemplates,
      message: `Found ${templateInfo.length} templates and ${sheetIds.length} old sheets to check`,
    };
  } catch (error) {
    var errorReport = errors.report("getTemplateAndsheetIds", error, {
      note: `Error getting template and old sheet IDs`,
      idMasterID: idMasterID,
      copyMode: copyMode,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Template id, version and old sheet for one sheet type.
 * @param {Object} idsMasterData
 * @param {string} sheetType
 * @param {string} [copyMode]
 * @returns {{success: boolean}} Plus template details. A failure envelope on error.
 */
function getTemplateInfo(idsMasterData, sheetType, copyMode) {
  try {
    var values = idsMasterData.values;
    var formulas = idsMasterData.formulas;
    var idMasterID = idsMasterData.idMasterID;
    copyMode = copyMode || "all";

    var spreadsheetInfo = labelUtils.findSheetTypeURL(
      idMasterID,
      "IDS",
      sheetType,
      values,
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
        `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
      );
      return {
        success: false,
        message: `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
      };
    }

    var oldSheetID = sheetRefs.extractSheetId(spreadsheetInfo.id);
    if (!oldSheetID) {
      console.log(`Could not extract old sheet ID from ${spreadsheetInfo.id}`);
      return {
        success: false,
        message: `Could not extract old sheet™ ID from ${spreadsheetInfo.id}`,
      };
    }

    var versionLoading =
      versionUtils.isVersionLoading(spreadsheetInfo.version.value) ||
      versionUtils.isVersionLoading(spreadsheetInfo.oldVersion.value);
    var templateVersion = versionUtils.isVersionLoading(spreadsheetInfo.version.value)
      ? ""
      : spreadsheetInfo.version.value;
    var oldVersion = versionUtils.isVersionLoading(spreadsheetInfo.oldVersion.value)
      ? ""
      : spreadsheetInfo.oldVersion.value;

    var templateStatus = versionUtils.getVersionStatus(templateVersion);
    if (templateStatus.blocked) {
      console.log(
        `${sheetType} template is ${templateStatus.label} (version cell: "${templateVersion}"), skipping`,
      );
      return {
        success: true,
        versionFiltered: true,
        skipReason: templateStatus.status,
        skipLabel: templateStatus.label,
        blocked: true,
        templateVersion: templateVersion,
        oldVersion: oldVersion,
        versionLoading: versionLoading,
        message: `${sheetType} is ${templateStatus.label}`,
      };
    }

    if (copyMode === "update") {
      if (!templateVersion || !oldVersion) {
        console.log(
          `Version information missing for ${sheetType} - template: ${templateVersion}, old: ${oldVersion}`,
        );
        return {
          success: true,
          versionFiltered: true,
          skipReason: "missingVersion",
          blocked: false,
          templateVersion: templateVersion,
          oldVersion: oldVersion,
          versionLoading: versionLoading,
          message: `Version information missing for ${sheetType}`,
        };
      }

      var versionComparison = versionUtils.compareVersions(
        oldVersion,
        templateVersion,
      );
      if (versionComparison !== "older") {
        console.log(
          `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}, skipping`,
        );
        return {
          success: true,
          versionFiltered: true,
          skipReason: "upToDate",
          blocked: false,
          templateVersion: templateVersion,
          oldVersion: oldVersion,
          versionLoading: versionLoading,
          message: `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}`,
        };
      }

      console.log(
        `${sheetType} template version ${templateVersion} is newer than old version ${oldVersion}, including`,
      );
    }

    var templateRow = spreadsheetInfo.template.row - 1;
    var templateCol = spreadsheetInfo.template.col - 1;

    var templateUrl = "";
    if (
      formulas &&
      formulas[templateRow] &&
      formulas[templateRow][templateCol]
    ) {
      templateUrl = sheetRefs.extractUrlFromHyperlink(
        formulas[templateRow][templateCol],
      );
    }

    if (!templateUrl) {
      console.log(`Template URL not found for ${sheetType}`);
      return {
        success: false,
        message: `Template URL not found for ${sheetType}`,
      };
    }

    var templateID = sheetRefs.extractSheetId(templateUrl);
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
      versionLoading: versionLoading,
      sheetType: sheetType,
      message: `Successfully got template info for ${sheetType}`,
    };
  } catch (error) {
    var errorReport = errors.report("getTemplateInfo", error, {
      note: `Error getting template info for ${sheetType}`,
      idsMasterData: idsMasterData,
      sheetType: sheetType,
      copyMode: copyMode,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Client-callable. Resolves a sheet's id and type from a link or ID.
 * @param {string} sheetID
 * @param {string} [sheetType]
 * @returns {{success: boolean, sheetID: string, sheetType: string}} A failure envelope on error.
 */
function findSheetIdAndType(sheetID, sheetType) {
  if (!sheetID) {
    console.log(`Missing sheetId parameter.`);
    return { error: "Missing sheetType parameter." };
  }
  sheetType = sheetType || "IDS Master's";
  var spreadsheetInfo = labelUtils.findSheetTypeID(sheetID, "IDS", sheetType);
  if (!spreadsheetInfo || !spreadsheetInfo.id) {
    return errors.reject(
      "findSheetIdAndType",
      errors.CODES.SHEET_STRUCTURE,
      `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
    );
  }
  console.log(`Found sheet type ID: ${spreadsheetInfo.id}`);
  var spreadsheetId = sheetRefs.extractSheetId(spreadsheetInfo.id);
  if (!spreadsheetId) {
    return errors.reject(
      "findSheetIdAndType",
      errors.CODES.SHEET_STRUCTURE,
      `Could not extract sheet™ ID from ${spreadsheetInfo.id}`,
    );
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

/**
 * Reads an IDS Master's IDS tab.
 * @param {string} idMasterID
 * @param {boolean} [forceRefresh]
 * @returns {{success: boolean, values: Array<Array<*>>}} A failure envelope on error.
 */
function fetchIdsMasterData(idMasterID, forceRefresh = false) {
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

    var idsValues = SheetsAPI.batchGetValues(
      idMasterID,
      ["IDS"],
      true,
      forceRefresh,
    );
    var idsFormulas = SheetsAPI.batchGetFormulas(
      idMasterID,
      ["IDS"],
      true,
      forceRefresh,
    );

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
    var errorReport = errors.report("fetchIdsMasterData", error, {
      note: `Error fetching IDS Master data`,
      idMasterID: idMasterID,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Resolves one sheet type's template and reports whether it is reachable.
 * @param {Object} idsMasterData
 * @param {string} sheetType
 * @param {string} [copyMode]
 * @returns {{success: boolean, accessDenied: boolean}} Plus template details.
 */
function processTemplateAccess(idsMasterData, sheetType, copyMode) {
  try {
    var values = idsMasterData.values;
    var formulas = idsMasterData.formulas;
    var idMasterID = idsMasterData.idMasterID;
    copyMode = copyMode || "all";

    var spreadsheetInfo = labelUtils.findSheetTypeURL(
      idMasterID,
      "IDS",
      sheetType,
      values,
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
        `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
      );
      return {
        success: false,
        message: `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
      };
    }

    var oldSheetID = sheetRefs.extractSheetId(spreadsheetInfo.id);
    if (!oldSheetID) {
      console.log(`Could not extract old sheet ID from ${spreadsheetInfo.id}`);
      return {
        success: false,
        message: `Could not extract old sheet™ ID from ${spreadsheetInfo.id}`,
      };
    }

    var versionLoading =
      versionUtils.isVersionLoading(spreadsheetInfo.version.value) ||
      versionUtils.isVersionLoading(spreadsheetInfo.oldVersion.value);
    var templateVersion = versionUtils.isVersionLoading(spreadsheetInfo.version.value)
      ? ""
      : spreadsheetInfo.version.value;
    var oldVersion = versionUtils.isVersionLoading(spreadsheetInfo.oldVersion.value)
      ? ""
      : spreadsheetInfo.oldVersion.value;

    var templateStatus = versionUtils.getVersionStatus(templateVersion);
    if (templateStatus.blocked) {
      console.log(
        `${sheetType} template is ${templateStatus.label} (version cell: "${templateVersion}"), blocking copy`,
      );
      return {
        success: false,
        blocked: true,
        skipReason: templateStatus.status,
        skipLabel: templateStatus.label,
        sheetType: sheetType,
        templateVersion: templateVersion,
        versionLoading: versionLoading,
        message: `${sheetType} is ${templateStatus.label}, so it cannot be copied right now. Please try again once the template has been released.`,
      };
    }

    if (copyMode === "update") {
      if (!templateVersion || !oldVersion) {
        console.log(
          `Version information missing for ${sheetType} - template: ${templateVersion}, old: ${oldVersion}`,
        );
        return {
          success: true,
          versionFiltered: true,
          versionLoading: versionLoading,
          message: `Version information missing for ${sheetType}`,
        };
      }

      var versionComparison = versionUtils.compareVersions(
        oldVersion,
        templateVersion,
      );
      if (versionComparison !== "older") {
        console.log(
          `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}, skipping`,
        );
        return {
          success: true,
          versionFiltered: true,
          versionLoading: versionLoading,
          message: `${sheetType} template version ${templateVersion} is not newer than old version ${oldVersion}`,
        };
      }

      console.log(
        `${sheetType} template version ${templateVersion} is newer than old version ${oldVersion}, including`,
      );
    }

    var templateRow = spreadsheetInfo.template.row - 1;
    var templateCol = spreadsheetInfo.template.col - 1;

    var templateUrl = "";
    if (
      formulas &&
      formulas[templateRow] &&
      formulas[templateRow][templateCol]
    ) {
      templateUrl = sheetRefs.extractUrlFromHyperlink(
        formulas[templateRow][templateCol],
      );
    }

    if (!templateUrl) {
      console.log(`Template URL not found for ${sheetType}`);
      return {
        success: false,
        message: `Template URL not found for ${sheetType}`,
      };
    }

    var templateID = sheetRefs.extractSheetId(templateUrl);
    if (!templateID) {
      console.log(`Could not extract template ID from URL: ${templateUrl}`);
      return {
        success: false,
        message: `Could not extract template ID from URL: ${templateUrl}`,
      };
    }

    try {
      var file = CacheManager.getFile(templateID);

      return {
        success: true,
        message: `Template access verified for ${sheetType}.`,
        accessDenied: false,
        templateID: templateID,
        templateVersion: templateVersion,
        oldVersion: oldVersion,
        versionLoading: versionLoading,
        oldFileId: oldSheetID,
        idMasterFileId: idMasterID,
        sheetType: sheetType,
      };
    } catch (error) {
      errors.report("processTemplateAccess", error, {
        note: `Error retrieving template file information`,
        idsMasterData: idsMasterData,
        sheetType: sheetType,
        copyMode: copyMode,
      }, errors.CODES.RECOVERED);
      console.log(`Template ID: ${templateID}, Sheet Type: ${sheetType}`);

      return {
        success: true,
        message: `Template access check completed. Access needed for ${sheetType} template.`,
        accessDenied: true,
        templateID: templateID,
        templateVersion: templateVersion,
        oldVersion: oldVersion,
        versionLoading: versionLoading,
        oldFileId: oldSheetID,
        idMasterFileId: idMasterID,
        sheetType: sheetType,
      };
    }
  } catch (error) {
    var errorReport = errors.report("processTemplateAccess", error, {
      note: `Error processing template access for ${sheetType}`,
      idsMasterData: idsMasterData,
      sheetType: sheetType,
      copyMode: copyMode,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Client-callable. The template id for one sheet, from its own IDS tab.
 * @param {string} sheetID
 * @param {string} sheetType
 * @returns {{success: boolean, templateID: string}} A failure envelope on error.
 */
function getTemplateIdForSingleSheet(sheetID, sheetType) {
  try {
    if (!sheetID) {
      return errors.reject(
        "getTemplateIdForSingleSheet",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing sheetID parameter." },
      );
    }
    if (!sheetType) {
      return errors.reject(
        "getTemplateIdForSingleSheet",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing sheetType parameter." },
      );
    }
    var spreadsheetInfo = labelUtils.findSheetTemplateID(
      sheetID,
      "Home Page",
      sheetType,
    );
    if (
      !spreadsheetInfo ||
      !spreadsheetInfo.templateID ||
      !spreadsheetInfo.templateVersion
    ) {
      return errors.reject(
        "getTemplateIdForSingleSheet",
        errors.CODES.SHEET_STRUCTURE,
        `Could not find sheet template for ${sheetType}`,
      );
    }
    return {
      success: true,
      templateID: spreadsheetInfo.templateID,
      templateVersion: spreadsheetInfo.templateVersion,
      message: `Successfully got template ID for ${sheetType}`,
    };
  } catch (error) {
    var errorReport = errors.report("getTemplateIdForSingleSheet", error, {
      note: `Error getting template ID for single sheet`,
      sheetID: sheetID,
      sheetType: sheetType,
    });
    return errors.fail(errorReport);
  }
}
