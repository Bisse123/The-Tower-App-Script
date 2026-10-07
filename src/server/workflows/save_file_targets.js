/**
 * Client-callable. The sheet each save-file category imports into.
 * @param {string} idMasterID
 * @param {string[]} [sheetTypes]
 * @returns {{success: boolean, targets: Object}} A failure envelope on error.
 */
function getSaveFileImportTargets(idMasterID, sheetTypes) {
  try {
    var resolvedIdMasterID = idMasterID
      ? sheetRefs.extractSheetId(String(idMasterID))
      : null;
    if (!resolvedIdMasterID) {
      return errors.reject(
        "getSaveFileImportTargets",
        errors.CODES.INVALID_LINK,
        null,
        {
          idMasterID: "",
          targets: {},
          missing: [],
          versions: {},
        },
      );
    }

    var requestedTypes =
      Array.isArray(sheetTypes) && sheetTypes.length > 0
        ? sheetTypes
        : [
            "Laboratory",
            "Workshop",
            "Ultimate Weapon",
            "Themes, Songs & Relics",
            "Bots",
            "Vault",
            "Cards",
            "Modules",
            "Guardians",
            "Player & Stuff",
            "IDS Master",
          ];

    /**
     * Reads a version out of a cell object.
     * @param {{value: *}} cell
     * @returns {string}
     */
    function readSheetVersion(cell) {
      return versionUtils.readVersion(cell && cell.value);
    }

    var targets, missing, versions, idsMasterData, incomplete;
    var attempt = 0;
    var maxAttempts = 2;
    do {
      if (attempt > 0) Utilities.sleep(700);
      attempt++;
      var forceRefresh = attempt > 1;

      idsMasterData = fetchIdsMasterData(resolvedIdMasterID, forceRefresh);
      if (!idsMasterData.success) {
        if (attempt >= maxAttempts) {
          return errors.reject(
            "getSaveFileImportTargets",
            errors.CODES.SHEET_STRUCTURE,
            idsMasterData.message || "Could not read IDS Master.",
            {
              idMasterID: resolvedIdMasterID,
              targets: {},
              missing: [],
              versions: {},
            },
          );
        }
        continue;
      }

      var values = idsMasterData.values;

      targets = {};
      missing = [];
      versions = {};
      incomplete = false;

      for (var i = 0; i < requestedTypes.length; i++) {
        var sheetType = requestedTypes[i];

        if (sheetType === "IDS Master") {
          targets[sheetType] = resolvedIdMasterID;
          var masterVersion = compareSheetVersions(
            resolvedIdMasterID,
            sheetType,
            forceRefresh,
          );
          var masterCurrent = String((masterVersion && masterVersion.currentVersion) || "").trim();
          var masterLatest = String((masterVersion && masterVersion.latestVersion) || "").trim();
          if (masterCurrent.toLowerCase().indexOf("loading") === 0) masterCurrent = "";
          if (masterLatest.toLowerCase().indexOf("loading") === 0) masterLatest = "";
          versions[sheetType] = {
            currentVersion: masterCurrent,
            latestVersion: masterLatest,
            upToDate: !!(
              masterVersion &&
              masterVersion.success &&
              masterVersion.comparisonResult !== "older"
            ),
          };
          if (!masterCurrent || !masterLatest) incomplete = true;
          continue;
        }

        var sheetTypeInfo = labelUtils.findSheetTypeURL(
          resolvedIdMasterID,
          "IDS",
          sheetType,
          values,
        );

        var targetID = sheetTypeInfo && sheetTypeInfo.id
          ? sheetRefs.extractSheetId(sheetTypeInfo.id)
          : null;

        if (!targetID) {
          missing.push(sheetType);
          continue;
        }

        targets[sheetType] = targetID;

        var latestVersion = readSheetVersion(sheetTypeInfo.version);
        var currentVersion = readSheetVersion(sheetTypeInfo.oldVersion);

        if (!latestVersion || !currentVersion) incomplete = true;

        var upToDate =
          latestVersion && currentVersion
            ? versionUtils.compareVersions(currentVersion, latestVersion) !== "older"
            : false;

        versions[sheetType] = {
          currentVersion: currentVersion,
          latestVersion: latestVersion,
          upToDate: upToDate,
        };
      }
    } while (incomplete && attempt < maxAttempts);

    return {
      success: true,
      message: `Resolved ${Object.keys(targets).length} target sheet(s) from IDS Master.`,
      idMasterID: resolvedIdMasterID,
      targets: targets,
      missing: missing,
      versions: versions,
    };
  } catch (error) {
    var errorReport = errors.report("getSaveFileImportTargets", error, {
      note: `Error resolving save-file import targets`,
      idMasterID: idMasterID,
      sheetTypes: sheetTypes,
    });
    return errors.fail(errorReport, null, {
      idMasterID: "",
      targets: {},
      missing: [],
      versions: {},
    });
  }
}

/**
 * Client-callable. Whether a linked file is an IDS Master or Collection,
 * and whether it is out of date.
 * @param {string} sheetID
 * @returns {{success: boolean, sheetType: string, outdated?: boolean}} A failure envelope on error.
 */
function getSaveFileSheetType(sheetID) {
  try {
    var resolvedID = sheetID
      ? sheetRefs.extractSheetId(String(sheetID)) || ""
      : "";
    if (!resolvedID) {
      return errors.reject(
        "getSaveFileSheetType",
        errors.CODES.INVALID_INPUT,
        "No sheet ID provided.",
        {
          sheetType: "",
        },
      );
    }

    var batchResult = SheetsAPI.batchGetValues(resolvedID, ["Home Page"]);
    var homePageValues =
      batchResult && batchResult[0] && batchResult[0].values
        ? batchResult[0].values
        : null;

    if (!homePageValues) {
      return errors.reject(
        "getSaveFileSheetType",
        errors.CODES.SHEET_STRUCTURE,
        "Could not read the Home Page tab of that file. The script may not have access to it yet.",
        {
          sheetType: "",
          idMasterID: resolvedID,
        },
      );
    }

    var sheetType =
      homePageValues[1] && homePageValues[1][1] != null
        ? String(homePageValues[1][1]).trim()
        : "";

    if (sheetType.indexOf("IDS Collection") !== -1) {
      sheetType = "IDS Collection";
    }

    if (sheetType !== "IDS Master" && sheetType !== "IDS Collection") {
      return errors.reject(
        "getSaveFileSheetType",
        errors.CODES.INTERNAL,
        sheetType
          ? `That file is not an IDS Master or an IDS Collection (its Home Page says ${sheetType}).`
          : "Could not tell whether that file is an IDS Master or an IDS Collection.",
        {
          sheetType: sheetType,
          idMasterID: resolvedID,
        },
      );
    }

    var result = {
      success: true,
      sheetType: sheetType,
      idMasterID: resolvedID,
    };

    if (sheetType === "IDS Collection") {
      var versionInfo = versionUtils.findSheetVersion(
        resolvedID,
        "Home Page",
        "IDS Collection",
        homePageValues,
      );
      if (
        versionInfo &&
        versionInfo.currentVersion &&
        versionInfo.latestVersion
      ) {
        result.currentVersion = versionInfo.currentVersion;
        result.latestVersion = versionInfo.latestVersion;
        result.outdated =
          versionUtils.compareVersions(
            versionInfo.currentVersion,
            versionInfo.latestVersion,
          ) === "older";
      }
    }

    return result;
  } catch (error) {
    var errorReport = errors.report("getSaveFileSheetType", error, {
      sheetID: sheetID,
    });
    return errors.fail(errorReport, null, {
      sheetType: "",
    });
  }
}
