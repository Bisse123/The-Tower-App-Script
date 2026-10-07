const versionUtils = {
  /**
   * Reads a sheet's current and latest version from its Home Page.
   * @param {string} sheetID
   * @param {string} sheetName
   * @param {string} sheetType
   * @param {Array<Array<*>>} [preLoadedValues]
   * @returns {{currentVersion: string, latestVersion: string}|null}
   */
  findSheetVersion: function (sheetID, sheetName, sheetType, preLoadedValues) {
    try {
      if (sheetType === "Effective Paths") {
        return versionUtils.getEPathsVersion(sheetID, sheetName, preLoadedValues);
      }
      var values;
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
            `No data found in sheet: ${sheetName} in spreadsheet: ${sheetID}`,
          );
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
              cell.toLowerCase().includes(w),
            ),
        );
        var latestVersionCol = values[row].findIndex(
          (cell) =>
            typeof cell === "string" &&
            ["latest remote version", "latest version"].some((w) =>
              cell.toLowerCase().includes(w),
            ),
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
      errors.report("versionUtils.findSheetVersion", error, {
        note: `Error finding sheet version`,
        sheetID: sheetID,
        sheetName: sheetName,
        sheetType: sheetType,
        preLoadedValues: preLoadedValues,
      });
      return null;
    }
  },

  /**
   * Reads an Effective Paths sheet's version.
   * @param {string} sheetID
   * @param {string} sheetName
   * @param {Array<Array<*>>} [preLoadedValues]
   * @returns {{currentVersion: string, latestVersion: string}|null}
   */
  getEPathsVersion: function (sheetID, sheetName, preLoadedValues) {
    try {
      var values;

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
            `No data found in sheet: ${sheetName} in spreadsheet: ${sheetID}`,
          );
          return null;
        }
        values = batchResult[0].values;
      }

      var currentVersion = null;
      var latestVersion = null;

      for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
          var cellValue = values[i] && values[i][j] ? values[i][j] : "";

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
      errors.report("versionUtils.getEPathsVersion", error, {
        note: `Error finding Effective Paths version`,
        sheetID: sheetID,
        sheetName: sheetName,
        preLoadedValues: preLoadedValues,
      });
      return null;
    }
  },

  /**
   * Whether a version cell is still calculating.
   * @param {*} value
   * @returns {boolean}
   */
  isVersionLoading: function (value) {
    return (
      String(value == null ? "" : value)
        .trim()
        .toLowerCase()
        .indexOf("loading") === 0
    );
  },

  /**
   * Normalises a version cell to a version string.
   * @param {*} value
   * @returns {string}
   */
  readVersion: function (value) {
    if (value == null) return "";
    var text = String(value).trim();
    return versionUtils.isVersionLoading(text) ? "" : text;
  },

  /**
   * Classifies a version string as loading, missing or present.
   * @param {*} version
   * @returns {string}
   */
  getVersionStatus: function (version) {
    var text = String(version == null ? "" : version).trim();
    if (!text) {
      return {
        status: "missing",
        label: "missing a version number",
        blocked: false,
        version: "",
      };
    }
    if (/maintenance/i.test(text)) {
      return {
        status: "maintenance",
        label: "under maintenance",
        blocked: true,
        version: text,
      };
    }
    if (/\bWIP\b|work[\s-]*in[\s-]*progress/i.test(text)) {
      return {
        status: "wip",
        label: "still in development (WIP)",
        blocked: true,
        version: text,
      };
    }
    return { status: "ok", label: "", blocked: false, version: text };
  },

  /**
   * Compares two version strings numerically, part by part.
   * @param {string} oldVersion
   * @param {string} newVersion
   * @returns {"older"|"same"|"newer"}
   */
  compareVersions: function (oldVersion, newVersion) {

    /**
     * Splits a version string into its numeric parts.
     * @param {*} v
     * @returns {number[]} Empty when there is no version in the string.
     */
    function parseVersion(v) {
      var match = String(v || "").match(/\d+(?:\.\d+)*/);
      if (!match) {
        return [];
      }
      return match[0].split(".").map(Number);
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
};
