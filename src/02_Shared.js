const CacheManager = {
  _userCache: null,

  get userCache() {
    if (!this._userCache) {
      try {
        this._userCache = CacheService.getUserCache();
      } catch (error) {
        errors.report("CacheManager.userCache", error, null, errors.CODES.RECOVERED);
        return null;
      }
    }
    return this._userCache;
  },

  CHUNK_SIZE: 90000,

  /**
   * Byte length of a string in UTF-8.
   * @param {string} str
   * @returns {number}
   */
  _byteLength: function (str) {
    let bytes = 0;

    for (let i = 0; i < str.length; i++) {
      const code = str.charCodeAt(i);

      if (code < 0x80) {
        bytes += 1;
      } else if (code < 0x800) {
        bytes += 2;
      } else if (code >= 0xd800 && code <= 0xdbff && i + 1 < str.length) {

        bytes += 4;
        i++;
      } else {
        bytes += 3;
      }
    }

    return bytes;
  },

  /**
   * Splits a string into chunks no larger than maxBytes.
   * @param {string} str
   * @param {number} [maxBytes]
   * @returns {string[]}
   */
  _chunkString: function (str, maxBytes = this.CHUNK_SIZE) {
    const chunks = [];
    let start = 0;
    let bytes = 0;
    let i = 0;

    while (i < str.length) {
      const code = str.charCodeAt(i);
      let charBytes = 3;
      let step = 1;

      if (code < 0x80) {
        charBytes = 1;
      } else if (code < 0x800) {
        charBytes = 2;
      } else if (code >= 0xd800 && code <= 0xdbff && i + 1 < str.length) {
        charBytes = 4;
        step = 2;
      }

      if (bytes + charBytes > maxBytes && i > start) {
        chunks.push(str.substring(start, i));
        start = i;
        bytes = 0;
      }

      bytes += charBytes;
      i += step;
    }

    chunks.push(str.substring(start));

    return chunks;
  },

  /**
   * How many chunks a cached key was split into.
   * @param {string} key
   * @returns {number} 0 when the key is not chunked.
   */
  _chunkCount: function (key) {
    const chunksCountStr = this.userCache.get(`${key}__chunks`);
    if (!chunksCountStr) {
      return 0;
    }

    const chunkCount = parseInt(chunksCountStr, 10);
    return isNaN(chunkCount) || chunkCount < 0 ? 0 : chunkCount;
  },

  /**
   * Every cache key a value occupies, chunks included.
   * @param {string} key
   * @returns {string[]}
   */
  _entryKeys: function (key) {
    const keys = [key];
    const chunkCount = this._chunkCount(key);

    for (let i = 0; i < chunkCount; i++) {
      keys.push(`${key}__chunk_${i}`);
    }
    if (chunkCount > 0) {
      keys.push(`${key}__chunks`);
    }

    return keys;
  },

  /**
   * Reads a cached value, rejoining its chunks.
   * @param {string} key
   * @returns {string|null}
   */
  _retrieveValue: function (key) {
    if (!this.userCache) {
      console.log(`Cache unavailable - cannot retrieve: ${key}`);
      return null;
    }

    const chunkCount = this._chunkCount(key);

    if (chunkCount === 0) {
      return this.userCache.get(key);
    }

    let combinedValue = "";

    for (let i = 0; i < chunkCount; i++) {
      const chunk = this.userCache.get(`${key}__chunk_${i}`);

      if (chunk === null || chunk === undefined) {
        console.log(
          `Chunk ${i + 1} of ${chunkCount} missing for ${key} - treating as a cache miss`,
        );
        return null;
      }

      combinedValue += chunk;
    }

    return combinedValue;
  },

  /**
   * Caches a value, chunking it when it is too large.
   * @param {string} key
   * @param {string} value
   * @returns {void}
   */
  _putValue: function (key, value) {
    this._putAllValues({ [key]: value });
  },

  /**
   * Caches several values in one call.
   * @param {Object} cacheData Key to value.
   * @returns {void}
   */
  _putAllValues: function (cacheData) {
    if (!this.userCache) {
      console.log(
        `Cache unavailable - cannot store: ${Object.keys(cacheData).join(", ")}`,
      );
      return;
    }

    const toStore = {};
    const staleKeys = [];

    for (const key in cacheData) {
      const value = cacheData[key];
      const previousChunkCount = this._chunkCount(key);
      const byteLength = this._byteLength(value);

      if (byteLength > this.CHUNK_SIZE) {
        const chunks = this._chunkString(value);

        for (let i = 0; i < chunks.length; i++) {
          toStore[`${key}__chunk_${i}`] = chunks[i];
        }
        toStore[`${key}__chunks`] = chunks.length.toString();

        staleKeys.push(key);
        for (let i = chunks.length; i < previousChunkCount; i++) {
          staleKeys.push(`${key}__chunk_${i}`);
        }

        console.log(
          `Chunked ${key} into ${chunks.length} parts (${byteLength} bytes total)`,
        );
      } else {
        toStore[key] = value;

        if (previousChunkCount > 0) {
          staleKeys.push(`${key}__chunks`);
          for (let i = 0; i < previousChunkCount; i++) {
            staleKeys.push(`${key}__chunk_${i}`);
          }
        }
      }
    }

    if (staleKeys.length > 0) {
      this.userCache.removeAll(staleKeys);
    }
    const cacheTimeMinutes = 1;
    this.userCache.putAll(toStore, cacheTimeMinutes * 60);
  },

  /**
   * Spreadsheet metadata, from cache or the Sheets API.
   * @param {string} spreadsheetTypeName Cache label, e.g. "Cards oldSpreadsheet".
   * @param {string} sheetID
   * @returns {Object|null}
   */
  getSpreadsheet: function (spreadsheetTypeName, sheetID) {
    if (!spreadsheetTypeName) {
      console.log("No spreadsheet type name provided");
      return null;
    }

    const cached = this._retrieveValue(spreadsheetTypeName);

    if (cached) {
      const cachedData = JSON.parse(cached);

      if (!sheetID) {
        console.log(
          `Using cached ${spreadsheetTypeName} (sheetID: ${cachedData.sheetID})`,
        );
        return cachedData.metadata;
      }

      if (cachedData.sheetID === sheetID) {
        console.log(
          `Cache hit for ${spreadsheetTypeName} (sheetID: ${sheetID})`,
        );
        return cachedData.metadata;
      }

      console.log(
        `Cache invalidated for ${spreadsheetTypeName}: sheetID changed from ${cachedData.sheetID} to ${sheetID}`,
      );
    }

    if (!sheetID) {
      console.log(
        `No cached entry and no sheetID provided for ${spreadsheetTypeName}`,
      );
      return null;
    }

    console.log(`Fetching fresh ${spreadsheetTypeName} (sheetID: ${sheetID})`);
    const metadata = SheetsAPI.fetchSpreadsheet(sheetID);

    if (metadata) {
      const cacheData = {
        metadata: metadata,
        sheetID: sheetID,
      };
      this._putValue(spreadsheetTypeName, JSON.stringify(cacheData));
      console.log(`Cached ${spreadsheetTypeName} (sheetID: ${sheetID})`);
      return metadata;
    }

    return null;
  },

  /**
   * Cached values for ranges.
   * @param {string} spreadsheetId
   * @param {string[]} ranges
   * @param {boolean} [forceRefresh]
   * @returns {Array<Object>|null} valueRanges.
   */
  getSheetValues: function (spreadsheetId, ranges, forceRefresh = false) {
    const cachedData = [];
    const uncachedRanges = [];
    const uncachedIndices = [];

    const cacheKeys = ranges.map((range) => `${spreadsheetId}|${range}|VALUE`);

    for (let i = 0; i < cacheKeys.length; i++) {

      const cached = forceRefresh ? null : this._retrieveValue(cacheKeys[i]);

      if (cached) {
        cachedData[i] = JSON.parse(cached);
        console.log(`Cache hit for values: ${ranges[i]}`);
      } else {
        uncachedRanges.push(ranges[i]);
        uncachedIndices.push(i);
      }
    }

    let result = [...cachedData];

    if (uncachedRanges.length > 0) {
      console.log(`Fetching uncached ranges: ${uncachedRanges.join(", ")}`);
      const fetchedData = SheetsAPI.batchGetValues(
        spreadsheetId,
        uncachedRanges,
        false,
      );

      if (fetchedData) {
        const cacheData = {};
        for (let i = 0; i < fetchedData.length; i++) {
          const cacheKey = `${spreadsheetId}|${uncachedRanges[i]}|VALUE`;
          cacheData[cacheKey] = JSON.stringify(fetchedData[i]);
          result[uncachedIndices[i]] = fetchedData[i];
        }
        this._putAllValues(cacheData);
      }
    }

    return result;
  },

  /**
   * Cached formulas for ranges.
   * @param {string} spreadsheetId
   * @param {string[]} ranges
   * @param {boolean} [forceRefresh]
   * @returns {Array<Object>|null} valueRanges.
   */
  getSheetFormulas: function (spreadsheetId, ranges, forceRefresh = false) {
    const cachedData = [];
    const uncachedRanges = [];
    const uncachedIndices = [];

    const cacheKeys = ranges.map(
      (range) => `${spreadsheetId}|${range}|FORMULA`,
    );

    for (let i = 0; i < cacheKeys.length; i++) {
      const cached = forceRefresh ? null : this._retrieveValue(cacheKeys[i]);

      if (cached) {
        cachedData[i] = JSON.parse(cached);
        console.log(`Cache hit for formulas: ${ranges[i]}`);
      } else {
        uncachedRanges.push(ranges[i]);
        uncachedIndices.push(i);
      }
    }

    let result = [...cachedData];

    if (uncachedRanges.length > 0) {
      console.log(`Fetching uncached formulas: ${uncachedRanges.join(", ")}`);
      const fetchedData = SheetsAPI.batchGetFormulas(
        spreadsheetId,
        uncachedRanges,
        false,
      );

      if (fetchedData) {
        const cacheData = {};
        for (let i = 0; i < fetchedData.length; i++) {
          const cacheKey = `${spreadsheetId}|${uncachedRanges[i]}|FORMULA`;
          cacheData[cacheKey] = JSON.stringify(fetchedData[i]);
          result[uncachedIndices[i]] = fetchedData[i];
        }
        this._putAllValues(cacheData);
      }
    }

    return result;
  },

  /**
   * Invalidates a spreadsheet's cache entries.
   * @param {string} spreadsheetTypeName
   * @returns {void}
   */
  RemoveSpreadsheet: function (spreadsheetTypeName) {
    if (!this.userCache) {
      console.log(`Cache unavailable - cannot remove: ${spreadsheetTypeName}`);
      return;
    }

    try {

      const cached = this._retrieveValue(spreadsheetTypeName);
      if (!cached) {
        console.log(
          `No cache entry found for ${spreadsheetTypeName} to invalidate`,
        );
        return;
      }

      const cachedData = JSON.parse(cached);
      const sheetID = cachedData.sheetID;
      const metadata = cachedData.metadata;

      let keysToRemove = this._entryKeys(spreadsheetTypeName);

      if (metadata && metadata.sheets) {
        for (let i = 0; i < metadata.sheets.length; i++) {
          const sheetName = metadata.sheets[i].properties.title;

          keysToRemove = keysToRemove
            .concat(this._entryKeys(`${sheetID}|${sheetName}|VALUE`))
            .concat(this._entryKeys(`${sheetID}|${sheetName}|FORMULA`));
        }
      }

      this.userCache.removeAll(keysToRemove);
      console.log(
        `Invalidated cache for ${spreadsheetTypeName} and ${keysToRemove.length - 1} sheet entries`,
      );
    } catch (error) {
      errors.report("cacheData.RemoveSpreadsheet", error, {
        note: `Error invalidating cache`,
        spreadsheetTypeName: spreadsheetTypeName,
      }, errors.CODES.RECOVERED);
    }
  },

  /**
   * Drive file metadata, from cache or Drive.
   * @param {string} fileID
   * @returns {Object|null} Null when it cannot be read.
   */
  getFile: function (fileID) {
    if (!fileID) {
      console.log("No file ID provided");
      return null;
    }

    const cacheKey = `File|${fileID}`;
    const cached = this._retrieveValue(cacheKey);

    if (cached) {
      console.log(`Cache hit for file: ${fileID}`);
      return JSON.parse(cached);
    }

    const allFieldsNeeded =
      "id, name, parents, owners/me, capabilities/canEdit, trashed";
    try {
      const file = Drive.Files.get(fileID, { fields: allFieldsNeeded });

      if (file) {
        this._putValue(cacheKey, JSON.stringify(file));
        console.log(`Cached file metadata: ${fileID}`);
        return file;
      }
    } catch (error) {
      errors.report("cacheData.getFile", error, {
        note: `Error fetching file`,
        fileID: fileID,
      }, errors.CODES.RECOVERED);
    }

    return null;
  },

  /**
   * Invalidates a file's cache entries.
   * @param {string} fileID
   * @returns {void}
   */
  RemoveFile: function (fileID) {
    if (!fileID) {
      console.log("No file ID provided");
      return;
    }

    if (!this.userCache) {
      console.log(`Cache unavailable - cannot remove file: ${fileID}`);
      return;
    }

    const keysToRemove = this._entryKeys(`File|${fileID}`);

    if (keysToRemove.length > 1) {
      console.log(
        `Removing ${keysToRemove.length} cache keys for file: ${fileID}`,
      );
    }

    this.userCache.removeAll(keysToRemove);
    console.log(`Invalidated cache for file: ${fileID}`);
  },
};

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

const shared = {
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
        return shared.getEPathsVersion(sheetID, sheetName, preLoadedValues);
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
      errors.report("shared.findSheetVersion", error, {
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
      errors.report("shared.getEPathsVersion", error, {
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
    return shared.isVersionLoading(text) ? "" : text;
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

  /**
   * Whether a cell is the ID label for a sheet type.
   * @param {*} cell
   * @param {string} sheetType
   * @returns {boolean}
   */
  isSheetTypeCell: function (cell, sheetType) {
    if (typeof cell !== "string" || !sheetType) {
      return false;
    }
    return (
      new RegExp(sheetType, "i").test(cell) &&
      /\bID\b/i.test(cell) &&
      cell.indexOf("script") === -1 &&
      cell.indexOf("More IDs are available") === -1
    );
  },

  /**
   * Finds a sheet type's ID in an IDS tab.
   * @param {string} spreadsheetId
   * @param {string} sheetName
   * @param {string} sheetType
   * @param {Array<Array<*>>} [values]
   * @returns {string} "" when not found.
   */
  findSheetTypeID: function (
    spreadsheetId,
    sheetName,
    sheetType,
    preLoadedValues,
  ) {
    var sheetType = sheetType || "IDS Master's";
    var values;

    if (preLoadedValues) {
      values = preLoadedValues;
    } else {
      var batchResult = SheetsAPI.batchGetValues(spreadsheetId, [sheetName]);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log(
          `No data found in sheet: ${sheetName} in spreadsheet: ${spreadsheetId}`,
        );
        return null;
      }
      values = batchResult[0].values;
    }

    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (shared.isSheetTypeCell(values[i][j], sheetType)) {
          var cellA1 = shared.columnToLetter(j + 2) + (i + 1);
          var accessA1 = shared.columnToLetter(j + 4) + (i + 1);
          var importedA1 = shared.columnToLetter(j + 4) + (i + 2);

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
          };
        }
      }
    }
    return null;
  },

  /**
   * Finds a sheet type's row in an IDS tab: id, template and version.
   * @param {string} spreadsheetId
   * @param {string} sheetName
   * @param {string} sheetType
   * @param {Array<Array<*>>} [values]
   * @returns {Object|null}
   */
  findSheetTypeURL: function (
    spreadsheetId,
    sheetName,
    sheetType,
    preLoadedValues,
  ) {
    var sheetType = sheetType || "IDS Master's";
    var values;

    if (preLoadedValues) {
      values = preLoadedValues;
    } else {
      var batchResult = SheetsAPI.batchGetValues(spreadsheetId, [sheetName]);
      if (!batchResult || batchResult.length === 0 || !batchResult[0].values) {
        console.log(
          `No data found in sheet: ${sheetName} in spreadsheet: ${spreadsheetId}`,
        );
        return null;
      }
      values = batchResult[0].values;
    }

    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (shared.isSheetTypeCell(values[i][j], sheetType)) {
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

  /**
   * Pulls a spreadsheet ID out of a URL or a bare ID.
   * @param {*} input
   * @returns {string} "" when it is not a sheet link or ID.
   */
  extractSheetId: function (input) {
    input = input.trim();
    var idPattern = /^[a-zA-Z0-9_-]{20,}$/;
    var urlPattern =
      /\/spreadsheets\/(?:u\/\d+\/)?d\/([a-zA-Z0-9_-]{20,})(?:[\/?#]|$)/;

    if (idPattern.test(input)) {
      return input;
    }
    var match = input.match(urlPattern);
    if (match && match[1]) {
      return match[1];
    }
    return null;
  },

  /**
   * 1-indexed column number to its A1 letters.
   * @param {number} column
   * @returns {string}
   */
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

  /**
   * Pulls the URL out of a HYPERLINK formula.
   * @param {*} formula
   * @returns {string}
   */
  extractUrlFromHyperlink: function (formula) {
    if (!formula || typeof formula !== "string") {
      return null;
    }

    if (!formula.startsWith("=")) {
      return null;
    }

    var hyperlinkMatch = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
    if (hyperlinkMatch && hyperlinkMatch[1]) {
      return hyperlinkMatch[1];
    }

    return null;
  },

  /**
   * Resolves a data-validation value against its named range.
   * @param {*} oldValue
   * @param {Object} dvtNamedRangesData
   * @returns {*}
   */
  getDVTValue: function (oldValue, dvtNamedRangesData) {
    if (!oldValue || !dvtNamedRangesData) {
      return oldValue;
    }

    var oldLevel = String(oldValue).split("|")[0].trim();

    for (var i = 0; i < dvtNamedRangesData.length; i++) {
      var row = dvtNamedRangesData[i];
      var val = row[0] ? row[0].split("|")[0].trim() : null;
      if (val && val === oldLevel) {
        return row[0];
      }
    }
    return oldValue;
  },

  templatePresetNames: ["Farming", "Tourney"],

  /**
   * Orders preset names, honouring any forced ordering.
   * @param {string[]} presetNames
   * @param {string[]} [forcedNames]
   * @returns {string[]}
   */
  resolvePresetOrder: function (presetNames, forcedNames) {
    var names = (presetNames || []).slice();
    var slotCount = names.length;
    var indices = new Array(slotCount).fill(null);
    var assignedSourceIndices = {};

    (forcedNames || []).forEach(function (forcedName, slot) {
      if (slot >= slotCount) {
        return;
      }
      var sourceIndex = names.findIndex(function (name, idx) {
        return name === forcedName && !assignedSourceIndices.hasOwnProperty(idx);
      });
      if (sourceIndex !== -1) {
        indices[slot] = sourceIndex;
        assignedSourceIndices[sourceIndex] = true;
      }
    });

    var remainingSourceIndices = names
      .map(function (_, idx) {
        return idx;
      })
      .filter(function (idx) {
        return !assignedSourceIndices.hasOwnProperty(idx);
      });

    var remainingCursor = 0;
    for (var slot = 0; slot < slotCount; slot++) {
      if (indices[slot] === null) {
        indices[slot] = remainingSourceIndices[remainingCursor++];
      }
    }

    var order = indices.map(function (sourceIndex, slot) {
      return names[sourceIndex] || `Preset ${slot + 1}`;
    });

    return { order: order, indices: indices };
  },

  /**
   * Finds a sheet type's template ID in an IDS tab.
   * @param {string} sheetID
   * @param {string} sheetName
   * @param {string} sheetType
   * @returns {string|null}
   */
  findSheetTemplateID: function (sheetID, sheetName, sheetType) {
    try {
      console.log(
        `Finding template ID for sheet: ${sheetID}, sheet name: ${sheetName}, type: ${sheetType}`,
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
        `Searching ${formulaData.length} rows for template HYPERLINK formulas`,
      );

      var templateID = null;
      var version = null;

      for (var i = 0; i < formulaData.length; i++) {
        for (var j = 0; j < formulaData[i].length; j++) {
          var formula = formulaData[i][j];

          if (
            formula &&
            typeof formula === "string" &&
            formula.toUpperCase().includes("HYPERLINK") &&
            formula.toLowerCase().includes("copy")
          ) {
            console.log(
              `Found potential template link in row ${i + 1}, col ${
                j + 1
              }: ${formula}`,
            );

            var templateUrl = shared.extractUrlFromHyperlink(formula);
            if (templateUrl) {
              templateID = shared.extractSheetId(templateUrl);
              if (templateID) {
                console.log(`Found template ID: ${templateID}`);
              }
            }
          }
        }
        if (templateID) {
          break;
        }
      }

      if (templateID) {
        var currentSheetVersionInfo = shared.findSheetVersion(
          sheetID,
          sheetName,
          sheetType,
          valueData,
        );
        if (currentSheetVersionInfo && currentSheetVersionInfo.latestVersion) {
          console.log(
            `Template version (from latest): ${currentSheetVersionInfo.latestVersion}`,
          );
          return {
            templateID: templateID,
            templateVersion: currentSheetVersionInfo.latestVersion,
          };
        }
      }

      console.log(
        `No template HYPERLINK with "copy" found in sheet: ${sheetName}`,
      );
      return null;
    } catch (error) {
      errors.report("shared.findSheetTemplateID", error, {
        note: `Error finding template ID`,
        sheetID: sheetID,
        sheetName: sheetName,
        sheetType: sheetType,
      });
      return null;
    }
  },

  /**
   * Zero-based column offset of an A1 range's first column.
   * @param {string} range
   * @returns {number}
   */
  getColumnOffsetFromRange: function (range) {
    var rangePart = range.split("!")[1];
    if (!rangePart) return 0;

    var startCell = rangePart.split(":")[0];
    if (!startCell) return 0;

    var columnLetters = startCell.replace(/[0-9]/g, "");

    var columnIndex = 0;
    for (var i = 0; i < columnLetters.length; i++) {
      columnIndex =
        columnIndex * 26 +
        (columnLetters.charCodeAt(i) - "A".charCodeAt(0) + 1);
    }

    return columnIndex - 1;
  },

  /**
   * Appends the IDS Master ID writes to a batch update.
   * @param {Array<Object>} batchUpdate Mutated in place.
   * @param {string} sheetType
   * @param {string} newSheetID
   * @param {Array<Array<*>>} idsData
   * @param {string} idMasterID
   * @returns {Array<Object>} The same batch.
   */
  addIDUpdatesToBatch: function (
    batchUpdate,
    sheetType,
    newSheetID,
    idsData,
    idMasterID,
  ) {
    try {
      if (newSheetID && idMasterID) {
        var thisSheetInfo = shared.findSheetTypeID(
          newSheetID,
          "IDS",
          "This Sheet ID",
          idsData,
        );
        var idMasterInfo = shared.findSheetTypeID(
          newSheetID,
          "IDS",
          "IDS Master's",
          idsData,
        );

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
      errors.report("shared.addIDUpdatesToBatch", error, {
        note: `Error adding ID updates to batch`,
        batchUpdate: batchUpdate,
        sheetType: sheetType,
        newSheetID: newSheetID,
        idsData: idsData,
        idMasterID: idMasterID,
      }, errors.CODES.RECOVERED);
      return batchUpdate;
    }
  },
};

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
    newVersionInfo = shared.findSheetVersion(
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

    var newVersionInfo = shared.findSheetVersion(
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
 * Writes sheet IDs into the IDS Master's IDS tab.
 * @param {string} idMasterID
 * @param {Array<{sheetType: string, sheetID: string}>} idDataEntries
 * @returns {{success: boolean, message: string}} A failure envelope on error.
 */
function updateIdsMaster(idMasterID, idDataEntries) {
  var idsMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
  if (!idsMasterSpreadsheet) {
    return errors.reject(
      "updateIdsMaster",
      errors.CODES.NOT_FOUND,
      "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
      null,
      { note: `IDS Master Spreadsheet™ not found with ID: ${idMasterID}` },
    );
  }

  var idMasterIDSheet = SheetsAPI.getSheetByName(idsMasterSpreadsheet, "IDS");
  if (!idMasterIDSheet) {
    return errors.reject(
      "updateIdsMaster",
      errors.CODES.SHEET_STRUCTURE,
      `IDS sheet™ not found in ID master spreadsheet™`,
    );
  }
  try {
    var batchUpdate = [];
    var idsMasterData = SheetsAPI.batchGetValues(idMasterID, ["IDS"]);
    var idsMasterValues = idsMasterData[0].values;
    var thisSheetInfo = shared.findSheetTypeID(
      idMasterID,
      "IDS",
      "This Sheet ID",
      idsMasterValues,
    );
    if (thisSheetInfo && thisSheetInfo.cell && thisSheetInfo.cell.range) {
      batchUpdate.push({
        range: thisSheetInfo.cell.range,
        values: [[idMasterID]],
      });
    }

    idDataEntries.forEach((entry) => {
      var sheetType = entry.sheetType;
      if (sheetType === "IDS Master") {
        return;
      }
      var idMasterSpreadsheetInfo = shared.findSheetTypeID(
        idMasterID,
        "IDS",
        sheetType,
        idsMasterValues,
      );
      if (
        idMasterSpreadsheetInfo &&
        idMasterSpreadsheetInfo.cell &&
        idMasterSpreadsheetInfo.cell.range
      ) {
        batchUpdate.push({
          range: idMasterSpreadsheetInfo.cell.range,
          values: [[entry.newSheetID]],
        });
      }
    });

    if (batchUpdate.length > 0) {
      try {
        SheetsAPI.batchUpdateValues(idMasterID, [batchUpdate]);
      } catch (error) {
        var errorReport = errors.report("updateIdsMaster", error, {
          note: `Error updating ID Master sheet`,
          idMasterID: idMasterID,
          idDataEntries: idDataEntries,
        });
        return errors.fail(errorReport);
      }
    }

    CacheManager.RemoveSpreadsheet("idMasterSpreadsheet");
    return {
      success: true,
      message: "New IDS Master set successfully",
      gid: idMasterIDSheet.sheetId,
    };
  } catch (error) {
    var errorReport = errors.report("setNewIdsMaster", error, {
      idMasterID: idMasterID,
      idDataEntries: idDataEntries,
    });
    return errors.fail(errorReport, null, {
      gid: idMasterIDSheet.sheetId,
    });
  }
}

/**
 * The grid id of the IDS Master's IDS tab.
 * @param {string} idMasterID
 * @returns {{success: boolean, gid: number}} A failure envelope on error.
 */
function getIdsMasterGid(idMasterID) {
  try {
    var idsMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
    if (!idsMasterSpreadsheet) {
      return errors.reject(
        "getIdsMasterGid",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `IDS Master Spreadsheet™ not found with ID: ${idMasterID}` },
      );
    }

    var idMasterIDSheet = SheetsAPI.getSheetByName(idsMasterSpreadsheet, "IDS");
    if (!idMasterIDSheet) {
      return errors.reject(
        "getIdsMasterGid",
        errors.CODES.SHEET_STRUCTURE,
        `IDS sheet™ not found in ID master spreadsheet™`,
      );
    }

    return {
      success: true,
      message: "IDS Master IDs already set during import",
      gid: idMasterIDSheet.sheetId,
    };
  } catch (error) {
    var errorReport = errors.report("getIdsMasterGid", error, {
      idMasterID: idMasterID,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Client-callable. Compares a sheet's version against its template's.
 * @param {string} sheetID
 * @param {string} sheetType
 * @param {boolean} [forceRefresh]
 * @returns {{success: boolean, comparisonResult: string}} A failure envelope on error.
 */
function compareSheetVersions(sheetID, sheetType, forceRefresh = false) {
  var spreadsheet = spreadsheets(`${sheetType} spreadsheet`, sheetID);
  if (!spreadsheet) {
    return errors.reject(
      "compareSheetVersions",
      errors.CODES.NOT_FOUND,
      "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
      null,
      { note: `Spreadsheet™ not found with ID: ${sheetID}` },
    );
  }
  var homePageSheet = SheetsAPI.getSheetByName(spreadsheet, "Home Page");
  if (!homePageSheet) {
    return errors.reject(
      "compareSheetVersions",
      errors.CODES.SHEET_STRUCTURE,
      `Home Page sheet™ not found in ${sheetType} spreadsheet™`,
    );
  }

  var homePageData = SheetsAPI.batchGetValues(
    sheetID,
    ["Home Page"],
    true,
    forceRefresh,
  );
  if (!homePageData || homePageData.length === 0) {
    return errors.reject(
      "compareSheetVersions",
      errors.CODES.SHEET_STRUCTURE,
      `Could not fetch Home Page data from sheet`,
    );
  }
  var homePageValues = homePageData[0].values;
  var versionInfo = shared.findSheetVersion(
    sheetID,
    "Home Page",
    sheetType,
    homePageValues,
  );
  if (
    !versionInfo ||
    !versionInfo.currentVersion ||
    !versionInfo.latestVersion
  ) {
    return errors.reject(
      "compareSheetVersions",
      errors.CODES.SHEET_STRUCTURE,
      `Could not find complete version information in Home Page sheet™`,
    );
  }
  var comparisonResult = shared.compareVersions(
    versionInfo.currentVersion,
    versionInfo.latestVersion,
  );
  if (comparisonResult !== "older") {
    console.log(`Sheet is up to date`);
    return {
      success: true,
      currentVersion: versionInfo.currentVersion,
      latestVersion: versionInfo.latestVersion,
      comparisonResult: comparisonResult,
    };
  }

  return {
    success: true,
    currentVersion: versionInfo.currentVersion,
    latestVersion: versionInfo.latestVersion,
    comparisonResult: comparisonResult,
  };
}

/**
 * Client-callable. An OAuth token for the Picker.
 * @returns {string}
 */
function getOAuthToken() {
  try {
    const token = ScriptApp.getOAuthToken();
    return {
      success: true,
      token: token,
      authorizationUrl: "",
      message: "Token retrieved successfully",
    };
  } catch (error) {
    var errorReport = errors.report("getOAuthToken", error, { note: `Error getting OAuth token` });

    return errors.fail(errorReport, null, {
      token: null,
      authorizationUrl: getScopeAuthorizationUrl(),
    });
  }
}

/**
 * Client-callable. The URL where the user grants missing scopes.
 * @returns {string} "" when it cannot be built.
 */
function getScopeAuthorizationUrl() {
  try {
    return (
      ScriptApp.getAuthorizationInfo(
        ScriptApp.AuthMode.FULL,
      ).getAuthorizationUrl() || ""
    );
  } catch (authInfoError) {
    errors.report("getScopeAuthorizationUrl", authInfoError, { note: `Error getting authorization URL` }, errors.CODES.RECOVERED);
    return "";
  }
}

/**
 * Client-callable. Whether every required scope is granted.
 * @returns {boolean}
 */
function checkScopePermissions() {
  try {
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
    return true;
  } catch (error) {
    errors.report(
      "checkScopePermissions",
      error,
      { note: `Scope permission check failed` }, errors.CODES.ACCESS_DENIED);
    return false;
  }
}

/**
 * Client-callable. Whether the script can reach a sheet, and how.
 * @param {string} sheetID
 * @returns {{success: boolean, accessible: boolean, owned: boolean, canEdit: boolean, sheetID: string}}
 */
function checkSheetAccess(sheetID) {
  try {
    if (!sheetID) {
      return errors.reject(
        "checkSheetAccess",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        {
          accessible: false,
          owned: false,
          canEdit: false,
        },
        { note: "Missing sheetID parameter" },
      );
    }

    try {
      const file = CacheManager.getFile(sheetID);
      const parentFolderID =
        file.parents && file.parents.length > 0 ? file.parents[0] : null;
      const owners = file.owners || [];
      const isOwner = owners.some((owner) => owner.me === true);
      const canEdit = !!(file.capabilities && file.capabilities.canEdit);

      return {
        success: true,
        message: `Sheet access verified`,
        accessible: true,
        owned: isOwner,
        canEdit: canEdit,
        sheetID: sheetID,
        name: file.name,
        parentFolderID: parentFolderID,
      };
    } catch (error) {
      errors.report(
        "checkSheetAccess",
        error,
        { note: `Sheet access denied for ${sheetID}`, sheetID: sheetID }, errors.CODES.ACCESS_DENIED);

      return {
        success: true,
        message: `Sheet access denied`,
        accessible: false,
        owned: false,
        canEdit: false,
        sheetID: sheetID,
      };
    }
  } catch (error) {
    var errorReport = errors.report("checkSheetAccess", error, {
      note: `Error checking sheet access`,
      sheetID: sheetID,
    });
    return errors.fail(errorReport, null, {
      accessible: false,
      owned: false,
      canEdit: false,
    });
  }
}

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

    var spreadsheetInfo = shared.findSheetTypeURL(
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

    var oldSheetID = shared.extractSheetId(spreadsheetInfo.id);
    if (!oldSheetID) {
      console.log(`Could not extract old sheet ID from ${spreadsheetInfo.id}`);
      return {
        success: false,
        message: `Could not extract old sheet™ ID from ${spreadsheetInfo.id}`,
      };
    }

    var versionLoading =
      shared.isVersionLoading(spreadsheetInfo.version.value) ||
      shared.isVersionLoading(spreadsheetInfo.oldVersion.value);
    var templateVersion = shared.isVersionLoading(spreadsheetInfo.version.value)
      ? ""
      : spreadsheetInfo.version.value;
    var oldVersion = shared.isVersionLoading(spreadsheetInfo.oldVersion.value)
      ? ""
      : spreadsheetInfo.oldVersion.value;

    var templateStatus = shared.getVersionStatus(templateVersion);
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

      var versionComparison = shared.compareVersions(
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
      templateUrl = shared.extractUrlFromHyperlink(
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
 * Client-callable. Whether the script can reach a template.
 * @param {string} templateID
 * @returns {{success: boolean, accessible: boolean, templateID: string}}
 */
function checkTemplateAccess(templateID) {
  try {
    if (!templateID) {
      return errors.reject(
        "checkTemplateAccess",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        {
          accessible: false,
        },
        { note: "Missing templateID parameter" },
      );
    }

    try {
      var file = CacheManager.getFile(templateID);

      return {
        success: true,
        message: `Template access verified`,
        accessible: true,
        templateID: templateID,
        name: file.name,
      };
    } catch (error) {
      errors.report(
        "checkTemplateAccess",
        error,
        { note: `Template access denied for ${templateID}`, templateID: templateID }, errors.CODES.ACCESS_DENIED);

      return {
        success: true,
        message: `Template access denied`,
        accessible: false,
        templateID: templateID,
      };
    }
  } catch (error) {
    var errorReport = errors.report("checkTemplateAccess", error, {
      note: `Error checking template access`,
      templateID: templateID,
    });
    return errors.fail(errorReport, null, {
      accessible: false,
    });
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
  var spreadsheetInfo = shared.findSheetTypeID(sheetID, "IDS", sheetType);
  if (!spreadsheetInfo || !spreadsheetInfo.id) {
    return errors.reject(
      "findSheetIdAndType",
      errors.CODES.SHEET_STRUCTURE,
      `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
    );
  }
  console.log(`Found sheet type ID: ${spreadsheetInfo.id}`);
  var spreadsheetId = shared.extractSheetId(spreadsheetInfo.id);
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

    var spreadsheetInfo = shared.findSheetTypeURL(
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

    var oldSheetID = shared.extractSheetId(spreadsheetInfo.id);
    if (!oldSheetID) {
      console.log(`Could not extract old sheet ID from ${spreadsheetInfo.id}`);
      return {
        success: false,
        message: `Could not extract old sheet™ ID from ${spreadsheetInfo.id}`,
      };
    }

    var versionLoading =
      shared.isVersionLoading(spreadsheetInfo.version.value) ||
      shared.isVersionLoading(spreadsheetInfo.oldVersion.value);
    var templateVersion = shared.isVersionLoading(spreadsheetInfo.version.value)
      ? ""
      : spreadsheetInfo.version.value;
    var oldVersion = shared.isVersionLoading(spreadsheetInfo.oldVersion.value)
      ? ""
      : spreadsheetInfo.oldVersion.value;

    var templateStatus = shared.getVersionStatus(templateVersion);
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

      var versionComparison = shared.compareVersions(
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
      templateUrl = shared.extractUrlFromHyperlink(
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

    var templateID = shared.extractSheetId(templateUrl);
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
 * Client-callable. Template access for one sheet type in an IDS Master.
 * @param {string} idMasterID
 * @param {string} sheetType
 * @returns {Object} What processTemplateAccess returned.
 */
function checkFileTemplateAccess(idMasterID, sheetType) {

  var result = null;
  for (var attempt = 1; attempt <= 2; attempt++) {
    if (attempt > 1) Utilities.sleep(700);

    var idsMasterData = fetchIdsMasterData(idMasterID, attempt > 1);
    if (!idsMasterData.success) {
      return idsMasterData;
    }

    result = processTemplateAccess(idsMasterData, sheetType, "all");
    if (!result || !result.versionLoading) break;
  }

  return result;
}

/**
 * Client-callable. The sheet each save-file category imports into.
 * @param {string} idMasterID
 * @param {string[]} [sheetTypes]
 * @returns {{success: boolean, targets: Object}} A failure envelope on error.
 */
function getSaveFileImportTargets(idMasterID, sheetTypes) {
  try {
    var resolvedIdMasterID = idMasterID
      ? shared.extractSheetId(String(idMasterID))
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
      return shared.readVersion(cell && cell.value);
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

        var sheetTypeInfo = shared.findSheetTypeURL(
          resolvedIdMasterID,
          "IDS",
          sheetType,
          values,
        );

        var targetID = sheetTypeInfo && sheetTypeInfo.id
          ? shared.extractSheetId(sheetTypeInfo.id)
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
            ? shared.compareVersions(currentVersion, latestVersion) !== "older"
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

    var versionInfo = shared.findSheetVersion(
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
 * Client-callable. Whether the new sheet points back at its IDS Master.
 * @param {string} newSheetID
 * @param {string} sheetType
 * @returns {{success: boolean}} A failure envelope on error.
 */
function checkNewSheetReference(newSheetID, sheetType) {
  try {
    if (!newSheetID) {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing newSheetID parameter." },
      );
    }

    var newSpreadsheet = spreadsheets(
      `${sheetType} newSpreadsheet`,
      newSheetID,
    );
    if (!newSpreadsheet) {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `New spreadsheet™ not found with ID: ${newSheetID}` },
      );
    }

    var newIdSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
    if (!newIdSheet) {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.SHEET_STRUCTURE,
        `IDS sheet™ not found in the new spreadsheet™.`,
      );
    }

    var newSpreadsheetInfo = shared.findSheetTypeID(newSheetID, "IDS");
    if (!newSpreadsheetInfo || !newSpreadsheetInfo.id) {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.SHEET_STRUCTURE,
        `Could not find sheet type ID for ${newSheetID}`,
      );
    }
    var accessStatus = newSpreadsheetInfo.accessStatus;
    if (!accessStatus || accessStatus.value !== "✅") {
      return errors.reject(
        "checkNewSheetReference",
        errors.CODES.ACCESS_DENIED,
        `New sheet have not been granted access to IDS Master.`,
      );
    }

    return {
      success: true,
      message: `New sheet reference is valid.`,
    };
  } catch (error) {
    var errorReport = errors.report("checkNewSheetReference", error, {
      note: `Error checking new sheet reference`,
      newSheetID: newSheetID,
      sheetType: sheetType,
    });
    return errors.fail(errorReport);
  }
}

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

      var sheetTypeInfo = shared.findSheetTypeURL(
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

      var oldSheetID = shared.extractSheetId(sheetTypeInfo.id);
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

      var oldVersion = shared.readVersion(sheetTypeInfo.oldVersion.value);
      var templateVersion = shared.readVersion(sheetTypeInfo.version.value);

      console.log(
        `oldVersion: ${oldVersion}, templateVersion: ${templateVersion}`,
      );
      var versionDifference = null;
      if (oldVersion && templateVersion) {
        var versionComparison = shared.compareVersions(
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
    var spreadsheetInfo = shared.findSheetTemplateID(
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

/**
 * Client-callable. Whether a sheet's version has a converter.
 * @param {string} oldSheetID
 * @param {string} sheetType
 * @returns {{success: boolean, versionDifference: string}} A failure envelope on error.
 */
function checkExportCompatibility(oldSheetID, sheetType) {
  try {
    if (!oldSheetID) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing oldSheetID parameter." },
      );
    }
    if (!sheetType) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing sheetType parameter." },
      );
    }

    var oldSpreadsheet = spreadsheets(
      `${sheetType} oldSpreadsheet`,
      oldSheetID,
    );
    if (!oldSpreadsheet) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `Could not access old ${sheetType} spreadsheet with ID: ${oldSheetID}` },
      );
    }

    var oldHomePageSheet = SheetsAPI.getSheetByName(
      oldSpreadsheet,
      "Home Page",
    );
    if (!oldHomePageSheet) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.SHEET_STRUCTURE,
        `Home Page sheet not found in old ${sheetType} spreadsheet`,
      );
    }

    var oldHomePageData = SheetsAPI.batchGetValues(oldSheetID, ["Home Page"]);
    if (!oldHomePageData || oldHomePageData.length === 0) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.SHEET_STRUCTURE,
        `Could not read Home Page data from old ${sheetType} spreadsheet`,
      );
    }

    var oldHomePageValues = oldHomePageData[0].values;

    var oldVersionInfo = shared.findSheetVersion(
      oldSheetID,
      "Home Page",
      sheetType,
      oldHomePageValues,
    );
    if (!oldVersionInfo || !oldVersionInfo.currentVersion) {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.INTERNAL,
        `Current Version not found in old ${sheetType} spreadsheet.`,
      );
    }
    var oldVersion = oldVersionInfo.currentVersion;

    var sheetTypeFunction = sheetVars(sheetType);
    if (sheetTypeFunction) {
      var versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
      if (!versionDifference) {
        return errors.reject(
          "checkExportCompatibility",
          errors.CODES.VERSION_OUTDATED,
          `Your ${sheetType} is version ${oldVersion}, which this script cannot read.`,
        );
      }
      return {
        success: true,
        message: `Old ${sheetType} version (${oldVersion}) is compatible for export`,
        oldVersion: oldVersion,
        versionDifference: versionDifference,
      };
    } else {
      return errors.reject(
        "checkExportCompatibility",
        errors.CODES.SHEET_STRUCTURE,
        `No compatibility function found for ${sheetType}. Cannot verify export compatibility.`,
      );
    }
  } catch (error) {
    var errorReport = errors.report("checkExportCompatibility", error, {
      note: `Error checking export compatibility`,
      oldSheetID: oldSheetID,
      sheetType: sheetType,
    });
    return errors.fail(errorReport);
  }
}

/**
 * Writes one sheet's ID into another sheet's IDS tab.
 * @param {string} spreadsheetID Where to write.
 * @param {string} sheetID What to write.
 * @param {string} sheetType
 * @returns {{success: boolean, message: string}} A failure envelope on error.
 */
function updateSheetID(spreadsheetID, sheetID, sheetType) {
  try {
    if (!spreadsheetID) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing spreadsheetID parameter." },
      );
    }
    if (!sheetID) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing sheetID parameter." },
      );
    }
    if (!sheetType) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.INVALID_INPUT,
        "Something was missing from that request. Please reload the page and try again.",
        null,
        { note: "Missing sheetType parameter." },
      );
    }
    var spreadsheet = spreadsheets(`${sheetType} spreadsheet`, spreadsheetID);
    if (!spreadsheet) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.NOT_FOUND,
        "The script could not open that sheet. It may have been deleted, or access to it was never granted.",
        null,
        { note: `Could not access spreadsheet with ID: ${spreadsheetID}` },
      );
    }
    var idSheet = SheetsAPI.getSheetByName(spreadsheet, "IDS");
    if (!idSheet) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.SHEET_STRUCTURE,
        "The script could not find a tab it needs in your sheet.",
        null,
        { note: `IDS sheet not found in spreadsheet with ID: ${spreadsheetID}` },
      );
    }
    var idValues = SheetsAPI.batchGetValues(spreadsheetID, ["IDS"]);
    if (!idValues || idValues.length === 0) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.SHEET_STRUCTURE,
        "The script could not find a tab it needs in your sheet.",
        null,
        { note: `Could not read IDS sheet data from spreadsheet with ID: ${spreadsheetID}` },
      );
    }
    var values = idValues[0].values;
    var sheetTypeInfo = shared.findSheetTypeID(
      spreadsheetID,
      "IDS",
      "IDS Master",
      values,
    );
    if (!sheetTypeInfo || !sheetTypeInfo.cell) {
      return errors.reject(
        "updateSheetID",
        errors.CODES.SHEET_STRUCTURE,
        `Could not find IDS Master entry in IDS sheet.`,
      );
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
    var errorReport = errors.report("updateSheetID", error, {
      note: `Error updating sheet ID`,
      spreadsheetID: spreadsheetID,
      sheetID: sheetID,
      sheetType: sheetType,
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
      var ownSheetInfo = shared.findSheetTypeID(
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
        var sheetInfo = shared.findSheetVersion(
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
        var sheetInfo = shared.findSheetVersion(
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

    batchUpdate = shared.addIDUpdatesToBatch(
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
      var sheetInfo = shared.findSheetVersion(
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

/**
 * Client-callable. Whether a linked file is an IDS Master or Collection,
 * and whether it is out of date.
 * @param {string} sheetID
 * @returns {{success: boolean, sheetType: string, outdated?: boolean}} A failure envelope on error.
 */
function getSaveFileSheetType(sheetID) {
  try {
    var resolvedID = sheetID
      ? shared.extractSheetId(String(sheetID)) || ""
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
      var versionInfo = shared.findSheetVersion(
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
          shared.compareVersions(
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
