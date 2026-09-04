const CacheManager = {
  _userCache: null,

  get userCache() {
    if (!this._userCache) {
      try {
        this._userCache = CacheService.getUserCache();
      } catch (error) {
        errors.reportFinal("CacheManager.userCache", error);
        return null;
      }
    }
    return this._userCache;
  },

  CHUNK_SIZE: 90000,

  /**
   * @private
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
        // Surrogate pair - four bytes across the two code units, counted once
        bytes += 4;
        i++;
      } else {
        bytes += 3;
      }
    }

    return bytes;
  },

  /**
   * Split a string into chunks that each fit the byte budget, never cutting a
   * surrogate pair in half
   * @private
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
   * How many chunks the value currently stored under a key is split into, or 0
   * when it is not stored chunked
   * @private
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
   * Every key a stored value occupies - the value itself plus its chunks and
   * their marker - so an entry can be invalidated whichever way it was stored
   * @private
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
   * Retrieve a value, automatically combining chunks if needed
   * Returns null if the cache is unavailable, or if any chunk of a chunked
   * value has been evicted - a partial value reads as corrupt, not as a miss
   * @private
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
   * @private
   */
  _putValue: function (key, value) {
    this._putAllValues({ [key]: value });
  },

  /**
   * @private
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

        // The unchunked copy, plus any chunk a shorter new value leaves behind
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
   * Get or fetch spreadsheet metadata with smart cache invalidation
   * Stores by spreadsheet type name (e.g., "Laboratory oldSpreadsheet")
   * Validates sheetID against cached value and overwrites if different
   * Automatically handles chunking for large spreadsheet metadata
   * @param {string} spreadsheetTypeName - e.g., "Laboratory oldSpreadsheet", "Workshop newSpreadsheet"
   * @param {string} sheetID - Sheet ID to fetch/validate
   * @returns {Object} { spreadsheet metadata, sheetID } or null
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
   * Cache sheet values with key "sheetID|sheetName|VALUE"
   * Handles partial caching - only fetches uncached ranges
   * Automatically chunks large values into multiple keys
   */
  getSheetValues: function (spreadsheetId, ranges, forceRefresh = false) {
    const cachedData = [];
    const uncachedRanges = [];
    const uncachedIndices = [];

    const cacheKeys = ranges.map((range) => `${spreadsheetId}|${range}|VALUE`);

    for (let i = 0; i < cacheKeys.length; i++) {
      // `forceRefresh` treats every range as a miss, so the values below are
      // fetched again and written back over the stale entry - for a caller
      // that has reason to believe what it read was a live formula caught
      // mid-calculation.
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
   * Cache sheet formulas with key "sheetID|sheetName|FORMULA"
   * Automatically chunks large values into multiple keys
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
   * Clear cache for a specific spreadsheet type and all its sheet data
   * Also removes chunked key variants (__chunk_0, __chunks, etc)
   */
  RemoveSpreadsheet: function (spreadsheetTypeName) {
    if (!this.userCache) {
      console.log(`Cache unavailable - cannot remove: ${spreadsheetTypeName}`);
      return;
    }

    try {
      // Read through _retrieveValue, not the raw key: a chunked entry has
      // nothing stored under the plain key, and reading it directly would make
      // this look like there was nothing to invalidate.
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
      errors.reportFinal("cacheData.RemoveSpreadsheet", error, {
        note: `Error invalidating cache`,
        spreadsheetTypeName: spreadsheetTypeName,
      });
    }
  },

  /**
   * Get or fetch file metadata from Drive with caching
   * Caches only the fields we need (id, name, parents, owners/me,
   * capabilities/canEdit, trashed)
   * Automatically handles chunking for large file metadata
   * @param {string} fileID - File ID to fetch/cache
   * @returns {Object} File metadata or null
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
      errors.reportFinal("cacheData.getFile", error, {
        note: `Error fetching file`,
        fileID: fileID,
      });
    }

    return null;
  },

  /**
   * Clear cache for a specific file
   * Also removes any chunked key variants
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
   * A version cell fed by a live formula reads "Loading..." for as long as
   * Sheets is still calculating it. That string is truthy and contains no
   * digits, so anything that trusts it draws the wrong conclusion — most
   * damagingly `compareVersions`, which parses it to no version at all and
   * therefore reports the template as "not newer", silently skipping a sheet
   * that really does need updating. Every read of a version cell tests for
   * it through here.
   */
  isVersionLoading: function (value) {
    return (
      String(value == null ? "" : value)
        .trim()
        .toLowerCase()
        .indexOf("loading") === 0
    );
  },

  /** Trimmed version text, or "" for a cell that is still calculating. */
  readVersion: function (value) {
    if (value == null) return "";
    var text = String(value).trim();
    return shared.isVersionLoading(text) ? "" : text;
  },

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

  compareVersions: function (oldVersion, newVersion) {
    // Dev only
    // if (newVersion.includes("WIP")) {
    //   return "older";
    // }
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

  /**
   * The named presets the sheet templates ship with. Everything else is called
   * "Preset N". None of these are guaranteed to exist in a user's data - they
   * are only the labels a freshly copied template starts out with.
   */
  templatePresetNames: ["Farming", "Tourney"],
  
  /**
   * Reorders a save file's preset names into the fixed slot order the sheet
   * templates expect (e.g. ["Farming", "Tourney", "Preset 3", "Preset 4", "Preset 5"]),
   * without requiring "Farming"/"Tourney" to already be first, or to exist at all.
   * Any name in forcedNames found anywhere in presetNames is pulled into its
   * matching slot; everything else fills the remaining slots in its original
   * relative order.
   * @param {Array} presetNames - preset names as they appear in the save file
   * @param {Array} forcedNames - names that should be pulled into the front slots when present, e.g. ["Farming", "Tourney"]
   * @returns {{order: Array, indices: Array}} order: slot-ordered names (empty slots fall back to "Preset N"); indices: for each slot, the original index in presetNames to pull parallel data from
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
      errors.reportFinal("shared.addIDUpdatesToBatch", error, {
        note: `Error adding ID updates to batch`,
        batchUpdate: batchUpdate,
        sheetType: sheetType,
        newSheetID: newSheetID,
        idsData: idsData,
        idMasterID: idMasterID,
      });
      return batchUpdate;
    }
  },
};

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
    // The merged sheet takes over from the two sheets it replaced, so the name
    // inherited from whichever one it was built from is renamed with it.
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

    // A merged sheet replaces more than one old sheet, so every sheet it was
    // built from is deleted. The new sheet is already moved and renamed at this
    // point, so a leftover here is logged rather than failing the whole move.
    var extraOldSheetIDs = (mergedOldSheetIDs || []).filter(function (sheetID) {
      return sheetID && sheetID !== oldSheetID;
    });
    for (var i = 0; i < extraOldSheetIDs.length; i++) {
      try {
        Drive.Files.update({ trashed: true }, extraOldSheetIDs[i]);
        CacheManager.RemoveFile(extraOldSheetIDs[i]);
        console.log(`Deleted merged old sheet: ${extraOldSheetIDs[i]}`);
      } catch (error) {
        errors.reportFinal("deleteOldSheet", error, {
          note: `Error deleting merged old sheet`,
          sheetID: extraOldSheetIDs[i],
        });
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
 * The combined update writes the new IDS Master's sheet IDs as part of its own
 * import, so nothing is left to set afterwards. This resolves just the IDS tab's
 * gid for the summary link, without the read and write updateIdsMaster spends on
 * IDs that are already in place.
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

function getScopeAuthorizationUrl() {
  try {
    return (
      ScriptApp.getAuthorizationInfo(
        ScriptApp.AuthMode.FULL,
      ).getAuthorizationUrl() || ""
    );
  } catch (authInfoError) {
    errors.reportFinal("getScopeAuthorizationUrl", authInfoError, { note: `Error getting authorization URL` });
    return "";
  }
}

function checkScopePermissions() {
  try {
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
    return true;
  } catch (error) {
    errors.reportFinal(
      "checkScopePermissions",
      error,
      { note: `Scope permission check failed` },
      errors.CODES.ACCESS_DENIED,
    );
    return false;
  }
}

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
      errors.reportFinal(
        "checkSheetAccess",
        error,
        { note: `Sheet access denied for ${sheetID}`, sheetID: sheetID },
        errors.CODES.ACCESS_DENIED,
      );

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

function getTemplateAndsheetIds(idMasterID, copyMode) {
  try {
    // "Themes, Songs & Relics" only exists on a v4.0 or later IDS Master; older
    // ones still have the two sheet types it replaced. It is looked up first
    // because "Relics" matches it by name too, and the two old sheet types are
    // skipped once it has been found.
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

    // Version cells still calculating read "Loading...", and deciding what to
    // update off one of those would skip a sheet that does need updating. One
    // re-read past the cache, then we go with what we have - see
    // `shared.isVersionLoading` and `getSaveFileImportTargets`, which does the
    // same thing for the save-file workflow.
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
          errors.reportFinal("getTemplateAndsheetIds", templateError, {
            note: `Error processing template for ${sheetType}`,
            idMasterID: idMasterID,
            copyMode: copyMode,
          });
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

    // Blank a version cell we caught mid-calculation rather than comparing
    // against it, and say so, so the caller can re-read instead of skipping a
    // sheet on the strength of a version nobody ever actually saw.
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
      errors.reportFinal(
        "checkTemplateAccess",
        error,
        { note: `Template access denied for ${templateID}`, templateID: templateID },
        errors.CODES.ACCESS_DENIED,
      );

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
 * @param {string} idMasterID
 * @param {boolean} [forceRefresh] Re-read the IDS sheet from the API even if
 *        it is in the cache. For a caller that read a live formula still
 *        showing "Loading..." and wants to give it another go - a plain
 *        re-call would otherwise be served the same cached copy.
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

    // Same as in `getTemplateInfo`: a cell still calculating is not a version,
    // and must not be compared against - see `shared.isVersionLoading`.
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
      errors.reportFinal("processTemplateAccess", error, {
        note: `Error retrieving template file information`,
        idsMasterData: idsMasterData,
        sheetType: sheetType,
        copyMode: copyMode,
      });
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

function checkFileTemplateAccess(idMasterID, sheetType) {
  // One re-read past the cache when a version cell was still calculating -
  // see `shared.isVersionLoading`.
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

    // When no specific types are requested, resolve every save-file category so
    // the caller can grant access to all linked subsheets up front.
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

    function readSheetVersion(cell) {
      return shared.readVersion(cell && cell.value);
    }

    // Version cells here come from live formulas pulling each template's
    // version, so if the IDS Master was opened moments ago Sheets can still
    // be showing "Loading..." in one or more of them at the exact moment we
    // read. A sheet whose link we DID resolve but whose version came back
    // blank is almost always that, not a structural problem, so it is worth
    // one re-read before giving up and showing "?". The retry has to go past
    // the cache: the read that returned "Loading..." cached that answer, and
    // a plain re-call would be handed the same placeholder straight back.
    // Still only the IDS Master and its Home Page - never a fetch per
    // template, which would cost real quota for no better odds.
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

        // The IDS Master is not one of its own linked subsheets - it is the sheet
        // we were handed - so there is no IDS row to look it up in. Its version
        // comes from its own Home Page instead.
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

      // A cell still calculating is not a version to compare against - blank
      // it and let the guard below skip the comparison, the same as it does
      // for a version that is genuinely absent.
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

function deleteOldSheet(sheetID) {
  try {
    console.log(`Attempting to delete sheet with ID: ${sheetID}`);

    var fileInfo;
    try {
      fileInfo = CacheManager.getFile(sheetID);
    } catch (error) {
      // Already gone is a fine outcome here - the function returns success
      // for exactly this case.
      errors.reportFinal(
        "deleteOldSheet",
        error,
        { note: `Sheet ${sheetID} not found or already deleted`, sheetID: sheetID },
        errors.CODES.NOT_FOUND,
      );
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
        errors.reportFinal("updateGetStartedSheetIdsAndReferences", error, {
          note: `Could not update file name with version info`,
          sheetID: sheetID,
          sheetType: sheetType,
          relatedSheetIDs: relatedSheetIDs,
        });
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
        errors.reportFinal("updateGetStartedSheetIdsAndReferences", error, {
          note: `Could not update file name with version info`,
          sheetID: sheetID,
          sheetType: sheetType,
          relatedSheetIDs: relatedSheetIDs,
        });
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
      errors.reportFinal("updateGetStartedSheetIdsAndReferences", error, {
        note: `Could not update file name with version info`,
        sheetID: sheetID,
        sheetType: sheetType,
        relatedSheetIDs: relatedSheetIDs,
      });
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

