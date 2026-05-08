// Cache Manager - handles spreadsheet metadata and sheet values caching
const CacheManager = {
  _userCache: null,

  get userCache() {
    if (!this._userCache) {
      try {
        this._userCache = CacheService.getUserCache();
      } catch (error) {
        console.error(`Failed to initialize userCache: ${error.message}`);
        return null;
      }
    }
    return this._userCache;
  },

  CHUNK_SIZE: 90000,

  /**
   * Split a large string into chunks for storage
   * @private
   */
  _chunkString: function (str, chunkSize = this.CHUNK_SIZE) {
    const chunks = [];
    for (let i = 0; i < str.length; i += chunkSize) {
      chunks.push(str.substring(i, i + chunkSize));
    }
    return chunks;
  },

  /**
   * Prepare cache data with automatic chunking for large values
   * Splits values > CHUNK_SIZE into multiple keys with metadata
   * @private
   */
  _prepareCacheData: function (cacheData) {
    const result = {};

    for (const key in cacheData) {
      const value = cacheData[key];

      if (value.length > this.CHUNK_SIZE) {
        const chunks = this._chunkString(value, this.CHUNK_SIZE);

        for (let i = 0; i < chunks.length; i++) {
          result[`${key}__chunk_${i}`] = chunks[i];
        }

        result[`${key}__chunks`] = chunks.length.toString();
        console.log(
          `Chunked ${key} into ${chunks.length} parts (${value.length} bytes total)`,
        );
      } else {
        result[key] = value;
      }
    }

    return result;
  },

  /**
   * Retrieve a value, automatically combining chunks if needed
   * Returns null if cache is unavailable
   * @private
   */
  _retrieveValue: function (key) {
    if (!this.userCache) {
      console.log(`Cache unavailable - cannot retrieve: ${key}`);
      return null;
    }

    const chunksCountStr = this.userCache.get(`${key}__chunks`);

    if (chunksCountStr) {
      const chunkCount = parseInt(chunksCountStr);
      let combinedValue = "";

      for (let i = 0; i < chunkCount; i++) {
        const chunkKey = `${key}__chunk_${i}`;
        const chunk = this.userCache.get(chunkKey);
        if (chunk) {
          combinedValue += chunk;
        }
      }

      return combinedValue;
    } else {
      return this.userCache.get(key);
    }
  },

  /**
   * Store a single value, automatically chunking if needed
   * Silently fails if cache is unavailable
   * @private
   */
  _putValue: function (key, value) {
    if (!this.userCache) {
      console.log(`Cache unavailable - cannot store: ${key}`);
      return;
    }

    if (value.length > this.CHUNK_SIZE) {
      const cacheData = this._prepareCacheData({ [key]: value });
      this.userCache.putAll(cacheData);
    } else {
      this.userCache.put(key, value);
    }
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
  getSheetValues: function (spreadsheetId, ranges) {
    const cachedData = [];
    const uncachedRanges = [];
    const uncachedIndices = [];

    const cacheKeys = ranges.map((range) => `${spreadsheetId}|${range}|VALUE`);

    for (let i = 0; i < cacheKeys.length; i++) {
      const cached = this._retrieveValue(cacheKeys[i]);

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
        const chunkedCacheData = this._prepareCacheData(cacheData);
        this.userCache.putAll(chunkedCacheData);
      }
    }

    return result;
  },

  /**
   * Cache sheet formulas with key "sheetID|sheetName|FORMULA"
   * Automatically chunks large values into multiple keys
   */
  getSheetFormulas: function (spreadsheetId, ranges) {
    const cachedData = [];
    const uncachedRanges = [];
    const uncachedIndices = [];

    const cacheKeys = ranges.map(
      (range) => `${spreadsheetId}|${range}|FORMULA`,
    );

    for (let i = 0; i < cacheKeys.length; i++) {
      const cached = this._retrieveValue(cacheKeys[i]);

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
        const chunkedCacheData = this._prepareCacheData(cacheData);
        this.userCache.putAll(chunkedCacheData);
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
      const cached = this.userCache.get(spreadsheetTypeName);
      if (!cached) {
        console.log(
          `No cache entry found for ${spreadsheetTypeName} to invalidate`,
        );
        return;
      }

      const cachedData = JSON.parse(cached);
      const sheetID = cachedData.sheetID;
      const metadata = cachedData.metadata;

      const keysToRemove = [spreadsheetTypeName];

      if (metadata && metadata.sheets) {
        for (let i = 0; i < metadata.sheets.length; i++) {
          const sheetName = metadata.sheets[i].properties.title;
          const valueKey = `${sheetID}|${sheetName}|VALUE`;
          const formulaKey = `${sheetID}|${sheetName}|FORMULA`;

          keysToRemove.push(valueKey);
          keysToRemove.push(formulaKey);

          const chunksCountValue = this.userCache.get(`${valueKey}__chunks`);
          if (chunksCountValue) {
            const chunkCount = parseInt(chunksCountValue);
            for (let j = 0; j < chunkCount; j++) {
              keysToRemove.push(`${valueKey}__chunk_${j}`);
            }
            keysToRemove.push(`${valueKey}__chunks`);
          }

          const chunksCountFormula = this.userCache.get(
            `${formulaKey}__chunks`,
          );
          if (chunksCountFormula) {
            const chunkCount = parseInt(chunksCountFormula);
            for (let j = 0; j < chunkCount; j++) {
              keysToRemove.push(`${formulaKey}__chunk_${j}`);
            }
            keysToRemove.push(`${formulaKey}__chunks`);
          }
        }
      }

      this.userCache.removeAll(keysToRemove);
      console.log(
        `Invalidated cache for ${spreadsheetTypeName} and ${keysToRemove.length - 1} sheet entries`,
      );
    } catch (error) {
      console.log(`Error invalidating cache: ${error}`);
    }
  },

  /**
   * Get or fetch file metadata from Drive with caching
   * Caches all fields (id, name, parents, owners, trashed)
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

    const allFieldsNeeded = "id, name, parents, owners, trashed";
    try {
      const file = Drive.Files.get(fileID, { fields: allFieldsNeeded });

      if (file) {
        this._putValue(cacheKey, JSON.stringify(file));
        console.log(`Cached file metadata: ${fileID}`);
        return file;
      }
    } catch (error) {
      console.error(`Error fetching file: ${error}`);
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

    const cacheKey = `File|${fileID}`;
    const keysToRemove = [cacheKey];

    const chunksCountStr = this.userCache.get(`${cacheKey}__chunks`);
    if (chunksCountStr) {
      const chunkCount = parseInt(chunksCountStr);
      for (let i = 0; i < chunkCount; i++) {
        keysToRemove.push(`${cacheKey}__chunk_${i}`);
      }
      keysToRemove.push(`${cacheKey}__chunks`);
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
      console.error(`Error getting spreadsheet: ${error}`);
      console.error(`Error details: ${JSON.stringify(error)}`);
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
      console.error(`Error getting sheet by name: ${error}`);
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
      console.error(`Error getting sheet by substring: ${error}`);
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
      console.error(`Error applying sheet visibility: ${error}`);
      return {
        success: false,
        message: `Error updating sheet visibility: ${error.toString()}`,
      };
    }
  },

  batchGetValues: function (spreadsheetId, ranges, useCache = true) {
    try {
      if (useCache) {
        return CacheManager.getSheetValues(spreadsheetId, ranges);
      }
      const response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
        ranges: ranges,
      });
      return response.valueRanges;
    } catch (error) {
      console.error(`Error in batchGetValues: ${error}`);
      return null;
    }
  },

  batchGetFormulas: function (spreadsheetId, ranges, useCache = true) {
    try {
      if (useCache) {
        return CacheManager.getSheetFormulas(spreadsheetId, ranges);
      }
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
      const requestBody = {
        data: updates,
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
      console.error(`Error finding sheet version: ${error}`);
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
      console.error(`Error finding Effective Paths version: ${error}`);
      return null;
    }
  },

  compareVersions: function (oldVersion, newVersion) {
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

    var regex = new RegExp(sheetType, "i");
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (regex.test(values[i][j]) && values[i][j].indexOf("script") === -1) {
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

    var regex = new RegExp(sheetType, "i");
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (
          regex.test(values[i][j]) &&
          values[i][j].indexOf("script") === -1 &&
          values[i][j].indexOf("More IDs are available") === -1
        ) {
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

    var oldLevel = oldValue.split("|")[0].trim();

    console.log(`Looking for DVT match for level: ${oldLevel}`);
    for (var i = 0; i < dvtNamedRangesData.length; i++) {
      var row = dvtNamedRangesData[i];
      var val = row[0] ? row[0].split("|")[0].trim() : null;
      if (val && val === oldLevel) {
        return row[0];
      }
    }
    return oldValue;
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
      console.error(`Error finding template ID: ${error.toString()}`);
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
      console.log(`Error adding ID updates to batch: ${error.toString()}`);
      return batchUpdate;
    }
  },
};

function moveSheet(sheetType, newSheetID, oldSheetID) {
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
    var oldFile = CacheManager.getFile(oldSheetID);
    if (!newFile || !oldFile) {
      console.log(`Could not retrieve file information for new or old sheet.`);
      return {
        success: false,
        message: `Could not retrieve file information for new or old sheet™.`,
      };
    }

    var newVersionInfo;
    newVersionInfo = shared.findSheetVersion(
      newSheetID,
      "Home Page",
      sheetType,
    );

    if (!newVersionInfo || !newVersionInfo.currentVersion) {
      console.log(`Could not find new sheet version`);
      return {
        success: false,
        message: `Could not find new sheet™ version`,
      };
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
    console.log(
      `Updating file name from "${oldFile.name}" to "${newFileName}"`,
    );

    parents = {};
    if (typeof oldFile.parents == "undefined") {
      console.log(`Could not find old file location.`);
      return {
        success: false,
        message: `Could not find old file location.`,
      };
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
      console.log(`Error moving new sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error moving new sheet™: ${error.toString()}`,
      };
    }

    try {
      Drive.Files.update({ trashed: true }, oldSheetID);
    } catch (error) {
      console.log(`Error deleting old sheet: ${error.toString()}`);
      return {
        success: false,
        message: `Error deleting old sheet™: ${error.toString()}`,
      };
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
    console.log(`Error in moveSheet: ${error.toString()}`);
    return {
      success: false,
      message: error.toString(),
    };
  }
}

function updateIdsMaster(idMasterID, idDataEntries) {
  var idsMasterSpreadsheet = spreadsheets("idMasterSpreadsheet", idMasterID);
  if (!idsMasterSpreadsheet) {
    console.log(`IDS Master Spreadsheet not found with ID: ${idMasterID}`);
    return {
      success: false,
      message: `IDS Master Spreadsheet™ not found with ID: ${idMasterID}`,
    };
  }

  var idMasterIDSheet = SheetsAPI.getSheetByName(idsMasterSpreadsheet, "IDS");
  if (!idMasterIDSheet) {
    console.log(`IDS sheet not found in ID master spreadsheet`);
    return {
      success: false,
      message: `IDS sheet™ not found in ID master spreadsheet™`,
    };
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
        console.log(`Error updating ID Master sheet: ${error.toString()}`);
        return {
          success: false,
          message: `Error updating ID Master sheet™: ${error.toString()}`,
        };
      }
    }

    CacheManager.RemoveSpreadsheet("idMasterSpreadsheet");
    return {
      success: true,
      message: "New IDS Master set successfully",
      gid: idMasterIDSheet.sheetId,
    };
  } catch (error) {
    console.log(`Error in setNewIdsMaster: ${error.toString()}`);
    return {
      success: false,
      message: error.toString(),
      gid: idMasterIDSheet.sheetId,
    };
  }
}

function checkImportStatusAndCompatibility(newSheetID, oldSheetID, sheetType) {
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
        imported: false,
      };
    }

    var sheetName = "IDS";
    var searchName = "IDS Master's";
    var requiresIDSSheet = true;

    if (sheetType === "IDS Collection") {
      sheetName = "Home Page";
      searchName = "Load your file here";
      requiresIDSSheet = false;
    }

    if (requiresIDSSheet) {
      var newIdSheet = SheetsAPI.getSheetByName(newSpreadsheet, "IDS");
      if (!newIdSheet) {
        console.log(`IDS sheet not found in new ${sheetType} spreadsheet.`);
        return {
          success: false,
          message: `IDS sheet™ not found in new ${sheetType} spreadsheet™.`,
          imported: false,
        };
      }
    }

    if (sheetType === "Effective Paths") {
      var eHPSheet = SheetsAPI.getSheetByName(newSpreadsheet, "eHP");
      var eDamageSheet = SheetsAPI.getSheetByName(newSpreadsheet, "eDamage");
      var eEconSheet = SheetsAPI.getSheetByName(newSpreadsheet, "eEcon");

      if (!eHPSheet || !eDamageSheet || !eEconSheet) {
        console.log(
          `Required Effective Paths sheets not found in new spreadsheet.`,
        );
        return {
          success: false,
          message: `New spreadsheet™ missing required sheets™ (eHP, eDamage, eEcon).`,
          imported: false,
        };
      }
    }

    var newHomePageSheet = SheetsAPI.getSheetByName(
      newSpreadsheet,
      "Home Page",
    );
    if (!newHomePageSheet) {
      console.log(`Home Page sheet not found in new ${sheetType} spreadsheet`);
      return {
        success: false,
        message: `Home Page sheet™ not found in new ${sheetType} spreadsheet™`,
        imported: false,
      };
    }

    var ranges = requiresIDSSheet ? ["IDS", "Home Page"] : ["Home Page"];
    var newSheetData = SheetsAPI.batchGetValues(newSheetID, ranges);
    if (!newSheetData || newSheetData.length < ranges.length) {
      console.log(`Could not fetch required data from new sheet`);
      return {
        success: false,
        message: `Could not fetch required data from new sheet`,
        imported: false,
      };
    }

    var newIDSValues = newSheetData[0].values;
    var newHomePageValues = requiresIDSSheet
      ? newSheetData[1].values
      : newSheetData[0].values;

    var newVersionInfo;
    newVersionInfo = shared.findSheetVersion(
      newSheetID,
      "Home Page",
      sheetType,
      newHomePageValues,
    );

    if (!newVersionInfo) {
      console.log(`Version not found in new ${sheetType} spreadsheet.`);
      return {
        success: false,
        message: `Version not found in new ${sheetType} spreadsheet™.`,
        imported: false,
      };
    }

    var newVersion = newVersionInfo.currentVersion;
    var latestVersion = newVersionInfo.latestVersion;
    if (!newVersion || !latestVersion) {
      console.log(
        `Version information is incomplete in new ${sheetType} spreadsheet.`,
      );
      return {
        success: false,
        message: `Version information is incomplete in new ${sheetType} spreadsheet™.`,
        imported: false,
      };
    }

    if (newVersion !== latestVersion) {
      console.log(
        `The version of the new sheet (${newVersion}) is not the latest version (${latestVersion}). Please update before importing.`,
      );
      return {
        success: false,
        message: `The version of the new sheet (${newVersion}) is not the latest version (${latestVersion}). Please update before importing.`,
        imported: false,
      };
    }

    var oldSpreadsheet = spreadsheets(
      `${sheetType} oldSpreadsheet`,
      oldSheetID,
    );
    if (!oldSpreadsheet) {
      console.log(`Old spreadsheet not found with ID: ${oldSheetID}`);
      return {
        success: false,
        message: `Old spreadsheet™ not found with ID: ${oldSheetID}`,
        imported: false,
      };
    }

    var oldHomePageSheet = SheetsAPI.getSheetByName(
      oldSpreadsheet,
      "Home Page",
    );
    if (!oldHomePageSheet) {
      console.log(`Home Page sheet not found in old ${sheetType} spreadsheet`);
      return {
        success: false,
        message: `Home Page sheet™ not found in old ${sheetType} spreadsheet™`,
        imported: false,
      };
    }

    var oldHomePageData = SheetsAPI.batchGetValues(oldSheetID, ["Home Page"]);
    if (!oldHomePageData || oldHomePageData.length === 0) {
      console.log(`Could not fetch Home Page data from old sheet`);
      return {
        success: false,
        message: `Could not fetch Home Page data from old sheet`,
        imported: false,
      };
    }

    var oldHomePageValues = oldHomePageData[0].values;
    var versionDifference = "None";
    if (oldHomePageValues[1][1] !== "IDS Master") {
      var oldVersionInfo;
      oldVersionInfo = shared.findSheetVersion(
        oldSheetID,
        "Home Page",
        sheetType,
        oldHomePageValues,
      );

      if (!oldVersionInfo || !oldVersionInfo.currentVersion) {
        console.log(`Version not found in old ${sheetType} spreadsheet.`);
        return {
          success: false,
          message: `Version not found in old ${sheetType} spreadsheet™.`,
          imported: false,
        };
      }

      var oldVersion = oldVersionInfo.currentVersion;
      var compareVersions = shared.compareVersions(oldVersion, newVersion);

      if (compareVersions === "newer") {
        console.log(
          `The version of the old sheet (${oldVersion}) is newer than the new sheet (${newVersion}). Import aborted.`,
        );
        return {
          success: false,
          message: `The version of the old sheet (${oldVersion}) is newer than the new sheet (${newVersion}). Import aborted.`,
          imported: false,
        };
      }

      var sheetTypeFunction = sheetVars(sheetType);
      if (sheetTypeFunction) {
        versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
        if (!versionDifference) {
          console.log(
            `Old version of ${sheetType} is incompatible for import (${oldVersion})`,
          );
          return {
            success: false,
            message: `Old version of ${sheetType} is incompatible for import (${oldVersion}).`,
            imported: false,
          };
        }
      } else {
        console.log(
          `No compatibility function found for ${sheetType}. Assuming incompatible.`,
        );
        return {
          success: false,
          message: `No compatibility function found for ${sheetType}. Assuming incompatible.`,
          imported: false,
        };
      }
    }

    var newSpreadsheetInfo = shared.findSheetTypeID(
      newSheetID,
      sheetName,
      searchName,
      newIDSValues,
    );
    if (!newSpreadsheetInfo || !newSpreadsheetInfo.importStatus) {
      console.log(`Can not find import cell in the new IDS sheet.`);
      return {
        success: false,
        message: `Can not find import cell in the new IDS sheet.`,
        imported: false,
      };
    }

    var importStatusValue = newSpreadsheetInfo.importStatus.value;
    var isImported = importStatusValue === "✅";

    if (isImported) {
      return {
        success: true,
        message: "Data is already imported.",
        imported: true,
        versionDifference: versionDifference,
      };
    } else {
      return {
        success: true,
        message: `Data is not imported yet. Ready for import.`,
        imported: false,
        versionDifference: versionDifference,
      };
    }
  } catch (error) {
    console.log(
      `Error checking import status and compatibility: ${error.message}`,
    );
    return {
      success: false,
      message: `Error checking import status and compatibility: ${error.message}`,
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
      authorizationUrl: "",
      message: "Token retrieved successfully",
    };
  } catch (error) {
    console.error("Error getting OAuth token:", error);

    return {
      success: false,
      token: null,
      authorizationUrl: getScopeAuthorizationUrl(),
      message: error.toString(),
    };
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
    console.error("Error getting authorization URL:", authInfoError);
    return "";
  }
}

function checkScopePermissions() {
  try {
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
    return true;
  } catch (error) {
    console.error("Scope permission check failed:", error);
    return false;
  }
}

function checkSheetAccess(sheetID, userEmail) {
  try {
    if (!sheetID || !userEmail) {
      return {
        success: false,
        message: "Missing sheetID or userEmail parameter",
        accessible: false,
        owned: false,
      };
    }

    try {
      const file = CacheManager.getFile(sheetID);

      const owners = file.owners || [];
      const isOwner = owners.some(
        (owner) =>
          owner.emailAddress &&
          owner.emailAddress.toLowerCase() === userEmail.toLowerCase(),
      );

      return {
        success: true,
        message: `Sheet access verified`,
        accessible: true,
        owned: isOwner,
        sheetID: sheetID,
        name: file.name,
      };
    } catch (error) {
      console.log(`Sheet access denied for ${sheetID}: ${error.toString()}`);

      return {
        success: true,
        message: `Sheet access denied`,
        accessible: false,
        owned: false,
        sheetID: sheetID,
      };
    }
  } catch (error) {
    console.error(`Error checking sheet access: ${error.toString()}`);
    return {
      success: false,
      message: `Error checking sheet access: ${error.toString()}`,
      accessible: false,
      owned: false,
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
      "Guardians",
      "Player & Stuff",
    ];

    copyMode = copyMode || "all";
    console.log(
      `Getting template and old sheet IDs for IDS Master: ${idMasterID}, mode: ${copyMode}`,
    );

    var idsMasterData = fetchIdsMasterData(idMasterID);
    if (!idsMasterData.success) {
      console.log(`Error fetching IDS Master data: ${idsMasterData.message}`);
      return {
        success: false,
        message: `Error fetching IDS Master data: ${idsMasterData.message}`,
        collection: idsMasterData.collection || false,
      };
    }

    var templateInfo = [];
    var sheetIds = [idMasterID];

    for (var i = 0; i < sheetTypes.length; i++) {
      var sheetType = sheetTypes[i];
      try {
        var templateResult = getTemplateInfo(
          idsMasterData,
          sheetType,
          copyMode,
        );

        if (templateResult && templateResult.success) {
          if (templateResult.versionFiltered) {
            console.log(`Skipping ${sheetType} - version filtering applied`);
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
        console.log(
          `Error processing template for ${sheetType}: ${templateError.toString()}`,
        );
      }
    }

    return {
      success: true,
      sheetIds: sheetIds,
      templateInfo: templateInfo,
      message: `Found ${templateInfo.length} templates and ${sheetIds.length} old sheets to check`,
    };
  } catch (error) {
    console.log(
      `Error getting template and old sheet IDs: ${error.toString()}`,
    );
    return {
      success: false,
      message: `Error getting template and old sheet IDs: ${error.message}`,
    };
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

    var templateVersion = spreadsheetInfo.version.value;
    var oldVersion = spreadsheetInfo.oldVersion.value;

    if (copyMode === "update") {
      if (!templateVersion || !oldVersion) {
        console.log(
          `Version information missing for ${sheetType} - template: ${templateVersion}, old: ${oldVersion}`,
        );
        return {
          success: true,
          versionFiltered: true,
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
      sheetType: sheetType,
      message: `Successfully got template info for ${sheetType}`,
    };
  } catch (error) {
    console.error(
      `Error getting template info for ${sheetType}: ${error.toString()}`,
    );
    return { success: false, message: `${error.toString()}` };
  }
}

function checkTemplateAccess(templateID) {
  try {
    if (!templateID) {
      return {
        success: false,
        message: "Missing templateID parameter",
        accessible: false,
      };
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
      console.log(
        `Template access denied for ${templateID}: ${error.toString()}`,
      );

      return {
        success: true,
        message: `Template access denied`,
        accessible: false,
        templateID: templateID,
      };
    }
  } catch (error) {
    console.error(`Error checking template access: ${error.toString()}`);
    return {
      success: false,
      message: `Error checking template access: ${error.toString()}`,
      accessible: false,
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
      `Could not find sheet ID for ${sheetType}. Please check that ${sheetType} ID is set in the IDS Master sheet.`,
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
        collection: true,
      };
    }

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
      idMasterID: idMasterID,
    };
  } catch (error) {
    console.error(`Error fetching IDS Master data: ${error.toString()}`);
    return { success: false, message: `${error.toString()}` };
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

    var templateVersion = spreadsheetInfo.version.value;
    var oldVersion = spreadsheetInfo.oldVersion.value;

    if (copyMode === "update") {
      if (!templateVersion || !oldVersion) {
        console.log(
          `Version information missing for ${sheetType} - template: ${templateVersion}, old: ${oldVersion}`,
        );
        return {
          success: true,
          versionFiltered: true,
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
        oldFileId: oldSheetID,
        idMasterFileId: idMasterID,
        sheetType: sheetType,
      };
    } catch (error) {
      console.log(
        `Error retrieving template file information: ${error.toString()}`,
      );
      console.log(`Template ID: ${templateID}, Sheet Type: ${sheetType}`);

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
    console.error(
      `Error processing template access for ${sheetType}: ${error.toString()}`,
    );
    return { success: false, message: `${error.toString()}` };
  }
}

function checkFileTemplateAccess(idMasterID, sheetType) {
  var idsMasterData = fetchIdsMasterData(idMasterID);
  if (!idsMasterData.success) {
    return idsMasterData;
  }

  return processTemplateAccess(idsMasterData, sheetType, "all");
}

function copyFileTemplate(
  templateID,
  sheetType,
  templateVersion,
  parentFolderID,
) {
  try {
    var resolvedTemplateVersion = templateVersion || "";
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
      console.log(`Error copying ${sheetType} template: no file returned`);
      return {
        success: false,
        message: `Error copying ${sheetType} template: no file returned`,
        copyUrl: templateCopyUrl,
      };
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
    console.error(`Error copying ${sheetType} template: ${error.toString()}`);
    return {
      success: false,
      message: `${error.toString()}`,
      copyUrl: templateID
        ? "https://docs.google.com/spreadsheets/d/" + templateID + "/copy"
        : "",
    };
  }
}

function moveGetStartedFileToFolder(fileId, parentFolderID) {
  try {
    if (!fileId) {
      return { success: false, message: "Missing fileId parameter." };
    }
    if (!parentFolderID) {
      return { success: false, message: "Missing parentFolderID parameter." };
    }

    var file = CacheManager.getFile(fileId);
    if (!file) {
      return {
        success: false,
        message: `File not found for ID: ${fileId}`,
      };
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
    console.error(
      `Error moving Get Started file (${fileId}): ${error.toString()}`,
    );
    return {
      success: false,
      message: `Error moving file: ${error.toString()}`,
    };
  }
}

function checkNewSheetReference(newSheetID, sheetType) {
  try {
    if (!newSheetID) {
      console.log(`Missing newSheetID parameter.`);
      return { success: false, message: "Missing newSheetID parameter." };
    }

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
      return {
        success: false,
        message: `Could not read IDS sheet data from IDS Master`,
        succeededTasks: [],
        failedTasks: [],
      };
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
        failedTasks.push({
          sheetType: sheetType,
          success: false,
          message: `Could not find old sheet information for ${sheetType}`,
        });
        continue;
      }

      var oldSheetID = shared.extractSheetId(sheetTypeInfo.id);
      if (!oldSheetID) {
        failedTasks.push({
          sheetType: sheetType,
          success: false,
          message: `Could not extract old sheet ID for ${sheetType}`,
        });
        continue;
      }

      var oldVersion = sheetTypeInfo.oldVersion.value;
      var templateVersion = sheetTypeInfo.version.value;

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
          failedTasks.push({
            sheetType: sheetType,
            success: false,
            message: `Old version (${oldVersion}) is newer than template version (${templateVersion})`,
          });
          continue;
        }

        var sheetTypeFunction = sheetVars(sheetType);
        if (sheetTypeFunction) {
          versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
          if (!versionDifference) {
            failedTasks.push({
              sheetType: sheetType,
              success: false,
              message: `Old version of ${sheetType} is incompatible for import`,
            });
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
    console.error(`Error preparing import data:`, error);
    return {
      success: false,
      message: `Error preparing import data: ${error.toString()}`,
      succeededTasks: [],
      failedTasks: [],
    };
  }
}

function deleteOldSheet(sheetID) {
  try {
    console.log(`Attempting to delete sheet with ID: ${sheetID}`);

    var fileInfo;
    try {
      fileInfo = CacheManager.getFile(sheetID);
    } catch (error) {
      console.log(
        `Sheet ${sheetID} not found or already deleted: ${error.toString()}`,
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

    Drive.Files.trash(sheetID);

    console.log(`Successfully deleted sheet: ${fileInfo.name} (${sheetID})`);
    return {
      success: true,
      message: `Successfully deleted sheet: "${fileInfo.name}"`,
    };
  } catch (error) {
    console.log(`Error deleting sheet ${sheetID}: ${error.toString()}`);
    return {
      success: false,
      message: `Error deleting sheet: ${error.toString()}`,
    };
  }
}

function getTemplateIdForSingleSheet(sheetID, sheetType) {
  try {
    if (!sheetID) {
      console.log(`Missing sheetID parameter.`);
      return { success: false, message: "Missing sheetID parameter." };
    }
    if (!sheetType) {
      console.log(`Missing sheetType parameter.`);
      return { success: false, message: "Missing sheetType parameter." };
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
      console.log(`Could not find sheet template for ${sheetType}`);
      return {
        success: false,
        message: `Could not find sheet template for ${sheetType}`,
      };
    }
    return {
      success: true,
      templateID: spreadsheetInfo.templateID,
      templateVersion: spreadsheetInfo.templateVersion,
      message: `Successfully got template ID for ${sheetType}`,
    };
  } catch (error) {
    console.error(
      `Error getting template ID for single sheet: ${error.toString()}`,
    );
    return { success: false, message: `${error.toString()}` };
  }
}

function checkExportCompatibility(oldSheetID, sheetType) {
  try {
    if (!oldSheetID) {
      return { success: false, message: "Missing oldSheetID parameter." };
    }
    if (!sheetType) {
      return { success: false, message: "Missing sheetType parameter." };
    }

    var oldSpreadsheet = spreadsheets(
      `${sheetType} oldSpreadsheet`,
      oldSheetID,
    );
    if (!oldSpreadsheet) {
      return {
        success: false,
        message: `Could not access old ${sheetType} spreadsheet with ID: ${oldSheetID}`,
      };
    }

    var oldHomePageSheet = SheetsAPI.getSheetByName(
      oldSpreadsheet,
      "Home Page",
    );
    if (!oldHomePageSheet) {
      return {
        success: false,
        message: `Home Page sheet not found in old ${sheetType} spreadsheet`,
      };
    }

    var oldHomePageData = SheetsAPI.batchGetValues(oldSheetID, ["Home Page"]);
    if (!oldHomePageData || oldHomePageData.length === 0) {
      return {
        success: false,
        message: `Could not read Home Page data from old ${sheetType} spreadsheet`,
      };
    }

    var oldHomePageValues = oldHomePageData[0].values;
    var oldVersionInfo;
    var oldVersion;

    oldVersionInfo = shared.findSheetVersion(
      oldSheetID,
      "Home Page",
      sheetType,
      oldHomePageValues,
    );
    if (!oldVersionInfo || !oldVersionInfo.currentVersion) {
      return {
        success: false,
        message: `Current Version not found in old ${sheetType} spreadsheet.`,
      };
    }
    oldVersion = oldVersionInfo.currentVersion;

    var sheetTypeFunction = sheetVars(sheetType);
    if (sheetTypeFunction) {
      var versionDifference = sheetTypeFunction.isCompatibleVersion(oldVersion);
      if (!versionDifference) {
        return {
          success: false,
          message: `Old version of ${sheetType} is incompatible for export (${oldVersion})`,
        };
      }
      return {
        success: true,
        message: `Old ${sheetType} version (${oldVersion}) is compatible for export`,
        oldVersion: oldVersion,
        versionDifference: versionDifference,
      };
    } else {
      return {
        success: false,
        message: `No compatibility function found for ${sheetType}. Cannot verify export compatibility.`,
      };
    }
  } catch (error) {
    console.error(`Error checking export compatibility: ${error.toString()}`);
    return {
      success: false,
      message: `Error checking export compatibility: ${error.toString()}`,
    };
  }
}

function updateSheetID(spreadsheetID, sheetID, sheetType) {
  try {
    if (!spreadsheetID) {
      return { success: false, message: "Missing spreadsheetID parameter." };
    }
    if (!sheetID) {
      return { success: false, message: "Missing sheetID parameter." };
    }
    if (!sheetType) {
      return { success: false, message: "Missing sheetType parameter." };
    }
    var spreadsheet = spreadsheets(`${sheetType} spreadsheet`, spreadsheetID);
    if (!spreadsheet) {
      return {
        success: false,
        message: `Could not access spreadsheet with ID: ${spreadsheetID}`,
      };
    }
    var idSheet = SheetsAPI.getSheetByName(spreadsheet, "IDS");
    if (!idSheet) {
      return {
        success: false,
        message: `IDS sheet not found in spreadsheet with ID: ${spreadsheetID}`,
      };
    }
    var idValues = SheetsAPI.batchGetValues(spreadsheetID, ["IDS"]);
    if (!idValues || idValues.length === 0) {
      return {
        success: false,
        message: `Could not read IDS sheet data from spreadsheet with ID: ${spreadsheetID}`,
      };
    }
    var values = idValues[0].values;
    var sheetTypeInfo = shared.findSheetTypeID(
      spreadsheetID,
      "IDS",
      "IDS Master",
      values,
    );
    if (!sheetTypeInfo || !sheetTypeInfo.cell) {
      return {
        success: false,
        message: `Could not find IDS Master entry in IDS sheet.`,
      };
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
    console.error(`Error updating sheet ID: ${error.toString()}`);
    return {
      success: false,
      message: `Error updating sheet ID: ${error.toString()}`,
    };
  }
}

function getOrCreateGetStartedFolder() {
  try {
    // Search for a folder named "The Tower" in the user's Drive
    var query =
      'name="The Tower" and mimeType="application/vnd.google-apps.folder" and trashed=false';
    var folderList = Drive.Files.list({
      q: query,
      spaces: "drive",
      fields: "files(id, name)",
      pageSize: 1,
    });

    if (folderList.files && folderList.files.length > 0) {
      // Folder exists, use it
      var folder = folderList.files[0];
      console.log(`Found existing "The Tower" folder: ${folder.id}`);
      return {
        success: true,
        id: folder.id,
        name: folder.name,
        url: `https://drive.google.com/drive/folders/${folder.id}`,
      };
    }

    // Folder doesn't exist, create it
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
    console.error(
      `Error getting or creating The Tower folder: ${error.toString()}`,
    );
    return {
      success: false,
      id: "",
      name: "",
      url: "",
      error: `Error locating or creating The Tower folder: ${error.toString()}`,
    };
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
        console.log(`Could not fetch Home Page from ${sheetType}`);
        return { success: false, message: "Could not fetch Home Page" };
      }

      var idsData = idsResult[0].values;
      var ownSheetInfo = shared.findSheetTypeID(
        sheetID,
        "Home Page",
        "Your ID:",
        idsData,
      );

      if (!ownSheetInfo || !ownSheetInfo.cell || !ownSheetInfo.cell.range) {
        return {
          success: false,
          message: "Could not find 'Your ID:' in IDS Collection",
        };
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
        console.log(
          `Could not update file name with version info: ${error.toString()}`,
        );
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
        console.log(
          `Could not update file name with version info: ${error.toString()}`,
        );
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
      console.log(`Could not fetch IDS sheet from ${sheetType}`);
      return { success: false, message: "Could not fetch IDS sheet" };
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
      console.log(
        `Could not update file name with version info: ${error.toString()}`,
      );
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
    console.error(`Error updating sheet IDs: ${error.toString()}`);
    return { success: false, message: error.toString() };
  }
}
