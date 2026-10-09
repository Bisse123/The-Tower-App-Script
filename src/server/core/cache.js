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
