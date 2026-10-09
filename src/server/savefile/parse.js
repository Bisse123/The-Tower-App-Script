/**
 * Client-callable. Ungzips and decodes a playerInfo.dat, then parses every
 * category independently so one failure does not cost the others.
 * @param {number[]} byteArray The raw file bytes.
 * @returns {{success: true, parsed: Object, order: string[], failedCategories: Array<Object>}}
 *   A failure envelope when the file is not a readable save.
 */
function parseSaveFileBytes(byteArray) {
  const uint8 = Uint8Array.from(byteArray);

  var decompressedBlob;
  try {
    decompressedBlob = Utilities.ungzip(
      Utilities.newBlob(Array.from(uint8), "application/x-gzip"),
    );
  } catch (error) {
    var errorReport = errors.report("parseSaveFileBytes", error, {
      byteLength: uint8.length,
    }, errors.CODES.INVALID_FILE);
    return errors.fail(errorReport);
  }

  const bytes = blobToUint8Array_(decompressedBlob);

  var data;
  try {
    data = parseNRBF(bytes);
  } catch (error) {
    var errorReport = errors.report("parseSaveFileBytes", error, {
      byteLength: bytes.length,
    });
    return errors.fail(errorReport);
  }

  /**
   * Pulls the save-file keys a sheet needs out of the decoded data.
   * @param {Object} headers Sheet key to save-file key.
   * @returns {Object} Sheet key to value, null where the key is absent.
   */
  function extractDataByHeaders(headers) {
    var values = {};
    for (const sheetKey of Object.keys(headers)) {
      const saveKey = headers[sheetKey];
      values[sheetKey] = data.hasOwnProperty(saveKey) ? data[saveKey] : null;
    }
    return values;
  }

  var vaultUpgradeLevels = {};
  Object.keys(data.vault || {}).forEach(function (key) {
    if (key.includes("UpgradesLevel")) vaultUpgradeLevels = data.vault[key] || {};
  });
  var presetCounts = presetUtils.unlockedPresetCounts(data.researchLevel, vaultUpgradeLevels);

  var labValues = extractDataByHeaders(labHeaders);
  var laboratoryData = labSaveFile.parseLabData(labValues);

  var workshopValues = extractDataByHeaders(workshopHeaders);
  var workshopData = workshopSaveFile.parseWorkshopData(workshopValues, presetCounts.workshop);

  var ultimateWeaponValues = extractDataByHeaders(ultimateWeaponHeaders);
  var ultimateWeaponData = ultimateSaveFile.parseUltimateWeaponData(ultimateWeaponValues);

  var themesAndRelicsValues = extractDataByHeaders(themesAndRelicsHeaders);
  var themesAndRelicsData =
    themesAndRelicsSaveFile.parseThemesAndRelicsData(themesAndRelicsValues);

  var botValues = extractDataByHeaders(botHeaders);
  var botData = botsSaveFile.parseBotsData(botValues, presetCounts.bots);

  var vaultValues = extractDataByHeaders(vaultHeaders);
  var vaultData = vaultSaveFile.parseVaultData(vaultValues);

  var cardsValues = extractDataByHeaders(cardsHeaders);
  var cardsData = cardsSaveFile.parseCardsData(cardsValues, presetCounts.cards);

  var moduleValues = extractDataByHeaders(moduleHeaders);
  var moduleData = modulesSaveFile.parseModulesData(moduleValues, presetCounts.modules);

  var guardianValues = extractDataByHeaders(guardianHeaders);
  var guardianData = guardiansSaveFile.parseGuardiansData(guardianValues, presetCounts.guardians);

  var playerStuffValues = extractDataByHeaders(PlayerStuffHeaders);
  var playerStuffdata = playerStuffSaveFile.parsePlayerStuffData(playerStuffValues);

  var masterValues = extractDataByHeaders(MasterHeaders);
  var masterData = masterSaveFile.parseMasterData(masterValues, presetCounts);

  const parsed = {
    "Laboratory": laboratoryData,
    "Workshop": workshopData,
    "Ultimate Weapon": ultimateWeaponData,
    "Themes, Songs & Relics": themesAndRelicsData,
    "Bots": botData,
    "Vault": vaultData,
    "Cards": cardsData,
    "Modules": moduleData,
    "Guardians": guardianData,
    "Player & Stuff": playerStuffdata,
    "IDS Master": masterData,
  };

  var order = Object.keys(parsed);
  var successfulParsed = {};
  var failedCategories = [];
  order.forEach(function (type) {
    var categoryResult = parsed[type];
    if (categoryResult && categoryResult.success === false) {
      var failure = errors.propagate(
        "parseSaveFileBytes",
        categoryResult,
        `${type} could not be read from your save file.`,
      );
      var failedCategory = {
        type: type,
        success: false,
        code: failure.code,
        expected: failure.expected,
        message: failure.message,
        reference: failure.reference,
        detail: failure.detail,
        trace: failure.trace,
      };
      if (failure.note) failedCategory.note = failure.note;
      if (failure.data) failedCategory.data = failure.data;
      if (failure.stack) failedCategory.stack = failure.stack;
      failedCategories.push(failedCategory);
      return;
    }
    successfulParsed[type] = categoryResult;
  });

  return {
    success: true,
    parsed: successfulParsed,
    order: order,
    failedCategories: failedCategories,
  };
}

/**
 * Blob to unsigned bytes.
 * @param {Blob} blob
 * @returns {Uint8Array}
 */
function blobToUint8Array_(blob) {
  const signedBytes = blob.getBytes();
  const out = new Uint8Array(signedBytes.length);
  for (let i = 0; i < signedBytes.length; i++) {
    out[i] = signedBytes[i] & 0xff;
  }
  return out;
}
