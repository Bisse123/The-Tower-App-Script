const presetUtils = {
  templatePresetNames: ["Farming", "Tourney"],

  presetUnlocks: {
    presetsLab: 40,
    globalPresetsLab: 252,
    maxPresets: 5,
    cardsLabPresets: 5,
    labPresets: 2,
    types: {
      cards: { labLevel: 1, vaultUpgrade: null },
      workshop: { labLevel: 2, vaultUpgrade: "2005" },
      bots: { labLevel: 3, vaultUpgrade: "1805" },
      modules: { labLevel: 4, vaultUpgrade: "1640" },
      guardians: { labLevel: 5, vaultUpgrade: "2205" },
      global: { labLevel: 1, vaultUpgrade: "1030" },
    },
  },

  /**
   * Counts the presets a player has unlocked for each preset type. The Presets lab's first level
   * unlocks five card presets and its levels 2 to 5 unlock two presets each for Workshop, Bots,
   * Modules and Guardians; the Global Presets lab unlocks two global presets; a type's vault
   * upgrade raises it to five. A type without presets still holds its one live build.
   * @param {number[]} researchLevel Lab levels by lab index, as stored in the save file.
   * @param {Object<string, number>} vaultLevels Vault upgrade levels keyed by upgrade index, as
   *   stored in the save file.
   * @returns {{cards: number, workshop: number, bots: number, modules: number, guardians: number,
   *   global: number}} Presets to import per type: 1 to 5, and 0 to 5 for global.
   */
  unlockedPresetCounts: function (researchLevel, vaultLevels) {
    var unlocks = presetUtils.presetUnlocks;
    var labs = researchLevel || [];
    var vault = vaultLevels || {};
    var presetsLabLevel = Number(labs[unlocks.presetsLab]) || 0;
    var globalLabLevel = Number(labs[unlocks.globalPresetsLab]) || 0;
    var counts = {};
    Object.keys(unlocks.types).forEach(function (type) {
      var rule = unlocks.types[type];
      var isGlobal = type === "global";
      var labUnlocked = isGlobal
        ? globalLabLevel >= rule.labLevel
        : presetsLabLevel >= rule.labLevel;
      var vaultUnlocked = rule.vaultUpgrade !== null && (Number(vault[rule.vaultUpgrade]) || 0) > 0;
      var labCount = type === "cards" ? unlocks.cardsLabPresets : unlocks.labPresets;
      if (vaultUnlocked) {
        counts[type] = unlocks.maxPresets;
      } else if (labUnlocked) {
        counts[type] = labCount;
      } else {
        counts[type] = isGlobal ? 0 : 1;
      }
    });
    return counts;
  },

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
};
