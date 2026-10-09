const presetUtils = {
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
};
