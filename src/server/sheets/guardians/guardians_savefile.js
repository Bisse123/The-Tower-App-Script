const guardianHeaders = {
  activePreset: "currentGuardianPreset",
  guardianChipSlot: "guardianChipSlot",
  guardianChipUnlocked: "guardianChipUnlocked",
  guardianChipLevel: "guardianChipLevel",
  guardianPresets: "guardianPresets",
}

const guardiansSaveFile = {

  /**
   * Parses Guardians data out of a decoded save file.
   * @param {Object} data
   * @param {number} presetCount How many of the stored presets the player has unlocked; only
   *   those are read.
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseGuardiansData: function (data, presetCount) {
    try {

      const targetGuardians = {
        0: { name: "Bounty", upgrades: ["Multiplier", "Cooldown", "Targets"] },
        2: { name: "Attack", upgrades: ["Percentage", "Cooldown", "Targets"], alwaysUnlocked: true },
        5: { name: "Ally", upgrades: ["Recovery Amount", "Max Recovery", "Cooldown"], alwaysUnlocked: true },
        6: { name: "Fetch", upgrades: ["Cooldown", "Find Chance", "Double Find Chance"] },
        7: { name: "Summon", upgrades: ["Cooldown", "Duration", "Cash Bonus"] },
        8: { name: "Scout", upgrades: ["Cooldown", "Range Bonus", "Duration"] },
      };
      const guardianIndices = Object.keys(targetGuardians)
        .map(Number)
        .sort(function (a, b) {
          return a - b;
        });

      const guardianUnlockedData = data.guardianChipUnlocked || [];
      const guardianPresetsData = (data.guardianPresets || []).slice(0, presetCount);

      var presets = guardianPresetsData.length
        ? guardianPresetsData
        : [
            {
              presetName: "Farming",
              chipSlot: data.guardianChipSlot || [],
              chipLevel: data.guardianChipLevel || [],
            },
          ];

      /**
       * Overlays the live chip levels and slots onto the preset the player
       * currently has selected, whose stored copy goes stale while they edit.
       * @param {Array<Object>} presetList
       * @param {number} presetIndex
       * @param {Array<number>} liveLevels
       * @param {Array<number>} liveSlots
       * @returns {Array<Object>} A copy; the input is not modified.
       */
      function withLiveValues(presetList, presetIndex, liveLevels, liveSlots) {
        var active = presetList[presetIndex];
        if (!active) return presetList;
        var merged = {};
        Object.keys(active).forEach(function (key) {
          merged[key] = active[key];
        });
        if (liveLevels && liveLevels.length) merged.chipLevel = liveLevels;
        if (liveSlots && liveSlots.length) merged.chipSlot = liveSlots;
        var copy = presetList.slice();
        copy[presetIndex] = merged;
        return copy;
      }

      presets = withLiveValues(
        presets,
        data.activePreset || 0,
        data.guardianChipLevel || [],
        data.guardianChipSlot || [],
      );

      var presetNames = [];
      var oldGuardians = {
        presetNames: presetNames,
        data: {},
      };

      presets.forEach(function (preset, index) {
        const chipLevel = preset.chipLevel || [];
        if (!chipLevel.some(function (level) {
          return level > 0;
        })) {
          return;
        }
      
        var presetName = preset.presetName;
        presetNames.push(presetName);

        const chipSlot = preset.chipSlot || [];

        guardianIndices.forEach(function (guardianIndex) {
          const { name, upgrades, alwaysUnlocked } = targetGuardians[guardianIndex];

          var chipLevels = {};
          upgrades.forEach(function (attr, j) {
            const level = chipLevel[guardianIndex * upgrades.length + j];
            chipLevels[attr] = level ? String(level).padStart(2, "0") : "00";
          });

          if (!oldGuardians.data[name]) {
            oldGuardians.data[name] = {
              unlocked: alwaysUnlocked
                ? null
                : guardianUnlockedData[guardianIndex] || false,
              presets: {},
            };
          }
          oldGuardians.data[name].presets[presetName] = {
            props: chipLevels,
            equipped: chipSlot.includes(guardianIndex),
          };
        });
      });

      oldGuardians.presetNames = presetUtils.resolvePresetOrder(
        presetNames,
        presetUtils.templatePresetNames,
      ).order;

      var targetGuardiansByName = {};
      var guardianOrder = [];
      guardianIndices.forEach(function (guardianIndex) {
        const { name, upgrades, alwaysUnlocked } = targetGuardians[guardianIndex];
        targetGuardiansByName[name] = alwaysUnlocked
          ? { upgrades: upgrades, alwaysUnlocked: alwaysUnlocked }
          : { upgrades: upgrades };
        guardianOrder.push(name);
      });

      return {
        success: true,
        oldGuardians: oldGuardians,
        targetGuardians: targetGuardiansByName,
        guardianOrder: guardianOrder,
      };
    } catch (error) {
      var errorReport = errors.report("guardiansSaveFile.parseGuardiansData", error, {
        data: data,
        oldGuardians: oldGuardians,
      });
      return errors.fail(errorReport);
    }
  },
};
