const workshopHeaders = {
  activePreset: "currentWorkshopPreset",
  presetNames: "workshopPresetName",
  upgradeAttackLevels: "upgradeWorkshopLevel",
  upgradeDefenseLevels: "upgradeWorkshopDefenseLevel",
  upgradeUtilityLevels: "upgradeWorkshopUtilityLevel",
  presetUpgradeAttackLevels: "presetUpgradeWorkshopLevel",
  presetUpgradeDefenseLevels: "presetUpgradeWorkshopDefenseLevel",
  presetUpgradeUtilityLevels: "presetUpgradeWorkshopUtilityLevel",
  enhancementAttackLevels: "enhancementLevel",
  enhancementDefenseLevels: "enhancementDefenseLevel",
  enhancementUtilityLevels: "enhancementUtilityLevel",
  presetEnhancementAttackLevels: "presetEnhancementLevel",
  presetEnhancementDefenseLevels: "presetEnhancementDefenseLevel",
  presetEnhancementUtilityLevels: "presetEnhancementUtilityLevel",
  upgradeAttackUnlocked: "upgradeTierUnlocked",
  upgradeDefenseUnlocked: "upgradeDefenseTierUnlocked",
  upgradeUtilityUnlocked: "upgradeUtilityTierUnlocked",
  presetUpgradeAttackUnlocked: "presetUpgradeTierUnlocked",
  presetUpgradeDefenseUnlocked: "presetUpgradeDefenseTierUnlocked",
  presetUpgradeUtilityUnlocked: "presetUpgradeUtilityTierUnlocked",
}

const workshopSaveFile = {

  /**
   * Parses Workshop data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseWorkshopData: function (data) {
    try {
      const attackUpgradeNamesByIndex = {
        0: "Damage",
        1: "Attack Speed",
        2: "Critical Chance",
        3: "Critical Factor",
        4: "Range",
        5: "Damage / Meter",
        6: "Multishot Chance",
        7: "Multishot Targets",
        8: "Rapid Fire Chance",
        9: "Rapid Fire Duration",
        10: "Bounce Shot Chance",
        11: "Bounce Shot Targets",
        12: "Bounce Shot Range",
        13: "Super Critical Chance",
        14: "Super Critical Mult",
        15: "Rend Armor Chance",
        16: "Rend Armor Mult",
      };
      const defenseUpgradeNamesByIndex = {
        0: "Health",
        1: "Health Regen",
        2: "Defense %",
        3: "Defense Absolute",
        4: "Thorn Damage",
        5: "Lifesteal",
        6: "Knockback Chance",
        7: "Knockback Force",
        8: "Orb Speed",
        9: "Orbs",
        10: "Shockwave Size",
        11: "Shockwave Frequency",
        12: "Land Mine Chance",
        13: "Land Mine Damage",
        14: "Land Mine Radius",
        15: "Death Defy",
        16: "Wall Health",
        17: "Wall Rebuild",
      };
      const utilityUpgradeNamesByIndex = {
        0: "Cash Bonus",
        1: "Cash / Wave",
        2: "Coin / Kill Bonus",
        3: "Coin / Wave",
        4: "Free Attack Upgrade",
        5: "Free Defense Upgrade",
        6: "Free Utility Upgrade",
        7: "Interest / Wave",
        8: "Recovery Amount",
        9: "Max Amount",
        10: "Package Chance",
        11: "Enemy Attack Level Skip",
        12: "Enemy Health Level Skip",
      };
      const attackEnhancementNamesByIndex = {
        0: "Damage +",
        1: "Rend Armor Mult +",
        2: "Critical Factor +",
        3: "Damage / Meter +",
        4: "Super Crit Multi +",
        5: "Attack Speed +",
      };
      const defenseEnhancementNamesByIndex = {
        0: "Health +",
        1: "Health Regen +",
        2: "Defense Absolute +",
        3: "Land Mine Damage +",
        4: "Wall Health +",
        5: "Orb Size +",
      };
      const utilityEnhancementNamesByIndex = {
        0: "Cash Bonus +",
        1: "Coin Bonus +",
        2: "Cells / Kill Bonus +",
        3: "Free Upgrades +",
        4: "Recovery Package +",
        5: "Enemy Level Skips +",
      };
      const attackUpgradeUnlockedNamesByIndex = {
        0: "Range",
        1: "Multishot Chance",
        2: "Rapid Fire Chance",
        3: "Bounce Shot Chance",
        4: "Super Critical Chance",
        5: "Rend Armor Chance",
      };
      const defenseUpgradeUnlockedNamesByIndex = {
        0: "Defense %",
        1: "Thorn Damage",
        2: "Lifesteal",
        3: "Knockback Chance",
        4: "Orb Speed",
        5: "Shockwave Size",
        6: "Land Mine Chance",
        7: "Death Defy",
        8: "Wall Health",
      };
      const utilityUpgradeUnlockedNamesByIndex = {
        0: "Cash Bonus",
        1: "Coin / Kill Bonus",
        2: "Free Attack Upgrade",
        3: "Interest / Wave",
        4: "Recovery Amount",
        5: "Enemy Attack Level Skip",
      };

      /**
       * Turns an index-keyed name map into a dense array.
       * @param {Object} namesByIndex
       * @returns {Array<string>}
       */
      function namesByIndexToArray(namesByIndex) {
        var array = [];
        Object.keys(namesByIndex).forEach(function (index) {
          array[Number(index)] = namesByIndex[index];
        });
        return array;
      }

      const attackUpgradeIndices = namesByIndexToArray(attackUpgradeNamesByIndex);
      const defenseUpgradeIndices = namesByIndexToArray(defenseUpgradeNamesByIndex);
      const utilityUpgradeIndices = namesByIndexToArray(utilityUpgradeNamesByIndex);
      const attackEnhancementIndices = namesByIndexToArray(attackEnhancementNamesByIndex);
      const defenseEnhancementIndices = namesByIndexToArray(defenseEnhancementNamesByIndex);
      const utilityEnhancementIndices = namesByIndexToArray(utilityEnhancementNamesByIndex);
      const attackUpgradeUnlockedIndices = namesByIndexToArray(attackUpgradeUnlockedNamesByIndex);
      const defenseUpgradeUnlockedIndices = namesByIndexToArray(defenseUpgradeUnlockedNamesByIndex);
      const utilityUpgradeUnlockedIndices = namesByIndexToArray(utilityUpgradeUnlockedNamesByIndex);

      const presetOrder = presetUtils.resolvePresetOrder(
        data.presetNames || [],
        presetUtils.templatePresetNames,
      );
      var presetNames = presetOrder.order;
      const presetIndices = presetOrder.indices;

      const activePreset = data.activePreset || 0;
      const attackUpgradeData = data.upgradeAttackLevels || [];
      const defenseUpgradeData = data.upgradeDefenseLevels || [];
      const utilityUpgradeData = data.upgradeUtilityLevels || [];

      const attackEnhancementData = data.enhancementAttackLevels || [];
      const defenseEnhancementData = data.enhancementDefenseLevels || [];
      const utilityEnhancementData = data.enhancementUtilityLevels || [];

      const attackUpgradeMax = attackUpgradeData.length;
      const defenseUpgradeMax = defenseUpgradeData.length;
      const utilityUpgradeMax = utilityUpgradeData.length;
      const attackEnhancementMax = attackEnhancementData.length;
      const defenseEnhancementMax = defenseEnhancementData.length;
      const utilityEnhancementMax = utilityEnhancementData.length;

      var attackPresetUpgradeData = data.presetUpgradeAttackLevels || [];
      var defensePresetUpgradeData = data.presetUpgradeDefenseLevels || [];
      var utilityPresetUpgradeData = data.presetUpgradeUtilityLevels || [];

      var attackPresetEnhancementData = data.presetEnhancementAttackLevels || [];
      var defensePresetEnhancementData = data.presetEnhancementDefenseLevels || [];
      var utilityPresetEnhancementData = data.presetEnhancementUtilityLevels || [];

      const attackUpgradeUnlocked = data.upgradeAttackUnlocked || [];
      const defenseUpgradeUnlocked = data.upgradeDefenseUnlocked || [];
      const utilityUpgradeUnlocked = data.upgradeUtilityUnlocked || [];

      var attackPresetUpgradeUnlocked = data.presetUpgradeAttackUnlocked || [];
      var defensePresetUpgradeUnlocked = data.presetUpgradeDefenseUnlocked || [];
      var utilityPresetUpgradeUnlocked = data.presetUpgradeUtilityUnlocked || [];

      const hasPresets =
        attackPresetUpgradeUnlocked.some((unlocked) => unlocked) ||
        defensePresetUpgradeUnlocked.some((unlocked) => unlocked) ||
        utilityPresetUpgradeUnlocked.some((unlocked) => unlocked);

      if (!hasPresets) {
        attackPresetUpgradeData = attackUpgradeData;
        defensePresetUpgradeData = defenseUpgradeData;
        utilityPresetUpgradeData = utilityUpgradeData;
        attackPresetEnhancementData = attackEnhancementData;
        defensePresetEnhancementData = defenseEnhancementData;
        utilityPresetEnhancementData = utilityEnhancementData;
        attackPresetUpgradeUnlocked = attackUpgradeUnlocked;
        defensePresetUpgradeUnlocked = defenseUpgradeUnlocked;
        utilityPresetUpgradeUnlocked = utilityUpgradeUnlocked;
      }

      /**
       * Overlays a preset slice with the live values for that preset.
       * @param {Array<*>} presetData
       * @param {Array<*>} liveData
       * @param {number} presetIndex
       * @returns {Array<*>} A copy; the input is not modified.
       */
      function withLiveValues(presetData, liveData, presetIndex) {
        if (!liveData || !liveData.length) return presetData;
        var copy = (presetData || []).slice();
        var offset = presetIndex * liveData.length;
        for (var i = 0; i < liveData.length; i++) {
          copy[offset + i] = liveData[i];
        }
        return copy;
      }

      attackPresetUpgradeData = withLiveValues(attackPresetUpgradeData, attackUpgradeData, activePreset);
      defensePresetUpgradeData = withLiveValues(defensePresetUpgradeData, defenseUpgradeData, activePreset);
      utilityPresetUpgradeData = withLiveValues(utilityPresetUpgradeData, utilityUpgradeData, activePreset);
      attackPresetEnhancementData = withLiveValues(attackPresetEnhancementData, attackEnhancementData, activePreset);
      defensePresetEnhancementData = withLiveValues(defensePresetEnhancementData, defenseEnhancementData, activePreset);
      utilityPresetEnhancementData = withLiveValues(utilityPresetEnhancementData, utilityEnhancementData, activePreset);
      attackPresetUpgradeUnlocked = withLiveValues(attackPresetUpgradeUnlocked, attackUpgradeUnlocked, activePreset);
      defensePresetUpgradeUnlocked = withLiveValues(defensePresetUpgradeUnlocked, defenseUpgradeUnlocked, activePreset);
      utilityPresetUpgradeUnlocked = withLiveValues(utilityPresetUpgradeUnlocked, utilityUpgradeUnlocked, activePreset);

      var oldWorkshopLevels = {
        presetNames: presetNames,
        data: {},
      };
      var oldWorkshopPlusLevels = {
        presetNames: presetNames,
        data: {},
      };

      /**
       * Fills oldWorkshopLevels with one preset's upgrade levels and unlocks.
       * @param {string[]} upgradeIndices
       * @param {Array<*>} upgradeData
       * @param {number} presetUpgradeIndex
       * @param {string[]} upgradeUnlockIndices
       * @param {Array<*>} upgradeUnlockedData
       * @returns {void}
       */
      function populateUpgradeLevels(upgradeIndices, upgradeData, presetUpgradeIndex, upgradeUnlockIndices, upgradeUnlockedData) {
        upgradeIndices.forEach((upgradeName, upgradeIndex) => {
          if (!oldWorkshopLevels.data[upgradeName]) {
            oldWorkshopLevels.data[upgradeName] = {
              unlocked: null,
              levels: [],
            };
          }
          const unlockedIndex = upgradeUnlockIndices.findIndex((unlocked) => unlocked === upgradeName);
          if (unlockedIndex !== -1 && !oldWorkshopLevels.data[upgradeName].unlocked) {
            oldWorkshopLevels.data[upgradeName].unlocked = upgradeUnlockedData[presetUpgradeIndex + unlockedIndex];
          }
          const levelValue = upgradeData[presetUpgradeIndex + upgradeIndex];
          oldWorkshopLevels.data[upgradeName].levels.push(levelValue, null);
        })
      }

      /**
       * Fills oldWorkshopPlusLevels with one preset's enhancement levels.
       * @param {string[]} enhancementIndices
       * @param {Array<*>} enhancementData
       * @param {number} presetEnhancementIndex
       * @returns {void}
       */
      function populateEnhancementLevels(enhancementIndices, enhancementData, presetEnhancementIndex) {
        enhancementIndices.forEach((enhancementName, enhancementIndex) => {
          if (!oldWorkshopPlusLevels.data[enhancementName]) {
            oldWorkshopPlusLevels.data[enhancementName] = [];
          }
          const levelValue = enhancementData[presetEnhancementIndex + enhancementIndex];
          oldWorkshopPlusLevels.data[enhancementName].push(levelValue);
        })
      }

      presetIndices.forEach((presetIndex) => {
        const attackUpgradeIndex = presetIndex * attackUpgradeMax;
        const defenseUpgradeIndex = presetIndex * defenseUpgradeMax;
        const utilityUpgradeIndex = presetIndex * utilityUpgradeMax;

        const attackEnhancementIndex = presetIndex * attackEnhancementMax;
        const defenseEnhancementIndex = presetIndex * defenseEnhancementMax;
        const utilityEnhancementIndex = presetIndex * utilityEnhancementMax;

        populateUpgradeLevels(attackUpgradeIndices, attackPresetUpgradeData, attackUpgradeIndex, attackUpgradeUnlockedIndices, attackPresetUpgradeUnlocked);
        populateUpgradeLevels(defenseUpgradeIndices, defensePresetUpgradeData, defenseUpgradeIndex, defenseUpgradeUnlockedIndices, defensePresetUpgradeUnlocked);
        populateUpgradeLevels(utilityUpgradeIndices, utilityPresetUpgradeData, utilityUpgradeIndex, utilityUpgradeUnlockedIndices, utilityPresetUpgradeUnlocked);

        populateEnhancementLevels(attackEnhancementIndices, attackPresetEnhancementData, attackEnhancementIndex);
        populateEnhancementLevels(defenseEnhancementIndices, defensePresetEnhancementData, defenseEnhancementIndex);
        populateEnhancementLevels(utilityEnhancementIndices, utilityPresetEnhancementData, utilityEnhancementIndex);
      })

      return {
        success: true,
        oldWorkshopLevels: oldWorkshopLevels,
        oldWorkshopPlusLevels: oldWorkshopPlusLevels,
        hasPresets: hasPresets,

        activePresetSlot:
          activePreset === -1 ? -1 : presetIndices.indexOf(activePreset),
        upgradeIndices: {
          Attack: attackUpgradeIndices,
          Defense: defenseUpgradeIndices,
          Utility: utilityUpgradeIndices,
        },
        enhancementIndices: {
          Attack: attackEnhancementIndices,
          Defense: defenseEnhancementIndices,
          Utility: utilityEnhancementIndices,
        },
      };
    } catch (error) {
      var errorReport = errors.report("workshopSaveFile.parseWorkshopData", error, {
        data: data,
        oldWorkshopLevels: oldWorkshopLevels,
        oldWorkshopPlusLevels: oldWorkshopPlusLevels,
      });
      return errors.fail(errorReport);
    }
  },
};
