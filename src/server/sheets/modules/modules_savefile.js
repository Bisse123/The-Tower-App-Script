const moduleHeaders = {
  moduleEquipped: "moduleEquipped",
  inventory: "inventory",
  assistModuleSlots: "assistModuleSlots",
  moduleLevels: "slotLevels",
  modulePresets: "modulePresets",
}

const modulesSaveFile = {

  /**
   * Parses Modules data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseModulesData: function (data) {
    try {
      const moduleNames = modulesCatalog.moduleNames;

      const moduleRarities = modulesCatalog.moduleRarities;

      const moduleCategories = modulesCatalog.moduleCategories;

      const assistRarities = ["Epic", "Legendary", "Mythic", "Ancestral"];

      /**
       * Resolves a module effect id to its substat name and rarity.
       * @param {number} effectID
       * @returns {Object|null}
       */
      function lookupEffectID(effectID) {
        const effectRarities = modulesCatalog.effectRarities;

        const substatClusters = modulesCatalog.substatClusters;
        if (effectID < 1 || effectID > 331) {
          return null;
        }
        var id = 1;
        for (var c = 0; c < substatClusters.length; c++) {
          var cluster = substatClusters[c];
          var numRarities = cluster[1];
          if (effectID < id + numRarities) {
            var label = cluster[0];
            return {
              success: true,
              label: label,
              rarity: effectRarities[6 - numRarities + (effectID - id)],
            };
          }
          id += numRarities;
        }
        return null;
      }

      const equippedModulesData = data.moduleEquipped || [];
      const assistSlotData = data.assistModuleSlots || [];
      const inventoryData = data.inventory || [];
      const modulePresetsData = data.modulePresets || [];
      const moduleLevelsData = data.moduleLevels || [];

      var oldModuleInventory = {};
      var oldModulesPresets = {};
      var moduleInstances = {};

      /**
       * Records a module instance by guid, keeping its effects and category.
       * @param {Object} module
       * @param {string} [fallbackCategory]
       * @returns {Object|null} The stored instance, or null without a guid.
       */
      function collectModuleInstance(module, fallbackCategory) {
        if (!module || !module.guid) {
          return null;
        }
        if (!moduleInstances.hasOwnProperty(module.guid)) {
          const moduleInfo = moduleNames[module.infoIndex];
          const category = (moduleInfo && moduleInfo.category) || fallbackCategory;
          if (!category) {
            return null;
          }
          const name = (moduleInfo && moduleInfo.name) || "Any Other";
          moduleInstances[module.guid] = {
            guid: module.guid,
            name: name,
            category: category.toLowerCase(),
            rarityLevel: module.currentRarity,
            rarity: moduleRarities.hasOwnProperty(module.currentRarity)
              ? moduleRarities[module.currentRarity]
              : "Epic",
            effects: module.effects || [],
          };
        }
        return moduleInstances[module.guid];
      }

      /**
       * Writes one module instance into the inventory under a display name.
       * @param {Object} instance
       * @param {string} moduleName
       * @returns {void}
       */
      function writeModuleEntry(instance, moduleName) {
        instance.displayName = moduleName;
        var moduleSubstats = [];
        instance.effects.forEach(function (effectID) {
          var substatInfo = lookupEffectID(effectID);
          if (substatInfo) {
            moduleSubstats.push([substatInfo.label, substatInfo.rarity]);
          } else {
            moduleSubstats.push([null, null]);
          }
        });
        if (!oldModuleInventory.hasOwnProperty(instance.category)) {
          oldModuleInventory[instance.category] = {};
        }
        oldModuleInventory[instance.category][moduleName] = {
          rarity: instance.rarity,
          substats: moduleSubstats,
        };
      }

      /**
       * The preset slot a module category occupies.
       * @param {number} index
       * @param {string} presetName
       * @returns {Object|null} Null when the category is unknown.
       */
      function presetSlot(index, presetName) {
        const moduleCategory = (moduleCategories[index] || "").toLowerCase();
        if (!moduleCategory) {
          return null;
        }
        if (!oldModulesPresets.hasOwnProperty(moduleCategory)) {
          oldModulesPresets[moduleCategory] = {};
        }
        if (!oldModulesPresets[moduleCategory].hasOwnProperty(presetName)) {
          oldModulesPresets[moduleCategory][presetName] = {
            primary: "Any Other",
            secondary: "",
          };
        }
        return oldModulesPresets[moduleCategory][presetName];
      }

      /**
       * Places a stored module instance into its preset slot.
       * @param {string} guid
       * @returns {Object|null} Null when the guid is unknown.
       */
      function placeModuleInstance(guid) {
        const instance = guid ? moduleInstances[guid] : null;
        if (!instance) {
          return null;
        }
        if (instance.displayName) {
          return instance.displayName;
        }
        const placed = oldModuleInventory[instance.category] || {};
        if (!placed.hasOwnProperty(instance.name)) {
          writeModuleEntry(instance, instance.name);
          return instance.displayName;
        }

        const spareName = `Spare ${instance.name}`;
        if (!placed.hasOwnProperty(spareName)) {
          writeModuleEntry(instance, spareName);
        } else {
          instance.displayName = spareName;
        }
        return instance.displayName;
      }

      moduleLevelsData.forEach(function (moduleLevel, index) {
        const moduleCategory = moduleCategories[index].toLowerCase();
        if (!moduleCategory) {
          return;
        }
        if (!oldModuleInventory.hasOwnProperty(moduleCategory)) {
          oldModuleInventory[moduleCategory] = {};
        }
        oldModuleInventory[moduleCategory]["Highest Level"] = moduleLevel || null;
      });

      equippedModulesData.forEach(function (module, index) {
        if (!module) {
          return;
        }
        const moduleCategory = moduleCategories[index].toLowerCase();
        if (!moduleCategory) {
          return;
        }
        collectModuleInstance(module, moduleCategory);
        if (!oldModuleInventory.hasOwnProperty(moduleCategory)) {
          oldModuleInventory[moduleCategory] = {};
        }
        if (!oldModulesPresets.hasOwnProperty(moduleCategory)) {
          oldModulesPresets[moduleCategory] = {};
        }
      });

      assistSlotData.forEach(function (assistSlot, index) {
        const assistCategory = moduleCategories[index].toLowerCase();
        if (!assistCategory) {
          return;
        }
        const assistUnlocked = assistSlot.unlocked || false;
        const assistMainEffiency = String(
          assistSlot.mainEffectEfficiencyLevel || 0,
        ).padStart(2, "0");
        const assistSubEffiency = String(
          assistSlot.substatEfficiencyLevel || 0,
        ).padStart(2, "0");
        const assistUniqueEffect = assistSlot.uniqueEffectEfficiencyLevel || 0;
        const assistRarity = assistRarities[assistUniqueEffect] || "Epic";
        if (!oldModulesPresets.hasOwnProperty(assistCategory)) {
          oldModulesPresets[assistCategory] = {};
        }
        oldModulesPresets[assistCategory]["Assist Slot"] = {
          locked: assistUnlocked,
          rarity: assistRarity,
          multiplier: assistMainEffiency,
          substat: assistSubEffiency,
        };

        if (!assistUnlocked) {
          return;
        }
        const equippedAssistModule = assistSlot.equippedModule || {};
        if (!equippedAssistModule) {
          return;
        }
        collectModuleInstance(equippedAssistModule, assistCategory);
        if (!oldModuleInventory.hasOwnProperty(assistCategory)) {
          oldModuleInventory[assistCategory] = {};
        }
        oldModuleInventory[assistCategory]["Assist Level"] =
          assistSlot.level || null;
        if (!oldModulesPresets.hasOwnProperty(assistCategory)) {
          oldModulesPresets[assistCategory] = {};
        }
      });

      inventoryData.forEach(function (module) {
        collectModuleInstance(module);
      });

      var presetNames = [];
      modulePresetsData.forEach(function (preset) {
        var presetName = preset.presetName || null;
        if (!presetName) {
          return;
        }
        if (presetName === "Preset 1") {
          presetName = "Farming";
        }
        presetNames.push(presetName);

        (preset.primaryModuleGuids || []).forEach(function (moduleGuid, index) {
          const slot = presetSlot(index, presetName);
          if (!slot) {
            return;
          }
          slot.primary = placeModuleInstance(moduleGuid) || "Any Other";
        });
        (preset.assistModuleGuids || []).forEach(function (moduleGuid, index) {
          const moduleName = placeModuleInstance(moduleGuid);
          if (!moduleName) {
            return;
          }
          const slot = presetSlot(index, presetName);
          if (slot) {
            slot.secondary = moduleName;
          }
        });
      });

      Object.keys(moduleNames).forEach(function (infoIndex) {
        const moduleInfo = moduleNames[infoIndex];
        const moduleCategory = moduleInfo.category.toLowerCase();
        if (
          oldModuleInventory.hasOwnProperty(moduleCategory) &&
          oldModuleInventory[moduleCategory].hasOwnProperty(moduleInfo.name)
        ) {
          return;
        }
        var bestCopy = null;
        Object.keys(moduleInstances).forEach(function (guid) {
          const instance = moduleInstances[guid];
          if (instance.name !== moduleInfo.name || instance.displayName) {
            return;
          }
          if (
            !bestCopy ||
            (instance.rarityLevel || 0) > (bestCopy.rarityLevel || 0)
          ) {
            bestCopy = instance;
          }
        });
        if (bestCopy) {
          writeModuleEntry(bestCopy, moduleInfo.name);
        }
      });
      oldModulesPresets.presetNames = presetUtils.resolvePresetOrder(
        presetNames,
        presetUtils.templatePresetNames,
      ).order;

      const moduleOrder = Object.keys(moduleNames)
        .map(Number)
        .sort(function (a, b) {
          return a - b;
        })
        .map(function (infoIndex) {
          return moduleNames[infoIndex].name;
        });

      return {
        oldModulesInventory: oldModuleInventory,
        oldModulesPresets: oldModulesPresets,
        moduleOrder: moduleOrder,
      };
    } catch (error) {
      var errorReport = errors.report("modulesSaveFile.parseModulesData", error, {
        data: data,
        oldModuleInventory: oldModuleInventory,
        oldModulesPresets: oldModulesPresets,
      });
      return errors.fail(errorReport);
    }
  },
};
