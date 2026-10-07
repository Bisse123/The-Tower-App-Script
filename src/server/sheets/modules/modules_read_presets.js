const modulesPresetsReader = {

  /**
   * Extracts ModulesPresets from a v5.0 sheet's values.
   * @param {Array<Array<*>>} oldModulesPresetsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_0ModulesPresets: function (oldModulesPresetsValues) {
    try {
      console.log("Called: modulesPresetsReader.getVersion5_0ModulesPresets");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = modulesLayout.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesPresetsValues,
      );

      var oldModulesPresets = {};
      var presetNames = [];
      var presetNamesLocked = false;
      targetModuleTypes.forEach(function (moduleType) {
        if (typeof oldModuleTypeIndex[moduleType] === "undefined") return;
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        oldModulesPresets[moduleType] = {};
        var row = oldModulesPresetsValues[rowIdx + 2];
        var presetIndex = 0;
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Primary Slot") {
            var rawPresetName = oldModulesPresetsValues[rowIdx][col]
              ? String(oldModulesPresetsValues[rowIdx][col]).trim()
              : "";
            if (!rawPresetName) {
              continue;
            }
            var presetName;
            if (presetIndex < presetNames.length) {
              presetName = presetNames[presetIndex];
            } else if (!presetNamesLocked) {
              presetName = rawPresetName;
              presetNames.push(rawPresetName);
            } else {
              break;
            }
            var primaryName = oldModulesPresetsValues[rowIdx + 2][col + 1]
              ? String(oldModulesPresetsValues[rowIdx + 2][col + 1]).trim()
              : "";
            var secondaryName = oldModulesPresetsValues[rowIdx + 3][col + 1]
              ? String(oldModulesPresetsValues[rowIdx + 3][col + 1]).trim()
              : "";
            oldModulesPresets[moduleType][presetName] = {
              primary: primaryName,
              secondary: secondaryName,
            };
            presetIndex++;
          }
        }
        if (presetIndex > 0) {
          presetNamesLocked = true;
        }
        var assistSlotCol =
          oldModulesPresetsValues[rowIdx].indexOf("Assist Slot");
        if (assistSlotCol !== -1) {
          var assistLocked = oldModulesPresetsValues[rowIdx][assistSlotCol + 2];

          var assistRarity = String(
            oldModulesPresetsValues[rowIdx + 1][assistSlotCol + 1],
          ).trim();
          var assistMultiplier = String(
            oldModulesPresetsValues[rowIdx + 2][assistSlotCol + 2],
          ).trim();
          var assistSubstat = String(
            oldModulesPresetsValues[rowIdx + 3][assistSlotCol + 2],
          ).trim();
          oldModulesPresets[moduleType]["Assist Slot"] = {
            locked: assistLocked,
            rarity: assistRarity,
            multiplier: assistMultiplier,
            substat: assistSubstat,
          };
        }
      });
      oldModulesPresets.presetNames = presetUtils.resolvePresetOrder(
        presetNames,
        presetUtils.templatePresetNames,
      ).order;

      return {
        success: true,
        oldModulesPresets: oldModulesPresets,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion5_0ModulesPresets", error, {
        oldModulesPresetsValues: oldModulesPresetsValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts ModulesPresets from a v4.0 sheet's values.
   * @param {Array<Array<*>>} oldModulesPresetsValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0ModulesPresets: function (oldModulesPresetsValues) {
    try {
      console.log("Called: modulesPresetsReader.getVersion4_0ModulesPresets");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = modulesLayout.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesPresetsValues,
      );

      var oldModulesPresets = {};
      var presetNames = [];
      var presetNamesLocked = false;
      targetModuleTypes.forEach(function (moduleType) {
        if (typeof oldModuleTypeIndex[moduleType] === "undefined") return;
        var rowIdx = oldModuleTypeIndex[moduleType] + 1;
        oldModulesPresets[moduleType] = {};
        var row = oldModulesPresetsValues[rowIdx];
        var presetIndex = 0;
        for (var col = 0; col < row.length; col++) {
          if (String(row[col]).trim() === "Module Name") {
            var rawPresetName = String(
              oldModulesPresetsValues[rowIdx - 1][col],
            ).trim();
            var moduleName = String(
              oldModulesPresetsValues[rowIdx + 1][col],
            ).trim();
            if (rawPresetName && moduleName) {
              var presetName;
              if (presetIndex < presetNames.length) {
                presetName = presetNames[presetIndex];
              } else if (!presetNamesLocked) {
                presetName = rawPresetName;
                presetNames.push(rawPresetName);
              } else {
                break;
              }
              oldModulesPresets[moduleType][presetName] = {
                primary: moduleName,
                secondary: "",
              };
              presetIndex++;
            }
          }
        }
        if (presetIndex > 0) {
          presetNamesLocked = true;
        }
      });
      oldModulesPresets.presetNames = presetNames;

      return {
        success: true,
        oldModulesPresets: oldModulesPresets,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion4_0ModulesPresets", error, {
        oldModulesPresetsValues: oldModulesPresetsValues,
      });
      return errors.fail(errorReport);
    }
  },
};
