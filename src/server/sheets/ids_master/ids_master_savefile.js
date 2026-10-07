const MasterHeaders = {
  globalPresets: "globalPresets",
  workshopPresetNames: "workshopPresetName",
  cardPresetNames: "presetName",
  botPresetNames: "botPresetName",
  modulePresets: "modulePresets",
  guardianPresets: "guardianPresets",
}

const masterSaveFile = {

  /**
   * Parses IDS_Master data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseMasterData: function (data) {
    try {
      const globalPresets = data.globalPresets || [];
      const workshopPresetNames = data.workshopPresetNames || [];
      const cardPresetNames = data.cardPresetNames || [];
      const botPresetNames = data.botPresetNames || [];
      const modulePresets = data.modulePresets || [];
      const guardianPresets = data.guardianPresets || [];
      const modulePresetNames = modulePresets.map((preset) => preset.presetName);
      const guardianPresetNames = guardianPresets.map((preset) => preset.presetName);

      const presetInfo = [
        {"type": "Workshop", "indexName": "workshopIndex", "data": workshopPresetNames},
        {"type": "Cards", "indexName": "cardsIndex", "data": cardPresetNames},
        {"type": "Bots", "indexName": "botsIndex", "data": botPresetNames},
        {"type": "Modules", "indexName": "modulesIndex", "data": modulePresetNames},
        {"type": "Guardians", "indexName": "guardiansIndex", "data": guardianPresetNames},
    ]

      var oldPresetNames = [];
      var oldPresetsData = {
        data: {},
      };
      globalPresets.forEach((preset, index) => {
        if (index == globalPresets.length - 1) return;
        const globalPresetName = preset.presetName;
        if (!globalPresetName) {
          return;
        }
        oldPresetNames.push(globalPresetName);
        if (!oldPresetsData.data.hasOwnProperty(globalPresetName)) {
          oldPresetsData.data[globalPresetName] = {};
        }
        presetInfo.forEach((info) => {
          const presetType = info.type;
          const indexName = info.indexName;
          const index = preset[indexName];
          const data = info.data;
          const presetName = data[index];
          if (presetName) {
            oldPresetsData.data[globalPresetName][presetType] = presetName;
          }
        });
      });

      oldPresetsData.presetNames = presetUtils.resolvePresetOrder(
          oldPresetNames,
          presetUtils.templatePresetNames,
        ).order;

      const presetTypesOrder = presetInfo.map((info) => info.type);
      return {
        success: true,
        success: true,
        oldPresetsData: oldPresetsData,
        presetTypesOrder: presetTypesOrder,
      };
    } catch (error) {
      var errorReport = errors.report("masterSaveFile.parseMasterData", error, {
        data: data,
        oldPresetsData: oldPresetsData,
      });
      return errors.fail(errorReport);
    }
  },
};
