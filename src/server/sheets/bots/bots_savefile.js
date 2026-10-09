const botHeaders = {
  presetNames: "botPresetName",
  flameBotPresets: "flameBotPresets",
  thunderBotPresets: "thunderBotPresets",
  goldenBotPresets: "goldenBotPresets",
  amplifyBotPresets: "amplifyBotPresets",
  botBotPresets: "botBotPresets",
  synchronicityPresets: "synchronicityPresets",
}

const botsSaveFile = {

  /**
   * Parses Bots data out of a decoded save file.
   * @param {Object} data
   * @param {number} presetCount How many of the stored presets the player has unlocked; only
   *   those are read.
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseBotsData: function (data, presetCount) {
    try {
      const targetBots = {
        "Flame Bot": {
          upgrades: ["Damage", "Range", "Cooldown", "Damage R."],
          plusUpgrade: "Wildfire",
        },
        "Thunder Bot": {
          upgrades: ["Linger", "Range", "Cooldown", "Duration"],
          plusUpgrade: "Titan Shock",
        },
        "Golden Bot": {
          upgrades: ["Bonus", "Range", "Cooldown", "Duration"],
          plusUpgrade: "Bonus Cell",
        },
        "Amplify Bot": {
          upgrades: ["Bonus", "Range", "Cooldown", "Duration"],
          plusUpgrade: "Echoing Shot",
        },
        "Bot Bot": {
          upgrades: ["Bonus", "Range", "Cooldown", "Duration"],
          plusUpgrade: "Maximum Power",
        },
      };

      const presetOrder = presetUtils.resolvePresetOrder(
        (data.presetNames || []).slice(0, presetCount),
        presetUtils.templatePresetNames,
      );
      const presetNames = presetOrder.order;
      const presetIndices = presetOrder.indices;

      const flameBotData = data.flameBotPresets || {};
      const thunderBotData = data.thunderBotPresets || {};
      const goldenBotData = data.goldenBotPresets || {};
      const amplifyBotData = data.amplifyBotPresets || {};
      const botBotData = data.botBotPresets || {};

      const syncPresets = data.synchronicityPresets || {};
      var oldBots = {
        presetNames: presetNames,
        data: {},
      };

      presetNames.forEach(function (presetName, slot) {
        const index = presetIndices[slot];
        const allBotPresets = [
          flameBotData[index] || {},
          thunderBotData[index] || {},
          goldenBotData[index] || {},
          amplifyBotData[index] || {},
          botBotData[index] || {},
        ];
        allBotPresets.forEach(function (botPreset, botIndex) {
          var botName = Object.keys(targetBots)[botIndex];
          var botLevels = botPreset.levels || [];
          var props = targetBots[botName].upgrades.reduce(function (
            acc,
            upgrade,
            idx,
          ) {
            var level = botLevels[idx];
            acc[upgrade] = level ? String(level).padStart(2, "0") : "00";
            return acc;
          }, {});

          props[targetBots[botName].plusUpgrade] = "Lo";
          if (botPreset.plusUnlocked) {
            props[targetBots[botName].plusUpgrade] =
              botPreset.plusLevel === null ||
              botPreset.plusLevel === undefined ||
              botPreset.plusLevel === ""
                ? "Lo"
                : String(botPreset.plusLevel).padStart(2, "0");
          }
          if (!oldBots.data.hasOwnProperty(botName)) {
            oldBots.data[botName] = {
              presets: {},
            };
          }
          const active = botPreset.unlocked !== undefined ? botPreset.unlocked : null;
          oldBots.data[botName].presets[presetName] = {
            props: props,
            sync: syncPresets[index][botIndex] || null,
            active: active,
          };
        });
      });

      return {
        success: true,
        oldBots: oldBots,
        targetBots: targetBots,
        botOrder: Object.keys(targetBots),
      };
    } catch (error) {
      var errorReport = errors.report("botsSaveFile.parseBotsData", error, {
        data: data,
        oldBots: oldBots,
      });
      return errors.fail(errorReport);
    }
  },
};
