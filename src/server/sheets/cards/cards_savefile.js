const cardsHeaders = {
  cardLevel: "cardLevel",
  cardMasteryUnlocked: "cardMasteryUnlocked",
  presetNames: "presetName",
  presetSlots: "slotPresetCardAssignedBool",
  presetCards: "slotPresetCardInt",
  slotsUnlocked: "slotsUnlocked",
}

const cardsSaveFile = {

  /**
   * Parses Cards data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseCardsData: function (data) {
    try {
      const cardNamesByIndex = {
        0: "Damage",
        1: "Attack Speed",
        2: "Health",
        3: "Health Regen",
        4: "Range",
        5: "Cash",
        6: "Coins",
        7: "Slow Aura",
        10: "Critical Chance",
        11: "Enemy Balance",
        12: "Extra Defense",
        13: "Fortress",
        15: "Free Upgrades",
        16: "Extra Orb",
        18: "Plasma Cannon",
        19: "Critical Coin",
        20: "Wave Skip",
        21: "Intro Sprint",
        22: "Land Mine Stun",
        23: "Recovery Package Chance",
        25: "Death Ray",
        26: "Energy Net",
        27: "Super Tower",
        28: "Second Wind",
        29: "Demon Mode",
        30: "Energy Shield",
        31: "Wave Accelerator",
        32: "Berserker",
        33: "Ultimate Crit",
        34: "Nuke",
        35: "Area of Effect",
        37: "Cells",
      };
      var cardNameIndices = [];
      Object.keys(cardNamesByIndex).forEach(function (index) {
        cardNameIndices[Number(index)] = cardNamesByIndex[index];
      });
      const cardLevel = data.cardLevel || [];
      const cardMasteryUnlocked = data.cardMasteryUnlocked || [];

      const presetOrder = presetUtils.resolvePresetOrder(
        data.presetNames || [],
        presetUtils.templatePresetNames,
      );
      const presetNames = presetOrder.order;
      const presetIndices = presetOrder.indices;

      const presetSlots = data.presetSlots || [];
      const presetCards = data.presetCards || [];
      const slotsUnlocked = data.slotsUnlocked || 0;

      var oldCardsLevel = {};
      cardNameIndices.forEach(function (cardName, i) {
        if (!cardName) return;
        oldCardsLevel[cardName] = [cardLevel[i], cardMasteryUnlocked[i]];
      });

      const numPresets = presetNames.length;
      const slotsPerPreset =
        numPresets > 0 ? Math.floor(presetSlots.length / numPresets) : 0;
      var oldCardsPreset = {};
      presetNames.forEach(function (name, slot) {
        if (!name) return;
        var sourceIndex = presetIndices[slot];
        var slotStart = sourceIndex * slotsPerPreset;
        var cards = [];
        presetSlots
          .slice(slotStart, slotStart + slotsPerPreset)
          .forEach(function (assigned, s) {
            if (assigned) {
              var resolvedName =
                cardNameIndices[presetCards[slotStart + s]] || null;
              if (resolvedName) {
                cards.push(resolvedName);
              }
            }
          });
        oldCardsPreset[name] = {
          cards: cards,
          remove: [],
          order: slot + 1,
        };
      });

      return {
        success: true,
        oldCardsLevel: oldCardsLevel,
        oldCardsPreset: oldCardsPreset,
        oldCardSlots: slotsUnlocked,
        cardNameIndices: cardNameIndices,
      };
    } catch (error) {
      var errorReport = errors.report("cardsSaveFile.parseCardsData", error, {
        data: data,
        oldCardsLevel: oldCardsLevel,
        oldCardsPreset: oldCardsPreset,
      });
      return errors.fail(errorReport);
    }
  },
};
