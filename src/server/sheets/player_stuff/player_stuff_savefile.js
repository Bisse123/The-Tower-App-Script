const PlayerStuffHeaders = {
  playerID: "playfabID",
  currentTier: "currentTier",
  tourneyID: "leagueID",
  addPack: "disableAdsUnlockedBool",
  starterPack: "starterPackUnlockedBool",
  epicPack: "epicPackUnlockedBool",
  highestWavePerTier: "highestWaveThisTier",
  premiumPass: "milestonesPremiumUnlocked",
  atkDissonance: "dissonanceDamageBoost",
  hpDissonance: "dissonanceHealthBoost",
  coinDissonance: "dissonanceCoinBoost",
  uwDissonance: "dissonanceUltDamageBoost",
  totalCoinsEarned: "totalCoinsEarned",
  totalStonesEarned: "totalStonesEarned",
  totalStonesBought: "totalStonesBought",
  totalGemsEarned: "totalGemsEarned",
  totalGemsBought: "totalGemsBought",
  totalKeysEarned: "totalKeysEarned",
  battleHistory: "battleHistory",
}

const playerStuffSaveFile = {

  /**
   * Parses Player_&_Stuff data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parsePlayerStuffData: function (data) {
    try {
      const tourneyNames = playerStuffCatalog.tourneyNames;

      const tourneyName = data.tourneyID ? tourneyNames[data.tourneyID] : null;
      const highestWavePerTier = data.highestWavePerTier || [];
      const premiumPass = data.premiumPass || [];
      const atkDissonance = data.atkDissonance || [];
      const hpDissonance = data.hpDissonance || [];
      const coinDissonance = data.coinDissonance || [];
      const uwDissonance = data.uwDissonance || [];

      const lifetimeCoins = data.totalCoinsEarned || 0;
      const lifetimeStones = data.totalStonesEarned + data.totalStonesBought || 0;
      const lifetimeGems = data.totalGemsEarned + data.totalGemsBought || 0;
      const lifetimeKeys = data.totalKeysEarned || 0;

      const battleHistory = data.battleHistory || [];

      /**
       * Formats a lifetime total with a magnitude suffix.
       * @param {*} value
       * @returns {string|null} Null when there is no usable number.
       */
      function formatLifeTime(value) {
        if (value === null || value === undefined || value === "") {
          return null;
        }
        if (typeof value !== "number" || !isFinite(value)) {
          return value;
        }

        /**
         * The magnitude suffix for a thousands group: K, M, B, then aa, ab, …
         * @param {number} group
         * @returns {string}
         */
        function getSuffix(group) {
          var named = ["", "K", "M", "B", "T", "q", "Q", "s", "S", "O", "N", "D"];
          if (group < named.length) {
            return named[group];
          }
          var n = group - named.length;
          var first = Math.floor(n / 26);
          var second = n % 26;
          return (
            String.fromCharCode(97 + first) + String.fromCharCode(97 + second)
          );
        }

        var negative = value < 0;
        var num = Math.abs(value);

        if (num < 1000) {
          return (negative ? "-" : "") + String(Math.round(num * 100) / 100);
        }

        var group = Math.floor(Math.log10(num) / 3);
        var mantissa = num / Math.pow(1000, group);

        if (mantissa >= 1000) {
          mantissa /= 1000;
          group += 1;
        }
        var rounded = Math.round(mantissa * 100) / 100;
        if (rounded >= 1000) {
          rounded /= 1000;
          group += 1;
        }

        return (negative ? "-" : "") + rounded.toFixed(2) + getSuffix(group);
      }

      var allBattlesCoinPerHour = battleHistory
        .map(function (battle) {
          if (battle && battle.coinsEarned && battle.realTime) {
            var hours = battle.realTime / 3600;
            if (hours > 0) {
              return battle.coinsEarned / hours;
            }
          }
          return null;
        })
        .filter(function (cph) {
          return cph !== null;
        })
        .sort(function (a, b) {
          return b - a;
        });

      var numBattles = 3;
      const coinPerHour =
        allBattlesCoinPerHour.length > 0
          ? allBattlesCoinPerHour.slice(0, numBattles).reduce(function (
              sum,
              cph,
            ) {
              return sum + cph;
            }, 0) / Math.min(numBattles, allBattlesCoinPerHour.length)
          : null;

      var oldPlayerStuffTierData = {};
      var oldPlayerStuffStatsData = {
        Stat: {
          "Player ID": data.playerID,
          "Farming Tier": "Tier " + data.currentTier,
          "Tourney League": tourneyName,
          "Lifetime Coins": formatLifeTime(lifetimeCoins),
          Stones: formatLifeTime(lifetimeStones),
          Gems: formatLifeTime(lifetimeGems),
          Keys: formatLifeTime(lifetimeKeys),
          "Coin / Hour": formatLifeTime(coinPerHour),
        },
        "Premium Packs": {
          "Disable Ads": data.addPack,
          "Starter Pack": data.starterPack,
          "Epic Pack": data.epicPack,
        },
      };

      var nextPremium = 0;
      for (var tier = 0; tier < highestWavePerTier.length; tier++) {

        var wave = highestWavePerTier[tier];
        if (wave <= 0) {
          continue;
        }
        var premium = null;
        if (tier % 3 === 1) {
          premium = premiumPass[nextPremium] || null;
          nextPremium++;
        }

        oldPlayerStuffTierData["Tier " + tier] = {
          wave: wave,
          diss: {
            attack: atkDissonance[tier] || 0,
            defense: hpDissonance[tier] || 0,
            utility: coinDissonance[tier] || 0,
            ultimate: uwDissonance[tier] || 0,
          },
          premium: premium,
        };
      }

      return {
        success: true,
        oldPlayerStuffTierData: oldPlayerStuffTierData,
        oldPlayerStuffStatsData: oldPlayerStuffStatsData,
        statOrder: Object.keys(oldPlayerStuffStatsData.Stat),
        premiumOrder: Object.keys(oldPlayerStuffStatsData["Premium Packs"]),
      };
    } catch (error) {
      var errorReport = errors.report("playerStuffSaveFile.parsePlayerStuffData", error, {
        data: data,
        oldPlayerStuffTierData: oldPlayerStuffTierData,
        oldPlayerStuffStatsData: oldPlayerStuffStatsData,
      });
      return errors.fail(errorReport);
    }
  },
};
