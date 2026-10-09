const ultimateWeaponHeaders = {
  ultimateWeaponLevel: "ultimateWeaponLevel",
  ultimateWeaponUnlocked: "ultimateWeaponUnlocked",
  ultimateWeaponPlusLevel: "ultimateWeaponPlusLevel",
  ultimateWeaponPlusUnlocked: "ultimateWeaponPlusUnlocked",
}

const ultimateSaveFile = {

  /**
   * Parses Ultimate_Weapons data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseUltimateWeaponData: function (data) {
    try {
      const targetWeapons = {
        "Chain Lightning": { upgrades: ["Damage", "Quantity", "Chance"], plusUpgrade: "Smite" },
        "Smart Missiles": { upgrades: ["Damage", "Quantity", "Cooldown"], plusUpgrade: "Cover Fire" },
        "Death Wave": { upgrades: ["Damage", "Quantity", "Cooldown"], plusUpgrade: "Kill Wall" },
        "Chrono Field": { upgrades: ["Duration", "Speed Reduction", "Cooldown"], plusUpgrade: "Chrono Loop" },
        "Inner Land Mines": { upgrades: ["Damage", "Quantity", "Cooldown"], plusUpgrade: "Charged Mines" },
        "Golden Tower": { upgrades: ["Multiplier", "Duration", "Cooldown"], plusUpgrade: "Golden Combo" },
        "Poison Swamp": { upgrades: ["Damage", "Duration", "Cooldown"], plusUpgrade: "Death Creep" },
        "Black Hole": { upgrades: ["Size", "Duration", "Cooldown"], plusUpgrade: "Consume" },
        "Spotlight": { upgrades: ["Multiplier", "Angle", "Quantity"], plusUpgrade: "Light Range" },
      };
      const ultimateWeaponUnlocked = data.ultimateWeaponUnlocked || [];
      const ultimateWeaponLevel = data.ultimateWeaponLevel || [];
      const ultimateWeaponPlusUnlocked = data.ultimateWeaponPlusUnlocked || [];
      const ultimateWeaponPlusLevel = data.ultimateWeaponPlusLevel || [];

      var oldUltimate = {};

      Object.keys(targetWeapons).forEach(function (weaponName, i) {
        var { upgrades, plusUpgrade } = targetWeapons[weaponName];
        var weaponLevels = {};
        upgrades.forEach(function (attr, j) {
          var idx = i * upgrades.length + j;
          var level = ultimateWeaponLevel[idx];
          weaponLevels[attr] = level ? String(level).padStart(2, "0") : "00";
        });
        if (!ultimateWeaponPlusUnlocked[i]) {
          weaponLevels[plusUpgrade] = "Lo";
        } else {
          var level = ultimateWeaponPlusLevel[i];
          weaponLevels[plusUpgrade] = level ? String(level).padStart(2, "0") : "00";
        }
        oldUltimate[weaponName] = {
          unlocked: ultimateWeaponUnlocked[i] || null,
          levels: weaponLevels,
        };
      });

      return {
        success: true,
        oldUltimate: oldUltimate,
        targetWeapons: targetWeapons,
        weaponOrder: Object.keys(targetWeapons),
      };
    } catch (error) {
      var errorReport = errors.report("ultimateSaveFile.parseUltimateWeaponData", error, {
        data: data,
        oldUltimate: oldUltimate,
      });
      return errors.fail(errorReport);
    }
  },
};
