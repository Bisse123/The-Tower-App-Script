const vaultHeaders = {
  vault: "vault",
}

const vaultSaveFile = {

  /**
   * Parses Vault data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseVaultData: function (data) {
    try {
      const vaultGroupsByIndex = vaultCatalog.vaultGroupsByIndex;

      const vault = data.vault || {};

      const unlockedGroupsName = "UnlockedGroups";
      const upgradesLevelName = "UpgradesLevel";

      var oldVault = {
        unlockedGroups: [],
        upgradesLevel: {},
      };

      var upgradeNamesByIndex = {};
      Object.keys(vaultGroupsByIndex).forEach((groupKey) => {
        const groupUpgrades = vaultGroupsByIndex[groupKey].upgrades || {};
        Object.keys(groupUpgrades).forEach((upgradeKey) => {
          upgradeNamesByIndex[upgradeKey] = groupUpgrades[upgradeKey];
        });
      });

      Object.keys(vault).forEach((key) => {
        if (key.includes(unlockedGroupsName)) {
          const unlockedGroups = vault[key];
          const unlockedElements = unlockedGroups.hasOwnProperty("Elements")
            ? unlockedGroups.Elements
            : [];
          unlockedElements.forEach((element) => {
            const groupInfo = vaultGroupsByIndex[element];
            if (groupInfo && !groupInfo.alwaysUnlocked) {
              oldVault.unlockedGroups.push(groupInfo.section);
            }
          });
        } else if (key.includes(upgradesLevelName)) {
          const upgradesLevel = vault[key];
          Object.keys(upgradesLevel).forEach((upgradeKey) => {
            const upgradeName = upgradeNamesByIndex[upgradeKey];
            if (upgradeName) {
              oldVault.upgradesLevel[upgradeName] = upgradesLevel[upgradeKey];
            }
          });
        }
      });

      var vaultIndices = [];
      var alwaysUnlocked = [];
      var columnsByName = {};

      Object.keys(vaultGroupsByIndex)
        .map(Number)
        .sort((a, b) => a - b)
        .forEach((groupIndex) => {
          const groupInfo = vaultGroupsByIndex[groupIndex];
          if (groupInfo.alwaysUnlocked) {
            alwaysUnlocked.push(groupInfo.section);
          }
          if (!columnsByName.hasOwnProperty(groupInfo.group)) {
            columnsByName[groupInfo.group] = {
              group: groupInfo.group,
              sections: [],
            };
            vaultIndices.push(columnsByName[groupInfo.group]);
          }

          const groupUpgrades = groupInfo.upgrades || {};
          const upgradeNames = Object.keys(groupUpgrades)
            .map(Number)
            .sort((a, b) => a - b)
            .map((upgradeIndex) => groupUpgrades[upgradeIndex]);

          columnsByName[groupInfo.group].sections.push({
            section: groupInfo.section,
            upgrades: upgradeNames,
          });
        });

      return {
        success: true,
        oldVault: oldVault,
        vaultIndices: vaultIndices,
        alwaysUnlocked: alwaysUnlocked,
      };
    } catch (error) {
      var errorReport = errors.report("vaultSaveFile.parseVaultData", error, {
        data: data,
        oldVault: oldVault,
      });
      return errors.fail(errorReport);
    }
  },
};
