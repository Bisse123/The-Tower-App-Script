const collectionConvertersV3 = {

  /**
   * Reads IDS_Collection data from a v3.2 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_2: function (oldSheetID) {
    try {
      console.log("Called: collectionConvertersV3.version3_2");

      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E",
          "Lab Planner": "Lab Planner",
          "Workshop Levels": "EXPORT_WS!B2:M",
          "Workshop Plus": "EXPORT_WS!P2:V",
          "Workshop Ratio": "Desired Ratios",
          "Ultimate Weapon": "EXPORT_UW!C5:H",
          "Themes & Songs": "Themes & Songs",
          Bots: "EXPORT_Bots!C4:L",
          Relics: "Relics",
          "Vault Harmony": "Vault_Harmony",
          "Vault Power": "Vault_Power",
          "Card Preset": "Card Preset",
          "Card Tracker": "Card and Mastery Tracker",
          "Cards Levels": "EXPORT_Cards!B5:D",
          "Cards Slots": "EXPORT_Cards!C2",
          "Modules Inventory": "Modules Inventory",
          "Modules Presets": "Modules Presets",
          "Modules Tracker": "Modules Tracker",
          Guardians: "EXPORT_Guardians!B5:F",
          "Player Tier": "EXPORT_Player!B3:H",
          "Player Stat": "EXPORT_Player!J3:K",
          "Player Perks": "Perk Preset",
        },
        formulas: {
          "Lab Planner": "Lab Planner",
          "UW Cost Calculator": "UW Cost Calculator v3",
          "Modules Tracker": "Modules Tracker",
        },
      };

      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      var batchValuesResults = [];
      var batchFormulasResults = [];

      if (valuesRanges.length > 0) {
        batchValuesResults = SheetsAPI.batchGetValues(oldSheetID, valuesRanges);
        if (!batchValuesResults) {
          console.log(`Could not read IDS Collection values data`);
          return {
            success: false,
            message: "Could not read IDS Collection values data",
          };
        }
      }

      if (formulasRanges.length > 0) {
        batchFormulasResults = SheetsAPI.batchGetFormulas(
          oldSheetID,
          formulasRanges,
        );
        if (!batchFormulasResults) {
          console.log(`Could not read IDS Collection formulas data`);
          return {
            success: false,
            message: "Could not read IDS Collection formulas data",
          };
        }
      }

      var getBatchResult = function (key, type) {
        type = type || "values";

        if (type === "values") {
          var index = Object.keys(rangeMap.values).indexOf(key);
          return index !== -1 && batchValuesResults[index]
            ? batchValuesResults[index]
            : null;
        } else if (type === "formulas") {
          var index = Object.keys(rangeMap.formulas).indexOf(key);
          return index !== -1 && batchFormulasResults[index]
            ? batchFormulasResults[index]
            : null;
        }
        return null;
      };

      var collectedData = {};

      var labLevelsResult = getBatchResult("Lab Levels", "values");
      var labPlannerValuesResult = getBatchResult("Lab Planner", "values");
      var labPlannerFormulasResult = getBatchResult("Lab Planner", "formulas");
      if (labLevelsResult && labLevelsResult.values) {
        var labLevelsValues = labLevelsResult.values;
        var labPlannerValues =
          labPlannerValuesResult && labPlannerValuesResult.values
            ? labPlannerValuesResult.values
            : null;
        var labPlannerFormulas =
          labPlannerFormulasResult && labPlannerFormulasResult.values
            ? labPlannerFormulasResult.values
            : null;

        var labLevelsData = labReader.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = labReader.getVersion1_0LabPlanner(
          labPlannerValues,
          labPlannerFormulas,
          labLevelsData.oldLabLevels,
          labLevelsData.oldLabMax,
        );

        var labSuccess = labLevelsData.success && labPlannerData.success;
        collectedData.Laboratory = {
          success: labSuccess,
          message: labSuccess
            ? "Laboratory data retrieved successfully"
            : "Error retrieving Laboratory data",
          oldLabLevels: labLevelsData.oldLabLevels,
          oldLabPlanner: labPlannerData.oldLabPlanner,
        };
      }

      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopPlusRatioResult = getBatchResult("Workshop Ratio", "values");
      if (
        workshopLevelsResult &&
        workshopLevelsResult.values &&
        workshopPlusResult &&
        workshopPlusResult.values &&
        workshopPlusRatioResult &&
        workshopPlusRatioResult.values
      ) {
        var workshopLevelsValues = workshopLevelsResult.values;
        var workshopPlusLevelsValues = workshopPlusResult.values;
        var workshopPlusRatioValues = workshopPlusRatioResult.values;

        var workshopLevelsData =
          workshopReader.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshopReader.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshopReader.getVersion2_2_8WorkshopPlusRatios(
          workshopPlusLevelsData.oldWorkshopPlusLevels.presetNames,
          workshopPlusRatioValues,
        );
        var workshopSuccess =
          workshopLevelsData.success &&
          workshopPlusLevelsData.success &&
          workshopPlusRatiosData.success;
        collectedData.Workshop = {
          success: workshopSuccess,
          message: workshopSuccess
            ? "Workshop data retrieved successfully"
            : "Error retrieving Workshop data",
          oldWorkshopLevels: workshopLevelsData.oldWorkshopLevels,
          oldWorkshopPlusLevels: workshopPlusLevelsData.oldWorkshopPlusLevels,
          oldWorkshopPlusRatios: workshopPlusRatiosData.oldWorkshopPlusRatios,
        };
      }

      var ultimateResult = getBatchResult("Ultimate Weapon", "values");
      var ultimateCostCalculatorResult = getBatchResult(
        "UW Cost Calculator",
        "formulas",
      );
      if (
        ultimateResult &&
        ultimateResult.values &&
        ultimateCostCalculatorResult &&
        ultimateCostCalculatorResult.values
      ) {
        var ultimateValues = ultimateResult.values;
        var ultimateCostCalculatorValues = ultimateCostCalculatorResult.values;

        var ultimateWeaponsData =
          ultimateReader.getVersion3_1_1UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimateReader.getVersion1_0CostCalculator(
          ultimateCostCalculatorValues,
        );

        var ultimateSuccess =
          ultimateWeaponsData.success && costCalculatorData.success;
        collectedData["Ultimate Weapon"] = {
          success: ultimateSuccess,
          message: ultimateSuccess
            ? "Ultimate Weapon data retrieved successfully"
            : "Error retrieving Ultimate Weapon data",
          oldUltimate: ultimateWeaponsData.oldUltimate,
          oldUltimateCostCalculator: costCalculatorData["UW Cost Calculator"],
        };
      }

      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themes.getVersion2_1_6Themes(themesResult.values);
        var relicsData = relics.getVersion1_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = botsReader.getVersion3_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      var harmonyResult = getBatchResult("Vault Harmony", "values");
      var powerResult = getBatchResult("Vault Power", "values");
      if (
        harmonyResult &&
        harmonyResult.values &&
        powerResult &&
        powerResult.values
      ) {
        var harmonyValues = harmonyResult.values;
        var powerValues = powerResult.values;

        var harmonyVaultData = vaultReader.getVersion3_1Vault(harmonyValues);
        var powerVaultData = vaultReader.getVersion3_1Vault(powerValues);

        var vaultSuccess = harmonyVaultData.success && powerVaultData.success;
        collectedData.Vault = {
          success: vaultSuccess,
          message: vaultSuccess
            ? "Vault data retrieved successfully"
            : "Error retrieving Vault data",
          oldVaultHarmony: harmonyVaultData.oldVault,
          oldVaultPower: powerVaultData.oldVault,
        };
      }

      var cardsPresetResult = getBatchResult("Card Preset", "values");
      var cardsTrackerResult = getBatchResult("Card Tracker", "values");
      var cardsLevelsResult = getBatchResult("Cards Levels", "values");
      var cardsSlotsResult = getBatchResult("Cards Slots", "values");
      if (
        cardsPresetResult &&
        cardsPresetResult.values &&
        cardsTrackerResult &&
        cardsTrackerResult.values &&
        cardsLevelsResult &&
        cardsLevelsResult.values &&
        cardsSlotsResult &&
        cardsSlotsResult.values
      ) {
        var cardsPresetValues = cardsPresetResult.values;
        var cardsTrackerValues = cardsTrackerResult.values;
        var cardsLevelValues = cardsLevelsResult.values;
        var cardsSlotsValues = cardsSlotsResult.values;

        var cardsPresetData = cardsReader.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cardsReader.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cardsReader.getVersion1_0CardsTracker(cardsTrackerValues);

        var cardsSuccess =
          cardsPresetData.success &&
          cardsLevelData.success &&
          cardsTrackerData.success;
        collectedData.Cards = {
          success: cardsSuccess,
          message: cardsSuccess
            ? "Cards data retrieved successfully"
            : "Error retrieving Cards data",
          oldCardsPreset: cardsPresetData.oldCardsPreset,
          shouldRemoveUsedCards: cardsPresetData.shouldRemoveUsedCards,
          oldCardsLevel: cardsLevelData.oldCardsLevel,
          oldCardSlots: cardsLevelData.oldCardSlots,
          oldCardsTracker: cardsTrackerData.oldCardsTracker,
        };
      }

      var modulesInventoryResult = getBatchResult(
        "Modules Inventory",
        "values",
      );
      var modulesPresetsResult = getBatchResult("Modules Presets", "values");
      var modulesTrackerResult = getBatchResult("Modules Tracker", "values");
      var modulesTrackerFormulasResult = getBatchResult(
        "Modules Tracker",
        "formulas",
      );
      if (
        modulesInventoryResult &&
        modulesInventoryResult.values &&
        modulesPresetsResult &&
        modulesPresetsResult.values &&
        modulesTrackerResult &&
        modulesTrackerResult.values &&
        modulesTrackerFormulasResult &&
        modulesTrackerFormulasResult.values
      ) {
        var modulesInventoryValues = modulesInventoryResult.values;
        var modulesPresetsValues = modulesPresetsResult.values;
        var modulesTrackerValues = modulesTrackerResult.values;
        var modulesTrackerFormulas = modulesTrackerFormulasResult.values;
        var modulesInventoryData = modulesInventoryReader.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modulesPresetsReader.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modulesTrackerReader.getVersion4_7ModulesTracker(
          modulesTrackerValues,
          modulesTrackerFormulas,
        );
        var modulesSuccess =
          modulesInventoryData.success &&
          modulesPresetsData.success &&
          modulesTrackerData.success;
        collectedData.Modules = {
          success: modulesSuccess,
          message: modulesSuccess
            ? "Modules data retrieved successfully"
            : "Error retrieving Modules data",
          oldModulesInventory: modulesInventoryData.oldModulesInventory,
          oldModulesPresets: modulesPresetsData.oldModulesPresets,
          oldModulesTracker: modulesTrackerData.oldModulesTracker,
        };
      }

      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardiansReader.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      var playerTierResult = getBatchResult("Player Tier", "values");
      var playerStatResult = getBatchResult("Player Stat", "values");
      var playerPerksResult = getBatchResult("Player Perks", "values");
      if (
        playerTierResult &&
        playerTierResult.values &&
        playerStatResult &&
        playerStatResult.values &&
        playerPerksResult &&
        playerPerksResult.values
      ) {
        var playerTierValues = playerTierResult.values;
        var playerStatValues = playerStatResult.values;
        var playerPerksValues = playerPerksResult.values;
        var playerTierData =
          playerStuffReader.getVersion4_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuffReader.getVersion3_2PlayerStuffStats(playerStatValues);
        var playerPerksData =
          playerStuffReader.getVersion4_2PlayerStuffPerks(playerPerksValues);
        var playerSuccess =
          playerTierData.success &&
          playerStatData.success &&
          playerPerksData.success;
        var playerData = {
          success: playerSuccess,
          message: playerSuccess
            ? "Player & Stuff data retrieved successfully"
            : "Error retrieving Player & Stuff data",
          oldPlayerStuffTierData: playerTierData.oldPlayerStuffTierData,
          oldPlayerStuffStatsData: playerStatData.oldPlayerStuffStatsData,
          oldPerksPreset: playerPerksData.oldPerksPreset,
          shouldRemoveUsedPerks: playerPerksData.shouldRemoveUsedPerks,
        };
        collectedData["Player & Stuff"] = playerData;
      }

      return {
        success: true,
        message: "IDS Collection data retrieved successfully",
        data: collectedData,
      };
    } catch (error) {
      var errorReport = errors.report("IDS", error, {
        note: `Error in IDS Collection version3_2`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads IDS_Collection data from a v3.0.4 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_0_4: function (oldSheetID) {
    try {
      console.log("Called: collectionConvertersV3.version3_0_4");

      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E",
          "Lab Planner": "Lab Planner",
          "Workshop Levels": "EXPORT_WS!B2:M",
          "Workshop Plus": "EXPORT_WS!P2:V",
          "Workshop Ratio": "Desired Ratios",
          "Ultimate Weapon": "EXPORT_UW!C5:H",
          "Themes & Songs": "Themes & Songs",
          Bots: "EXPORT_Bots!C4:L",
          Relics: "Relics",
          "Vault Harmony": "Vault_Harmony",
          "Vault Power": "Vault_Power",
          "Card Preset": "Card Preset",
          "Card Tracker": "Card and Mastery Tracker",
          "Cards Levels": "EXPORT_Cards!B5:D",
          "Cards Slots": "EXPORT_Cards!C2",
          "Modules Inventory": "Modules Inventory",
          "Modules Presets": "Modules Presets",
          "Modules Tracker": "Modules Tracker",
          Guardians: "EXPORT_Guardians!B5:F",
          "Player Tier": "EXPORT_Player!B3:H",
          "Player Stat": "EXPORT_Player!J3:K",
        },
        formulas: {
          "Lab Planner": "Lab Planner",
          "UW Cost Calculator": "UW Cost Calculator v3",
          "Modules Tracker": "Modules Tracker",
        },
      };

      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      var batchValuesResults = [];
      var batchFormulasResults = [];

      if (valuesRanges.length > 0) {
        batchValuesResults = SheetsAPI.batchGetValues(oldSheetID, valuesRanges);
        if (!batchValuesResults) {
          console.log(`Could not read IDS Collection values data`);
          return {
            success: false,
            message: "Could not read IDS Collection values data",
          };
        }
      }

      if (formulasRanges.length > 0) {
        batchFormulasResults = SheetsAPI.batchGetFormulas(
          oldSheetID,
          formulasRanges,
        );
        if (!batchFormulasResults) {
          console.log(`Could not read IDS Collection formulas data`);
          return {
            success: false,
            message: "Could not read IDS Collection formulas data",
          };
        }
      }

      var getBatchResult = function (key, type) {
        type = type || "values";

        if (type === "values") {
          var index = Object.keys(rangeMap.values).indexOf(key);
          return index !== -1 && batchValuesResults[index]
            ? batchValuesResults[index]
            : null;
        } else if (type === "formulas") {
          var index = Object.keys(rangeMap.formulas).indexOf(key);
          return index !== -1 && batchFormulasResults[index]
            ? batchFormulasResults[index]
            : null;
        }
        return null;
      };

      var collectedData = {};

      var labLevelsResult = getBatchResult("Lab Levels", "values");
      var labPlannerValuesResult = getBatchResult("Lab Planner", "values");
      var labPlannerFormulasResult = getBatchResult("Lab Planner", "formulas");
      if (labLevelsResult && labLevelsResult.values) {
        var labLevelsValues = labLevelsResult.values;
        var labPlannerValues =
          labPlannerValuesResult && labPlannerValuesResult.values
            ? labPlannerValuesResult.values
            : null;
        var labPlannerFormulas =
          labPlannerFormulasResult && labPlannerFormulasResult.values
            ? labPlannerFormulasResult.values
            : null;

        var labLevelsData = labReader.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = labReader.getVersion1_0LabPlanner(
          labPlannerValues,
          labPlannerFormulas,
          labLevelsData.oldLabLevels,
          labLevelsData.oldLabMax,
        );

        var labSuccess = labLevelsData.success && labPlannerData.success;
        collectedData.Laboratory = {
          success: labSuccess,
          message: labSuccess
            ? "Laboratory data retrieved successfully"
            : "Error retrieving Laboratory data",
          oldLabLevels: labLevelsData.oldLabLevels,
          oldLabPlanner: labPlannerData.oldLabPlanner,
        };
      }

      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopPlusRatioResult = getBatchResult("Workshop Ratio", "values");
      if (
        workshopLevelsResult &&
        workshopLevelsResult.values &&
        workshopPlusResult &&
        workshopPlusResult.values &&
        workshopPlusRatioResult &&
        workshopPlusRatioResult.values
      ) {
        var workshopLevelsValues = workshopLevelsResult.values;
        var workshopPlusLevelsValues = workshopPlusResult.values;
        var workshopPlusRatioValues = workshopPlusRatioResult.values;

        var workshopLevelsData =
          workshopReader.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshopReader.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshopReader.getVersion2_2_8WorkshopPlusRatios(
          workshopPlusLevelsData.oldWorkshopPlusLevels.presetNames,
          workshopPlusRatioValues,
        );
        var workshopSuccess =
          workshopLevelsData.success &&
          workshopPlusLevelsData.success &&
          workshopPlusRatiosData.success;
        collectedData.Workshop = {
          success: workshopSuccess,
          message: workshopSuccess
            ? "Workshop data retrieved successfully"
            : "Error retrieving Workshop data",
          oldWorkshopLevels: workshopLevelsData.oldWorkshopLevels,
          oldWorkshopPlusLevels: workshopPlusLevelsData.oldWorkshopPlusLevels,
          oldWorkshopPlusRatios: workshopPlusRatiosData.oldWorkshopPlusRatios,
        };
      }

      var ultimateResult = getBatchResult("Ultimate Weapon", "values");
      var ultimateCostCalculatorResult = getBatchResult(
        "UW Cost Calculator",
        "formulas",
      );
      if (
        ultimateResult &&
        ultimateResult.values &&
        ultimateCostCalculatorResult &&
        ultimateCostCalculatorResult.values
      ) {
        var ultimateValues = ultimateResult.values;
        var ultimateCostCalculatorValues = ultimateCostCalculatorResult.values;

        var ultimateWeaponsData =
          ultimateReader.getVersion3_1_1UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimateReader.getVersion1_0CostCalculator(
          ultimateCostCalculatorValues,
        );

        var ultimateSuccess =
          ultimateWeaponsData.success && costCalculatorData.success;
        collectedData["Ultimate Weapon"] = {
          success: ultimateSuccess,
          message: ultimateSuccess
            ? "Ultimate Weapon data retrieved successfully"
            : "Error retrieving Ultimate Weapon data",
          oldUltimate: ultimateWeaponsData.oldUltimate,
          oldUltimateCostCalculator: costCalculatorData["UW Cost Calculator"],
        };
      }

      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themes.getVersion2_1_6Themes(themesResult.values);
        var relicsData = relics.getVersion1_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = botsReader.getVersion3_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      var harmonyResult = getBatchResult("Vault Harmony", "values");
      var powerResult = getBatchResult("Vault Power", "values");
      if (
        harmonyResult &&
        harmonyResult.values &&
        powerResult &&
        powerResult.values
      ) {
        var harmonyValues = harmonyResult.values;
        var powerValues = powerResult.values;

        var harmonyVaultData = vaultReader.getVersion3_1Vault(harmonyValues);
        var powerVaultData = vaultReader.getVersion3_1Vault(powerValues);

        var vaultSuccess = harmonyVaultData.success && powerVaultData.success;
        collectedData.Vault = {
          success: vaultSuccess,
          message: vaultSuccess
            ? "Vault data retrieved successfully"
            : "Error retrieving Vault data",
          oldVaultHarmony: harmonyVaultData.oldVault,
          oldVaultPower: powerVaultData.oldVault,
        };
      }

      var cardsPresetResult = getBatchResult("Card Preset", "values");
      var cardsTrackerResult = getBatchResult("Card Tracker", "values");
      var cardsLevelsResult = getBatchResult("Cards Levels", "values");
      var cardsSlotsResult = getBatchResult("Cards Slots", "values");
      if (
        cardsPresetResult &&
        cardsPresetResult.values &&
        cardsTrackerResult &&
        cardsTrackerResult.values &&
        cardsLevelsResult &&
        cardsLevelsResult.values &&
        cardsSlotsResult &&
        cardsSlotsResult.values
      ) {
        var cardsPresetValues = cardsPresetResult.values;
        var cardsTrackerValues = cardsTrackerResult.values;
        var cardsLevelValues = cardsLevelsResult.values;
        var cardsSlotsValues = cardsSlotsResult.values;

        var cardsPresetData = cardsReader.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cardsReader.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cardsReader.getVersion1_0CardsTracker(cardsTrackerValues);

        var cardsSuccess =
          cardsPresetData.success &&
          cardsLevelData.success &&
          cardsTrackerData.success;
        collectedData.Cards = {
          success: cardsSuccess,
          message: cardsSuccess
            ? "Cards data retrieved successfully"
            : "Error retrieving Cards data",
          oldCardsPreset: cardsPresetData.oldCardsPreset,
          shouldRemoveUsedCards: cardsPresetData.shouldRemoveUsedCards,
          oldCardsLevel: cardsLevelData.oldCardsLevel,
          oldCardSlots: cardsLevelData.oldCardSlots,
          oldCardsTracker: cardsTrackerData.oldCardsTracker,
        };
      }

      var modulesInventoryResult = getBatchResult(
        "Modules Inventory",
        "values",
      );
      var modulesPresetsResult = getBatchResult("Modules Presets", "values");
      var modulesTrackerResult = getBatchResult("Modules Tracker", "values");
      var modulesTrackerFormulasResult = getBatchResult(
        "Modules Tracker",
        "formulas",
      );
      if (
        modulesInventoryResult &&
        modulesInventoryResult.values &&
        modulesPresetsResult &&
        modulesPresetsResult.values &&
        modulesTrackerResult &&
        modulesTrackerResult.values &&
        modulesTrackerFormulasResult &&
        modulesTrackerFormulasResult.values
      ) {
        var modulesInventoryValues = modulesInventoryResult.values;
        var modulesPresetsValues = modulesPresetsResult.values;
        var modulesTrackerValues = modulesTrackerResult.values;
        var modulesTrackerFormulas = modulesTrackerFormulasResult.values;
        var modulesInventoryData = modulesInventoryReader.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modulesPresetsReader.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modulesTrackerReader.getVersion4_7ModulesTracker(
          modulesTrackerValues,
          modulesTrackerFormulas,
        );
        var modulesSuccess =
          modulesInventoryData.success &&
          modulesPresetsData.success &&
          modulesTrackerData.success;
        collectedData.Modules = {
          success: modulesSuccess,
          message: modulesSuccess
            ? "Modules data retrieved successfully"
            : "Error retrieving Modules data",
          oldModulesInventory: modulesInventoryData.oldModulesInventory,
          oldModulesPresets: modulesPresetsData.oldModulesPresets,
          oldModulesTracker: modulesTrackerData.oldModulesTracker,
        };
      }

      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardiansReader.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      var playerTierResult = getBatchResult("Player Tier", "values");
      var playerStatResult = getBatchResult("Player Stat", "values");
      if (
        playerTierResult &&
        playerTierResult.values &&
        playerStatResult &&
        playerStatResult.values
      ) {
        var playerTierValues = playerTierResult.values;
        var playerStatValues = playerStatResult.values;
        var playerTierData =
          playerStuffReader.getVersion4_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuffReader.getVersion3_2PlayerStuffStats(playerStatValues);
        var playerSuccess = playerTierData.success && playerStatData.success;
        var playerData = {
          success: playerSuccess,
          message: playerSuccess
            ? "Player & Stuff data retrieved successfully"
            : "Error retrieving Player & Stuff data",
          oldPlayerStuffTierData: playerTierData.oldPlayerStuffTierData,
          oldPlayerStuffStatsData: playerStatData.oldPlayerStuffStatsData,
        };
        collectedData["Player & Stuff"] = playerData;
      }

      return {
        success: true,
        message: "IDS Collection data retrieved successfully",
        data: collectedData,
      };
    } catch (error) {
      var errorReport = errors.report("IDS", error, {
        note: `Error in IDS Collection version3_0_4`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Reads IDS_Collection data from a v3.0 sheet.
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  version3_0: function (oldSheetID) {
    try {
      console.log("Called: collectionConvertersV3.version3_0");

      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E",
          "Lab Planner": "Lab Planner",
          "Workshop Levels": "EXPORT_WS!B2:M",
          "Workshop Plus": "EXPORT_WS!P2:V",
          "Workshop Ratio": "Desired Ratios",
          "Ultimate Weapon": "EXPORT_UW!C5:G",
          "Themes & Songs": "Themes & Songs",
          Bots: "EXPORT_Bots!C4:L",
          Relics: "Relics",
          "Vault Harmony": "Vault_Harmony",
          "Vault Power": "Vault_Power",
          "Card Preset": "Card Preset",
          "Card Tracker": "Card and Mastery Tracker",
          "Cards Levels": "EXPORT_Cards!B5:D",
          "Cards Slots": "EXPORT_Cards!C2",
          "Modules Inventory": "Modules Inventory",
          "Modules Presets": "Modules Presets",
          "Modules Tracker": "Modules Tracker",
          Guardians: "EXPORT_Guardians!B5:F",
          "Player Tier": "EXPORT_Player!B3:H",
          "Player Stat": "EXPORT_Player!J3:K",
        },
        formulas: {
          "Lab Planner": "Lab Planner",
          "UW Cost Calculator": "UW Cost Calculator v3",
          "Modules Tracker": "Modules Tracker",
        },
      };

      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      var batchValuesResults = [];
      var batchFormulasResults = [];

      if (valuesRanges.length > 0) {
        batchValuesResults = SheetsAPI.batchGetValues(oldSheetID, valuesRanges);
        if (!batchValuesResults) {
          console.log(`Could not read IDS Collection values data`);
          return {
            success: false,
            message: "Could not read IDS Collection values data",
          };
        }
      }

      if (formulasRanges.length > 0) {
        batchFormulasResults = SheetsAPI.batchGetFormulas(
          oldSheetID,
          formulasRanges,
        );
        if (!batchFormulasResults) {
          console.log(`Could not read IDS Collection formulas data`);
          return {
            success: false,
            message: "Could not read IDS Collection formulas data",
          };
        }
      }

      var getBatchResult = function (key, type) {
        type = type || "values";

        if (type === "values") {
          var index = Object.keys(rangeMap.values).indexOf(key);
          return index !== -1 && batchValuesResults[index]
            ? batchValuesResults[index]
            : null;
        } else if (type === "formulas") {
          var index = Object.keys(rangeMap.formulas).indexOf(key);
          return index !== -1 && batchFormulasResults[index]
            ? batchFormulasResults[index]
            : null;
        }
        return null;
      };

      var collectedData = {};

      var labLevelsResult = getBatchResult("Lab Levels", "values");
      var labPlannerValuesResult = getBatchResult("Lab Planner", "values");
      var labPlannerFormulasResult = getBatchResult("Lab Planner", "formulas");
      if (labLevelsResult && labLevelsResult.values) {
        var labLevelsValues = labLevelsResult.values;
        var labPlannerValues =
          labPlannerValuesResult && labPlannerValuesResult.values
            ? labPlannerValuesResult.values
            : null;
        var labPlannerFormulas =
          labPlannerFormulasResult && labPlannerFormulasResult.values
            ? labPlannerFormulasResult.values
            : null;

        var labLevelsData = labReader.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = labReader.getVersion1_0LabPlanner(
          labPlannerValues,
          labPlannerFormulas,
          labLevelsData.oldLabLevels,
          labLevelsData.oldLabMax,
        );

        var labSuccess = labLevelsData.success && labPlannerData.success;
        collectedData.Laboratory = {
          success: labSuccess,
          message: labSuccess
            ? "Laboratory data retrieved successfully"
            : "Error retrieving Laboratory data",
          oldLabLevels: labLevelsData.oldLabLevels,
          oldLabPlanner: labPlannerData.oldLabPlanner,
        };
      }

      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopPlusRatioResult = getBatchResult("Workshop Ratio", "values");
      if (
        workshopLevelsResult &&
        workshopLevelsResult.values &&
        workshopPlusResult &&
        workshopPlusResult.values &&
        workshopPlusRatioResult &&
        workshopPlusRatioResult.values
      ) {
        var workshopLevelsValues = workshopLevelsResult.values;
        var workshopPlusLevelsValues = workshopPlusResult.values;
        var workshopPlusRatioValues = workshopPlusRatioResult.values;

        var workshopLevelsData =
          workshopReader.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshopReader.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshopReader.getVersion2_2_8WorkshopPlusRatios(
          workshopPlusLevelsData.oldWorkshopPlusLevels.presetNames,
          workshopPlusRatioValues,
        );
        var workshopSuccess =
          workshopLevelsData.success &&
          workshopPlusLevelsData.success &&
          workshopPlusRatiosData.success;
        collectedData.Workshop = {
          success: workshopSuccess,
          message: workshopSuccess
            ? "Workshop data retrieved successfully"
            : "Error retrieving Workshop data",
          oldWorkshopLevels: workshopLevelsData.oldWorkshopLevels,
          oldWorkshopPlusLevels: workshopPlusLevelsData.oldWorkshopPlusLevels,
          oldWorkshopPlusRatios: workshopPlusRatiosData.oldWorkshopPlusRatios,
        };
      }

      var ultimateResult = getBatchResult("Ultimate Weapon", "values");
      var ultimateCostCalculatorResult = getBatchResult(
        "UW Cost Calculator",
        "formulas",
      );
      if (
        ultimateResult &&
        ultimateResult.values &&
        ultimateCostCalculatorResult &&
        ultimateCostCalculatorResult.values
      ) {
        var ultimateValues = ultimateResult.values;
        var ultimateCostCalculatorValues = ultimateCostCalculatorResult.values;

        var ultimateWeaponsData =
          ultimateReader.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimateReader.getVersion1_0CostCalculator(
          ultimateCostCalculatorValues,
        );

        var ultimateSuccess =
          ultimateWeaponsData.success && costCalculatorData.success;
        collectedData["Ultimate Weapon"] = {
          success: ultimateSuccess,
          message: ultimateSuccess
            ? "Ultimate Weapon data retrieved successfully"
            : "Error retrieving Ultimate Weapon data",
          oldUltimate: ultimateWeaponsData.oldUltimate,
          oldUltimateCostCalculator: costCalculatorData["UW Cost Calculator"],
        };
      }

      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themes.getVersion2_1_6Themes(themesResult.values);
        var relicsData = relics.getVersion1_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = botsReader.getVersion3_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      var harmonyResult = getBatchResult("Vault Harmony", "values");
      var powerResult = getBatchResult("Vault Power", "values");
      if (
        harmonyResult &&
        harmonyResult.values &&
        powerResult &&
        powerResult.values
      ) {
        var harmonyValues = harmonyResult.values;
        var powerValues = powerResult.values;

        var harmonyVaultData = vaultReader.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vaultReader.getVersion1_0Vault(powerValues);

        var vaultSuccess = harmonyVaultData.success && powerVaultData.success;
        collectedData.Vault = {
          success: vaultSuccess,
          message: vaultSuccess
            ? "Vault data retrieved successfully"
            : "Error retrieving Vault data",
          oldVaultHarmony: harmonyVaultData.oldVault,
          oldVaultPower: powerVaultData.oldVault,
        };
      }

      var cardsPresetResult = getBatchResult("Card Preset", "values");
      var cardsTrackerResult = getBatchResult("Card Tracker", "values");
      var cardsLevelsResult = getBatchResult("Cards Levels", "values");
      var cardsSlotsResult = getBatchResult("Cards Slots", "values");
      if (
        cardsPresetResult &&
        cardsPresetResult.values &&
        cardsTrackerResult &&
        cardsTrackerResult.values &&
        cardsLevelsResult &&
        cardsLevelsResult.values &&
        cardsSlotsResult &&
        cardsSlotsResult.values
      ) {
        var cardsPresetValues = cardsPresetResult.values;
        var cardsTrackerValues = cardsTrackerResult.values;
        var cardsLevelValues = cardsLevelsResult.values;
        var cardsSlotsValues = cardsSlotsResult.values;

        var cardsPresetData = cardsReader.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cardsReader.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cardsReader.getVersion1_0CardsTracker(cardsTrackerValues);

        var cardsSuccess =
          cardsPresetData.success &&
          cardsLevelData.success &&
          cardsTrackerData.success;
        collectedData.Cards = {
          success: cardsSuccess,
          message: cardsSuccess
            ? "Cards data retrieved successfully"
            : "Error retrieving Cards data",
          oldCardsPreset: cardsPresetData.oldCardsPreset,
          shouldRemoveUsedCards: cardsPresetData.shouldRemoveUsedCards,
          oldCardsLevel: cardsLevelData.oldCardsLevel,
          oldCardSlots: cardsLevelData.oldCardSlots,
          oldCardsTracker: cardsTrackerData.oldCardsTracker,
        };
      }

      var modulesInventoryResult = getBatchResult(
        "Modules Inventory",
        "values",
      );
      var modulesPresetsResult = getBatchResult("Modules Presets", "values");
      var modulesTrackerResult = getBatchResult("Modules Tracker", "values");
      var modulesTrackerFormulasResult = getBatchResult(
        "Modules Tracker",
        "formulas",
      );
      if (
        modulesInventoryResult &&
        modulesInventoryResult.values &&
        modulesPresetsResult &&
        modulesPresetsResult.values &&
        modulesTrackerResult &&
        modulesTrackerResult.values &&
        modulesTrackerFormulasResult &&
        modulesTrackerFormulasResult.values
      ) {
        var modulesInventoryValues = modulesInventoryResult.values;
        var modulesPresetsValues = modulesPresetsResult.values;
        var modulesTrackerValues = modulesTrackerResult.values;
        var modulesTrackerFormulas = modulesTrackerFormulasResult.values;
        var modulesInventoryData = modulesInventoryReader.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modulesPresetsReader.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modulesTrackerReader.getVersion4_7ModulesTracker(
          modulesTrackerValues,
          modulesTrackerFormulas,
        );
        var modulesSuccess =
          modulesInventoryData.success &&
          modulesPresetsData.success &&
          modulesTrackerData.success;
        collectedData.Modules = {
          success: modulesSuccess,
          message: modulesSuccess
            ? "Modules data retrieved successfully"
            : "Error retrieving Modules data",
          oldModulesInventory: modulesInventoryData.oldModulesInventory,
          oldModulesPresets: modulesPresetsData.oldModulesPresets,
          oldModulesTracker: modulesTrackerData.oldModulesTracker,
        };
      }

      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardiansReader.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      var playerTierResult = getBatchResult("Player Tier", "values");
      var playerStatResult = getBatchResult("Player Stat", "values");
      if (
        playerTierResult &&
        playerTierResult.values &&
        playerStatResult &&
        playerStatResult.values
      ) {
        var playerTierValues = playerTierResult.values;
        var playerStatValues = playerStatResult.values;
        var playerTierData =
          playerStuffReader.getVersion4_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuffReader.getVersion3_2PlayerStuffStats(playerStatValues);
        var playerSuccess = playerTierData.success && playerStatData.success;
        var playerData = {
          success: playerSuccess,
          message: playerSuccess
            ? "Player & Stuff data retrieved successfully"
            : "Error retrieving Player & Stuff data",
          oldPlayerStuffTierData: playerTierData.oldPlayerStuffTierData,
          oldPlayerStuffStatsData: playerStatData.oldPlayerStuffStatsData,
        };
        collectedData["Player & Stuff"] = playerData;
      }

      return {
        success: true,
        message: "IDS Collection data retrieved successfully",
        data: collectedData,
      };
    } catch (error) {
      var errorReport = errors.report("IDS", error, {
        note: `Error in IDS Collection version3_0`,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },
};
