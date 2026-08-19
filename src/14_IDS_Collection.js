const collection = {
  // #region Export Functions
  exportData: function (versionDifference) {
    try {
      console.log("Called: collection.exportData");
      var getVersionFunction = this.convertVersionFunctions[versionDifference];
      if (!getVersionFunction) {
        console.log(`Unsupported version: ${versionDifference}`);
        return {
          success: false,
          message: `Unsupported version: ${versionDifference}`,
        };
      }

      var oldDataResult = getVersionFunction();
      if (!oldDataResult || !oldDataResult.success) {
        console.log(`${oldDataResult.message}`);
        return oldDataResult;
      }

      return {
        success: true,
        message: "IDS Collection export completed successfully",
        data: oldDataResult.data,
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting IDS Collection data: " + error.message,
      };
    }
  },

  // #endregion
  // #region Import Functions
  importData: function (data) {
    try {
      console.log("Called: collection.importData");
      var newSpreadsheet = spreadsheets("IDS Collection newSpreadsheet");
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }
      var newSheetID = newSpreadsheet.spreadsheetId;

      // Define DVT named ranges for each Ultimate Weapon, Bot, Module, and Guardian
      var dvtNamedRangesUW = {
        "Chain Lightning": {
          Damage: "DVT_UW_UG_CL_DMG",
          Quantity: "DVT_UW_UG_CL_QNT",
          Chance: "DVT_UW_UG_CL_CH",
          Smite: "DVT_UW_UG_CL_SM",
        },
        "Smart Missiles": {
          Damage: "DVT_UW_UG_SM_DMG",
          Quantity: "DVT_UW_UG_SM_QNT",
          Cooldown: "DVT_UW_UG_SM_CD",
          "Cover Fire": "DVT_UW_UG_SM_CF",
        },
        "Death Wave": {
          Damage: "DVT_UW_UG_DW_DMG",
          Quantity: "DVT_UW_UG_DW_QNT",
          Cooldown: "DVT_UW_UG_DW_CD",
          "Kill Wall": "DVT_UW_UG_DW_KW",
        },
        "Chrono Field": {
          Duration: "DVT_UW_UG_CF_DU",
          "Speed Reduction": "DVT_UW_UG_CF_SP",
          Cooldown: "DVT_UW_UG_CF_CD",
          "Chrono Loop": "DVT_UW_UG_CF_CL",
        },
        "Inner Land Mines": {
          Damage: "DVT_UW_UG_ILM_DMG",
          Quantity: "DVT_UW_UG_ILM_QNT",
          Cooldown: "DVT_UW_UG_ILM_CD",
          "Charged Mines": "DVT_UW_UG_ILM_CM",
        },
        "Golden Tower": {
          Multiplier: "DVT_UW_UG_GT_M",
          Duration: "DVT_UW_UG_GT_DU",
          Cooldown: "DVT_UW_UG_GT_CD",
          "Golden Combo": "DVT_UW_UG_GT_GC",
        },
        "Poison Swamp": {
          Damage: "DVT_UW_UG_PS_DMG",
          Duration: "DVT_UW_UG_PS_DU",
          Cooldown: "DVT_UW_UG_PS_CH",
          "Death Creep": "DVT_UW_UG_PS_DC",
        },
        "Black Hole": {
          Size: "DVT_UW_UG_BH_SZ",
          Duration: "DVT_UW_UG_BH_DU",
          Cooldown: "DVT_UW_UG_BH_CD",
          Consume: "DVT_UW_UG_BH_C",
        },
        Spotlight: {
          Multiplier: "DVT_UW_UG_SL_MU",
          Angle: "DVT_UW_UG_SL_AN",
          Quantity: "DVT_UW_UG_SL_QNT",
          "Light Range": "DVT_UW_UG_SL_LR",
        },
      };

      var dvtNamedRangesBots = {
        "Flame Bot": {
          "Damage R.": "DVT_BOT_UG_FB_DMGR",
          Cooldown: "DVT_BOT_UG_FB_CD",
          Damage: "DVT_BOT_UG_FB_DMG",
          Range: "DVT_BOT_UG_FB_RANGE",
          Wildfire: "DVT_BOT_UG_FB_WILDFIRE",
        },
        "Thunder Bot": {
          Duration: "DVT_BOT_UG_TB_DUR",
          Cooldown: "DVT_BOT_UG_TB_CD",
          Linger: "DVT_BOT_UG_TB_LINGER",
          Range: "DVT_BOT_UG_TB_RANGE",
          "Titan Shock": "DVT_BOT_UG_TB_TITANSHOCK",
        },
        "Golden Bot": {
          Duration: "DVT_BOT_UG_GB_DUR",
          Cooldown: "DVT_BOT_UG_GB_CD",
          Bonus: "DVT_BOT_UG_GB_BONUS",
          Range: "DVT_BOT_UG_GB_RANGE",
          "Bonus Cell": "DVT_BOT_UG_GB_BONUSCELL",
        },
        "Amplify Bot": {
          Duration: "DVT_BOT_UG_AB_DUR",
          Cooldown: "DVT_BOT_UG_AB_CD",
          Bonus: "DVT_BOT_UG_AB_BONUS",
          Range: "DVT_BOT_UG_AB_RANGE",
          "Echoing Shot": "DVT_BOT_UG_AB_ECHOINGSHOT",
        },
        "Bot Bot": {
          Duration: "DVT_BOT_UG_BB_DUR",
          Cooldown: "DVT_BOT_UG_BB_CD",
          Bonus: "DVT_BOT_UG_BB_BONUS",
          Range: "DVT_BOT_UG_BB_RANGE",
          "Maximum Power": "DVT_BOT_UG_BB_MAXIMUMPOWER",
        },
      };

      var dvtNamedRangesModules = {
        "Main Efficiency": "DVT_Mod_Assist_Bonus_Level",
        "Substat Efficiency": "DVT_Mod_Assist_Substat_Level",
      };

      var dvtNamedRangesGuardians = {
        Attack: {
          Percentage: "DVT_GAR_UG_AT_PER",
          Cooldown: "DVT_GAR_UG_AT_COO",
          Targets: "DVT_GAR_UG_AT_TAR",
        },
        Ally: {
          "Recovery Amount": "DVT_GAR_UG_AL_REC",
          "Max Recovery": "DVT_GAR_UG_AL_MAX",
          Cooldown: "DVT_GAR_UG_AL_COO",
        },
        Bounty: {
          Multiplier: "DVT_GAR_UG_BO_MUL",
          Cooldown: "DVT_GAR_UG_BO_COO",
          Targets: "DVT_GAR_UG_BO_TAR",
        },
        Fetch: {
          Cooldown: "DVT_GAR_UG_FE_COO",
          "Find Chance": "DVT_GAR_UG_FE_FIN",
          "Double Find Chance": "DVT_GAR_UG_FE_DOU",
        },
        Summon: {
          Cooldown: "DVT_GAR_UG_SU_COO",
          Duration: "DVT_GAR_UG_SU_DUR",
          "Cash Bonus": "DVT_GAR_UG_SU_CAS",
        },
        Scout: {
          Cooldown: "DVT_GAR_UG_SC_COO",
          "Range Bonus": "DVT_GAR_UG_SC_RAN",
          Duration: "DVT_GAR_UG_SC_DUR",
        },
      };

      // Configuration dictionary mapping sheet types to their required ranges
      var sheetRequiredRanges = {
        values: {
          "Home Page": { sheetName: "Home Page", range: "Home Page" },
          Lab_MS: { sheetName: "Lab_MS", range: "Lab_MS" },
          "Workshop Ratio": {
            sheetName: "Desired Ratios",
            range: "Desired Ratios",
          },
          UW_MS: { sheetName: "UW_MS", range: "UW_MS" },
          "Themes & Songs": {
            sheetName: "Themes & Songs",
            range: "Themes & Songs",
          },
          Bots_MS: { sheetName: "Bots_MS", range: "Bots_MS" },
          Relics: { sheetName: "Relics", range: "Relics" },
          Vault_Harmony: { sheetName: "Vault_Harmony", range: "Vault_Harmony" },
          Vault_Power: { sheetName: "Vault_Power", range: "Vault_Power" },
          Cards_MS: { sheetName: "Cards_MS", range: "Cards_MS" },
          "Card Preset": { sheetName: "Card Preset", range: "Card Preset" },
          Cards_Tracker: {
            sheetName: "Card and Mastery Tracker",
            range: "Card and Mastery Tracker",
          },
          "Modules Inventory": {
            sheetName: "Modules Inventory",
            range: "Modules Inventory",
          },
          "Modules Presets": {
            sheetName: "Modules Presets",
            range: "Modules Presets",
          },
          "Modules Tracker": {
            sheetName: "Modules Tracker",
            range: "Modules Tracker",
          },
          Guardians_MS: { sheetName: "Guardians_MS", range: "Guardians_MS" },
          player_MS: { sheetName: "player_MS", range: "player_MS" },
          "Perk Preset": { sheetName: "Perk Preset", range: "Perk Preset" },
        },
        formulas: {
          "Lab Planner": { sheetName: "Lab Planner", range: "Lab Planner" },
          Workshop_MS: { sheetName: "Workshop_MS", range: "Workshop_MS" },
          "UW Cost Calculator": {
            sheetName: "UW Cost Calculator v3",
            range: "UW Cost Calculator v3",
          },
        },
      };

      // Add Ultimate Weapons DVT ranges to sheetRequiredRanges
      Object.keys(dvtNamedRangesUW).forEach(function (weapon) {
        Object.keys(dvtNamedRangesUW[weapon]).forEach(function (prop) {
          var rangeName = dvtNamedRangesUW[weapon][prop];
          sheetRequiredRanges.values[rangeName] = {
            sheetName: rangeName,
            range: rangeName,
          };
        });
      });

      // Add Bots DVT ranges to sheetRequiredRanges
      Object.keys(dvtNamedRangesBots).forEach(function (bot) {
        Object.keys(dvtNamedRangesBots[bot]).forEach(function (prop) {
          var rangeName = dvtNamedRangesBots[bot][prop];
          sheetRequiredRanges.values[rangeName] = {
            sheetName: rangeName,
            range: rangeName,
          };
        });
      });

      // Add Modules DVT ranges to sheetRequiredRanges
      Object.keys(dvtNamedRangesModules).forEach(function (item) {
        var rangeName = dvtNamedRangesModules[item];
        sheetRequiredRanges.values[rangeName] = {
          sheetName: rangeName,
          range: rangeName,
        };
      });

      // Add Guardians DVT ranges to sheetRequiredRanges
      Object.keys(dvtNamedRangesGuardians).forEach(function (guardian) {
        Object.keys(dvtNamedRangesGuardians[guardian]).forEach(function (prop) {
          var rangeName = dvtNamedRangesGuardians[guardian][prop];
          sheetRequiredRanges.values[rangeName] = {
            sheetName: rangeName,
            range: rangeName,
          };
        });
      });

      // Create ranges arrays from the configuration
      var allValuesRanges = Object.keys(sheetRequiredRanges.values).map(
        function (key) {
          return sheetRequiredRanges.values[key].range;
        },
      );

      var allFormulasRanges = Object.keys(sheetRequiredRanges.formulas).map(
        function (key) {
          return sheetRequiredRanges.formulas[key].range;
        },
      );

      // Batch fetch all required data
      var batchValuesResults = [];
      var batchFormulasResults = [];

      if (allValuesRanges.length > 0) {
        batchValuesResults = SheetsAPI.batchGetValues(
          newSheetID,
          allValuesRanges,
        );
        if (!batchValuesResults) {
          console.log(
            `Error fetching values ranges from IDS Collection spreadsheet`,
          );
          return {
            success: false,
            message:
              "Error fetching values ranges from IDS Collection spreadsheet",
          };
        }
      }

      if (allFormulasRanges.length > 0) {
        batchFormulasResults = SheetsAPI.batchGetFormulas(
          newSheetID,
          allFormulasRanges,
        );
        if (!batchFormulasResults) {
          console.log(
            `Error fetching formulas ranges from IDS Collection spreadsheet`,
          );
          return {
            success: false,
            message:
              "Error fetching formulas ranges from IDS Collection spreadsheet",
          };
        }
      }

      // Update each sheet type with its data
      var updateResults = [];
      var batchUpdate = [];

      // Helper function to get range data
      var getRangeData = function (sheetName, type) {
        type = type || "values"; // Default to values if not specified

        if (type === "formulas") {
          var index = Object.keys(sheetRequiredRanges.formulas).indexOf(
            sheetName,
          );
          return index !== -1 && batchFormulasResults[index]
            ? batchFormulasResults[index].values
            : null;
        } else {
          var index = Object.keys(sheetRequiredRanges.values).indexOf(
            sheetName,
          );
          return index !== -1 && batchValuesResults[index]
            ? batchValuesResults[index].values
            : null;
        }
      };

      // Build DVT named ranges data for each module
      var buildDVTNamedRangesData = function (dvtNamedRanges) {
        var dvtNamedRangesData = {};
        Object.keys(dvtNamedRanges).forEach(function (item) {
          var value = dvtNamedRanges[item];
          if (typeof value === "object" && value !== null) {
            // Nested structure: { item: { prop: rangeName } }
            dvtNamedRangesData[item] = {};
            Object.keys(value).forEach(function (prop) {
              var rangeData = getRangeData(value[prop], "values");
              dvtNamedRangesData[item][prop] = rangeData || [];
            });
          } else {
            // Flat structure: { item: rangeName }
            var rangeData = getRangeData(value, "values");
            dvtNamedRangesData[item] = rangeData || [];
          }
        });
        return dvtNamedRangesData;
      };

      // Laboratory updates
      if (data.Laboratory) {
        try {
          var labData = data.Laboratory;
          var labMasterSheetData = getRangeData("Lab_MS", "values");
          var labPlannerData = getRangeData("Lab Planner", "formulas");
          var labSuccess = true;
          var labMessages = [];
          var labResult, labPlannerResult;
          if (labData.hasOwnProperty("oldLabLevels") && labMasterSheetData) {
            labResult = lab.updateLabLevels(
              sheetRequiredRanges.values["Lab_MS"].sheetName,
              labData.oldLabLevels,
              labMasterSheetData,
            );
            if (labResult && labResult.success) {
              batchUpdate = batchUpdate.concat(labResult.batchUpdate || []);
            } else {
              labSuccess = false;
              labMessages.push(
                labResult ? labResult.message : "Unknown error in LabLevels",
              );
            }
          }
          if (labData.hasOwnProperty("oldLabPlanner") && labPlannerData) {
            labPlannerResult = lab.updateLabPlanner(
              sheetRequiredRanges.formulas["Lab Planner"].sheetName,
              labData.oldLabPlanner,
              labPlannerData,
            );
            if (labPlannerResult && labPlannerResult.success) {
              batchUpdate = batchUpdate.concat(
                labPlannerResult.batchUpdate || [],
              );
            } else {
              labSuccess = false;
              labMessages.push(
                labPlannerResult
                  ? labPlannerResult.message
                  : "Unknown error in LabPlanner",
              );
            }
          }
          updateResults.push({
            sheetType: "Laboratory",
            success: labSuccess,
            message: labSuccess
              ? "Laboratory updated successfully"
              : "Laboratory update failed: " + labMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Laboratory",
            success: false,
            message: "Error in Laboratory update: " + error.message,
          });
        }
      }

      // Workshop updates
      if (data.Workshop) {
        try {
          var workshopData = data.Workshop;
          var workshopMasterSheetData = getRangeData("Workshop_MS", "formulas");
          var workshopPlusRatioData = getRangeData("Workshop Ratio", "values");
          var workshopSuccess = true;
          var workshopMessages = [];
          var workshopResult, workshopPlusRatioResult;
          if (
            workshopData.hasOwnProperty("oldWorkshopLevels") &&
            workshopData.hasOwnProperty("oldWorkshopPlusLevels") &&
            workshopMasterSheetData
          ) {
            var hasPresets = workshopData.hasOwnProperty("hasPresets") ? workshopData.hasPresets : true;
            workshopResult = workshop.updateWorkshopLevels(
              sheetRequiredRanges.formulas["Workshop_MS"].sheetName,
              workshopData.oldWorkshopLevels,
              workshopData.oldWorkshopPlusLevels,
              hasPresets,
              workshopMasterSheetData,
            );
            if (workshopResult && workshopResult.success) {
              batchUpdate = batchUpdate.concat(
                workshopResult.batchUpdate || [],
              );
            } else {
              workshopSuccess = false;
              workshopMessages.push(
                workshopResult
                  ? workshopResult.message
                  : "Unknown error in WorkshopLevels",
              );
            }
          }
          if (workshopData.hasOwnProperty("oldWorkshopPlusRatios")) {
            workshopPlusRatioResult = workshop.updateWorkshopPlusRatios(
              sheetRequiredRanges.values["Workshop Ratio"].sheetName,
              workshopData.oldWorkshopPlusRatios,
              workshopPlusRatioData,
            );
            if (workshopPlusRatioResult && workshopPlusRatioResult.success) {
              batchUpdate = batchUpdate.concat(
                workshopPlusRatioResult.batchUpdate || [],
              );
            } else {
              workshopSuccess = false;
              workshopMessages.push(
                workshopPlusRatioResult
                  ? workshopPlusRatioResult.message
                  : "Unknown error in WorkshopPlusRatios",
              );
            }
          }
          updateResults.push({
            sheetType: "Workshop",
            success: workshopSuccess,
            message: workshopSuccess
              ? "Workshop updated successfully"
              : "Workshop update failed: " + workshopMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Workshop",
            success: false,
            message: "Error in Workshop update: " + error.message,
          });
        }
      }

      // Ultimate Weapon updates
      if (data["Ultimate Weapon"]) {
        try {
          var ultimateData = data["Ultimate Weapon"];
          var ultimateMasterSheetData = getRangeData("UW_MS", "values");
          var ultimateCostCalculatorData = getRangeData(
            "UW Cost Calculator",
            "formulas",
          );
          var ultimateDVTData = buildDVTNamedRangesData(dvtNamedRangesUW);
          var ultimateSuccess = true;
          var ultimateMessages = [];
          var ultimateResult, ultimateCostCalculatorResult;
          if (
            ultimateData.hasOwnProperty("oldUltimate") &&
            ultimateMasterSheetData &&
            ultimateDVTData
          ) {
            ultimateResult = ultimate.updateUltimateLevels(
              sheetRequiredRanges.values["UW_MS"].sheetName,
              ultimateData.oldUltimate,
              ultimateMasterSheetData,
              ultimateDVTData,
            );
            if (ultimateResult && ultimateResult.success) {
              batchUpdate = batchUpdate.concat(
                ultimateResult.batchUpdate || [],
              );
            } else {
              ultimateSuccess = false;
              ultimateMessages.push(
                ultimateResult
                  ? ultimateResult.message
                  : "Unknown error in UltimateLevels",
              );
            }
          }
          if (
            ultimateData.hasOwnProperty("oldUltimateCostCalculator") &&
            ultimateCostCalculatorData
          ) {
            ultimateCostCalculatorResult =
              ultimate.updateUltimateCostCalculator(
                sheetRequiredRanges.formulas["UW Cost Calculator"].sheetName,
                ultimateData.oldUltimateCostCalculator,
                ultimateCostCalculatorData,
              );
            if (
              ultimateCostCalculatorResult &&
              ultimateCostCalculatorResult.success
            ) {
              batchUpdate = batchUpdate.concat(
                ultimateCostCalculatorResult.batchUpdate || [],
              );
            } else {
              ultimateSuccess = false;
              ultimateMessages.push(
                ultimateCostCalculatorResult
                  ? ultimateCostCalculatorResult.message
                  : "Unknown error in UltimateCostCalculator",
              );
            }
          }
          updateResults.push({
            sheetType: "Ultimate Weapon",
            success: ultimateSuccess,
            message: ultimateSuccess
              ? "Ultimate Weapon updated successfully"
              : "Ultimate Weapon update failed: " + ultimateMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Ultimate Weapon",
            success: false,
            message: "Error in Ultimate Weapon update: " + error.message,
          });
        }
      }

      // Themes, Songs & Relics updates - one sheet type, but still written to
      // the two separate sheets the IDS Collection keeps them in.
      if (data["Themes, Songs & Relics"]) {
        try {
          var themesAndRelicsData = data["Themes, Songs & Relics"];
          var themesMasterSheetData = getRangeData("Themes & Songs", "values");
          var relicsMasterSheetData = getRangeData("Relics", "values");
          var themesAndRelicsSuccess = true;
          var themesAndRelicsMessages = [];
          var themesResult;
          var relicsResult;
          if (
            themesAndRelicsData.hasOwnProperty("oldThemesNames") &&
            themesMasterSheetData
          ) {
            themesResult = themesAndRelics.updateThemes(
              sheetRequiredRanges.values["Themes & Songs"].sheetName,
              themesAndRelicsData.oldThemesNames,
              themesMasterSheetData,
            );
            if (themesResult && themesResult.success) {
              batchUpdate = batchUpdate.concat(themesResult.batchUpdate || []);
            } else {
              themesAndRelicsSuccess = false;
              themesAndRelicsMessages.push(
                themesResult ? themesResult.message : "Unknown error in Themes",
              );
            }
          }
          if (
            themesAndRelicsData.hasOwnProperty("oldRelics") &&
            relicsMasterSheetData
          ) {
            relicsResult = themesAndRelics.updateRelics(
              sheetRequiredRanges.values["Relics"].sheetName,
              themesAndRelicsData.oldRelics,
              relicsMasterSheetData,
            );
            if (relicsResult && relicsResult.success) {
              batchUpdate = batchUpdate.concat(relicsResult.batchUpdate || []);
            } else {
              themesAndRelicsSuccess = false;
              themesAndRelicsMessages.push(
                relicsResult ? relicsResult.message : "Unknown error in Relics",
              );
            }
          }
          updateResults.push({
            sheetType: "Themes, Songs & Relics",
            success: themesAndRelicsSuccess,
            message: themesAndRelicsSuccess
              ? "Themes, Songs & Relics updated successfully"
              : "Themes, Songs & Relics update failed: " +
                themesAndRelicsMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Themes, Songs & Relics",
            success: false,
            message: "Error in Themes, Songs & Relics update: " + error.message,
          });
        }
      }

      // Bots updates
      if (data.Bots) {
        try {
          var botsData = data.Bots;
          var botsMasterSheetData = getRangeData("Bots_MS", "values");
          var botsDVTData = buildDVTNamedRangesData(dvtNamedRangesBots);
          var botsSuccess = true;
          var botsMessages = [];
          var botsResult;
          if (
            botsData.hasOwnProperty("oldBots") &&
            botsMasterSheetData &&
            botsDVTData
          ) {
            botsResult = bots.updateBotLevels(
              sheetRequiredRanges.values["Bots_MS"].sheetName,
              botsData.oldBots,
              botsMasterSheetData,
              botsDVTData,
            );
            if (botsResult && botsResult.success) {
              batchUpdate = batchUpdate.concat(botsResult.batchUpdate || []);
            } else {
              botsSuccess = false;
              botsMessages.push(
                botsResult ? botsResult.message : "Unknown error in Bots",
              );
            }
          }
          updateResults.push({
            sheetType: "Bots",
            success: botsSuccess,
            message: botsSuccess
              ? "Bots updated successfully"
              : "Bots update failed: " + botsMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Bots",
            success: false,
            message: "Error in Bots update: " + error.message,
          });
        }
      }

      // Vault updates
      if (data.Vault) {
        try {
          var vaultData = data.Vault;
          var harmonyData = getRangeData("Vault_Harmony", "values");
          var powerData = getRangeData("Vault_Power", "values");
          var vaultSuccess = true;
          var vaultMessages = [];
          var harmonyResult, powerResult;
          if (vaultData.hasOwnProperty("oldVaultHarmony") && harmonyData) {
            harmonyResult = vault.updateVault(
              sheetRequiredRanges.values["Vault_Harmony"].sheetName,
              harmonyData,
              vaultData.oldVaultHarmony,
            );
            if (harmonyResult && harmonyResult.success) {
              batchUpdate = batchUpdate.concat(harmonyResult.batchUpdate || []);
            } else {
              vaultSuccess = false;
              vaultMessages.push(
                "Harmony: " +
                  (harmonyResult ? harmonyResult.message : "Unknown error"),
              );
            }
          }
          if (vaultData.hasOwnProperty("oldVaultPower") && powerData) {
            powerResult = vault.updateVault(
              sheetRequiredRanges.values["Vault_Power"].sheetName,
              powerData,
              vaultData.oldVaultPower,
            );
            if (powerResult && powerResult.success) {
              batchUpdate = batchUpdate.concat(powerResult.batchUpdate || []);
            } else {
              vaultSuccess = false;
              vaultMessages.push(
                "Power: " +
                  (powerResult ? powerResult.message : "Unknown error"),
              );
            }
          }
          updateResults.push({
            sheetType: "Vault",
            success: vaultSuccess,
            message: vaultSuccess
              ? "Vault updated successfully"
              : "Vault update failed: " + vaultMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Vault",
            success: false,
            message: "Error in Vault update: " + error.message,
          });
        }
      }

      // Cards updates
      if (data.Cards) {
        try {
          var cardsData = data.Cards;
          var cardsMasterSheetData = getRangeData("Cards_MS", "values");
          var cardsTrackerData = getRangeData("Cards_Tracker", "values");
          var cardsPresetData = getRangeData("Card Preset", "values");
          var cardsSuccess = true;
          var cardsMessages = [];
          var cardsLevelsResult, cardsPresetResult;
          if (
            cardsData.hasOwnProperty("oldCardsLevel") &&
            cardsData.hasOwnProperty("oldCardSlots") &&
            cardsMasterSheetData
          ) {
            cardsLevelsResult = cards.updateCardsLevels(
              sheetRequiredRanges.values["Cards_MS"].sheetName,
              cardsData.oldCardsLevel,
              cardsData.oldCardSlots,
              cardsMasterSheetData,
            );
            if (cardsLevelsResult && cardsLevelsResult.success) {
              batchUpdate = batchUpdate.concat(
                cardsLevelsResult.batchUpdate || [],
              );
            } else {
              cardsSuccess = false;
              cardsMessages.push(
                "Levels: " +
                  (cardsLevelsResult
                    ? cardsLevelsResult.message
                    : "Unknown error"),
              );
            }
          }
          if (cardsData.hasOwnProperty("oldCardsPreset") && cardsPresetData) {
            var shouldRemoveUsedCards = cardsData.hasOwnProperty("shouldRemoveUsedCards")
              ? cardsData.shouldRemoveUsedCards
              : true;
            cardsPresetResult = cards.updateCardsPreset(
              sheetRequiredRanges.values["Card Preset"].sheetName,
              cardsData.oldCardsPreset,
              shouldRemoveUsedCards,
              cardsPresetData,
            );
            if (cardsPresetResult && cardsPresetResult.success) {
              batchUpdate = batchUpdate.concat(
                cardsPresetResult.batchUpdate || [],
              );
            } else {
              cardsSuccess = false;
              cardsMessages.push(
                "Preset: " +
                  (cardsPresetResult
                    ? cardsPresetResult.message
                    : "Unknown error"),
              );
            }
          }
          if (cardsData.hasOwnProperty("oldCardsTracker") && cardsTrackerData) {
            var cardsTrackerResult = cards.updateCardsTracker(
              sheetRequiredRanges.values["Cards_Tracker"].sheetName,
              cardsData.oldCardsTracker,
              cardsTrackerData,
            );
            if (cardsTrackerResult && cardsTrackerResult.success) {
              batchUpdate = batchUpdate.concat(
                cardsTrackerResult.batchUpdate || [],
              );
            } else {
              cardsSuccess = false;
              cardsMessages.push(
                "Tracker: " +
                  (cardsTrackerResult
                    ? cardsTrackerResult.message
                    : "Unknown error"),
              );
            }
          }
          updateResults.push({
            sheetType: "Cards",
            success: cardsSuccess,
            message: cardsSuccess
              ? "Cards updated successfully"
              : "Cards update failed: " + cardsMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Cards",
            success: false,
            message: "Error in Cards update: " + error.message,
          });
        }
      }

      // Modules updates
      if (data.Modules) {
        try {
          var modulesData = data.Modules;
          var modulesInventoryData = getRangeData(
            "Modules Inventory",
            "values",
          );
          var modulesPresetsData = getRangeData("Modules Presets", "values");
          var modulesTrackerData = getRangeData("Modules Tracker", "values");
          var modulesDVTData = buildDVTNamedRangesData(dvtNamedRangesModules);
          var modulesSuccess = true;
          var modulesMessages = [];
          var inventoryResult, presetsResult;
          if (
            modulesData.hasOwnProperty("oldModulesInventory") &&
            modulesInventoryData
          ) {
            inventoryResult = modules.updateModulesInventory(
              sheetRequiredRanges.values["Modules Inventory"].sheetName,
              modulesData.oldModulesInventory,
              modulesInventoryData,
            );
            if (inventoryResult && inventoryResult.success) {
              batchUpdate = batchUpdate.concat(
                inventoryResult.batchUpdate || [],
              );
            } else {
              modulesSuccess = false;
              modulesMessages.push(
                "Inventory: " +
                  (inventoryResult ? inventoryResult.message : "Unknown error"),
              );
            }
          }
          if (
            modulesData.hasOwnProperty("oldModulesPresets") &&
            modulesPresetsData
          ) {
            presetsResult = modules.updateModulesPresets(
              sheetRequiredRanges.values["Modules Presets"].sheetName,
              modulesData.oldModulesPresets,
              modulesPresetsData,
              modulesDVTData,
            );
            if (presetsResult && presetsResult.success) {
              batchUpdate = batchUpdate.concat(presetsResult.batchUpdate || []);
            } else {
              modulesSuccess = false;
              modulesMessages.push(
                "Presets: " +
                  (presetsResult ? presetsResult.message : "Unknown error"),
              );
            }
          }
          if (
            modulesData.hasOwnProperty("oldModulesTracker") &&
            modulesTrackerData
          ) {
            var trackerResult = modules.updateModulesTracker(
              sheetRequiredRanges.values["Modules Tracker"].sheetName,
              modulesData.oldModulesTracker,
              modulesTrackerData,
            );
            if (trackerResult && trackerResult.success) {
              batchUpdate = batchUpdate.concat(trackerResult.batchUpdate || []);
            } else {
              modulesSuccess = false;
              modulesMessages.push(
                "Tracker: " +
                  (trackerResult ? trackerResult.message : "Unknown error"),
              );
            }
          }
          updateResults.push({
            sheetType: "Modules",
            success: modulesSuccess,
            message: modulesSuccess
              ? "Modules updated successfully"
              : "Modules update failed: " + modulesMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Modules",
            success: false,
            message: "Error in Modules update: " + error.message,
          });
        }
      }

      // Guardians updates
      if (data.Guardians) {
        try {
          var guardiansData = data.Guardians;
          var guardiansMasterSheetData = getRangeData("Guardians_MS", "values");
          var guardiansDVTData = buildDVTNamedRangesData(
            dvtNamedRangesGuardians,
          );
          var guardiansSuccess = true;
          var guardiansMessages = [];
          var guardiansResult;
          if (
            guardiansData.hasOwnProperty("oldGuardians") &&
            guardiansMasterSheetData &&
            guardiansDVTData
          ) {
            guardiansResult = guardians.updateGuardianLevels(
              sheetRequiredRanges.values["Guardians_MS"].sheetName,
              guardiansData.oldGuardians,
              guardiansMasterSheetData,
              guardiansDVTData,
            );
            if (guardiansResult && guardiansResult.success) {
              batchUpdate = batchUpdate.concat(
                guardiansResult.batchUpdate || [],
              );
            } else {
              guardiansSuccess = false;
              guardiansMessages.push(
                guardiansResult
                  ? guardiansResult.message
                  : "Unknown error in Guardians",
              );
            }
          }
          updateResults.push({
            sheetType: "Guardians",
            success: guardiansSuccess,
            message: guardiansSuccess
              ? "Guardians updated successfully"
              : "Guardians update failed: " + guardiansMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Guardians",
            success: false,
            message: "Error in Guardians update: " + error.message,
          });
        }
      }

      // Player updates
      if (data["Player & Stuff"]) {
        try {
          var playerData = data["Player & Stuff"];
          var playerMasterSheetData = getRangeData("player_MS", "values");
          var playerPerkPresetData = getRangeData("Perk Preset", "values");
          var playerSuccess = true;
          var playerMessages = [];
          var playerResult;
          if (
            playerData.hasOwnProperty("oldPlayerStuffTierData") &&
            playerData.hasOwnProperty("oldPlayerStuffStatsData") &&
            playerMasterSheetData
          ) {
            playerResult = playerStuff.updatePlayerStuffData(
              sheetRequiredRanges.values["player_MS"].sheetName,
              playerData.oldPlayerStuffTierData,
              playerData.oldPlayerStuffStatsData,
              playerMasterSheetData,
            );
            if (playerResult && playerResult.success) {
              batchUpdate = batchUpdate.concat(playerResult.batchUpdate || []);
            } else {
              playerSuccess = false;
              playerMessages.push(
                playerResult
                  ? playerResult.message
                  : "Unknown error in Player & Stuff",
              );
            }
          }
          if (
            playerData.hasOwnProperty("oldPerksPreset") &&
            playerPerkPresetData
          ) {
            var shouldRemoveUsedPerks = playerData.hasOwnProperty("shouldRemoveUsedPerks")
              ? playerData.shouldRemoveUsedPerks
              : true;
            var perksResult = playerStuff.updatePlayerPerksPreset(
              sheetRequiredRanges.values["Perk Preset"].sheetName,
              playerData.oldPerksPreset,
              shouldRemoveUsedPerks,
              playerPerkPresetData,
            );
            if (perksResult && perksResult.success) {
              batchUpdate = batchUpdate.concat(perksResult.batchUpdate || []);
            } else {
              playerSuccess = false;
              playerMessages.push(
                perksResult
                  ? perksResult.message
                  : "Unknown error in Player & Stuff Perks Preset",
              );
            }
          }

          updateResults.push({
            sheetType: "Player & Stuff",
            success: playerSuccess,
            message: playerSuccess
              ? "Player & Stuff updated successfully"
              : "Player & Stuff update failed: " + playerMessages.join(", "),
          });
        } catch (error) {
          updateResults.push({
            sheetType: "Player & Stuff",
            success: false,
            message: "Error in Player & Stuff update: " + error.message,
          });
        }
      }

      // Check if any updates failed
      var failedUpdates = updateResults.filter(function (result) {
        return !result.success;
      });

      // Set the newSheetID in "Your ID:" field for performance optimization
      var homePageData = getRangeData("Home Page", "values");
      var yourIdInfo = shared.findSheetTypeID(
        newSheetID,
        "Home Page",
        "Your ID:",
        homePageData,
      );
      if (yourIdInfo && yourIdInfo.cell && yourIdInfo.cell.range) {
        batchUpdate.push({
          range: yourIdInfo.cell.range,
          values: [[newSheetID]],
        });
        console.log(
          `Added IDS Collection sheet ID update to batch: ${newSheetID}`,
        );
      }

      if (failedUpdates.length === 0) {
        var newSheetInfo = shared.findSheetTypeID(
          newSheetID,
          "Home Page",
          "Load your file here",
          homePageData,
        );
        if (
          !newSheetInfo ||
          !newSheetInfo.importStatus ||
          !newSheetInfo.importStatus.range
        ) {
          console.log(`Could not find import status range in IDS sheet`);
          return {
            success: false,
            message: "Could not find import status range in IDS sheet",
          };
        }
        // Add import status update to batch
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });
      }
      // Apply all batch updates at once if we have any
      if (batchUpdate.length > 0) {
        var finalUpdateResult = SheetsAPI.batchUpdateValues(
          newSheetID,
          batchUpdate,
        );
        if (!finalUpdateResult) {
          console.log(
            `Error applying batch updates to IDS Collection spreadsheet`,
          );
          return {
            success: false,
            message:
              "Error applying batch updates to IDS Collection spreadsheet",
          };
        }
      }

      if (failedUpdates.length > 0) {
        var failedSheets = failedUpdates
          .map(function (result) {
            return result.sheetType;
          })
          .join(", ");

        return {
          success: false,
          message: `Failed to update sheets: ${failedSheets}`,
          failedUpdates: failedUpdates,
        };
      }

      return {
        success: true,
        message: "IDS Collection import completed successfully",
      };
    } catch (error) {
      console.log(`Error in IDS Collection importData: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection importData: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Convert Versions
  version4_0: function () {
    try {
      console.log("Called: collection.version4_0");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "Ultimate Weapon": "EXPORT_UW!C5:H", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C4:N", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardians!B4:O", // Guardians data
          "Player Tier": "EXPORT_Player!B3:H", // Player tier data
          "Player Stat": "EXPORT_Player!J3:K", // Player stat data
          "Player Perks": "Perk Preset", // Player perks data (full sheet)
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_2_8WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion3_1_1UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themesAndRelics.getVersion4_0Themes(themesResult.values);
        var relicsData = themesAndRelics.getVersion4_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion3_2Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion3_1Vault(harmonyValues);
        var powerVaultData = vault.getVersion3_1Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion3_1Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion4_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
        var playerPerksData =
          playerStuff.getVersion4_2PlayerStuffPerks(playerPerksValues);
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
      console.log(`Error in IDS Collection version4_0: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version4_0: ${error.message}`,
      };
    }
  },

  version3_2: function () {
    try {
      console.log("Called: collection.version3_2");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "Ultimate Weapon": "EXPORT_UW!C5:H", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C4:L", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardians!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B3:H", // Player tier data
          "Player Stat": "EXPORT_Player!J3:K", // Player stat data
          "Player Perks": "Perk Preset", // Player perks data (full sheet)
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_2_8WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion3_1_1UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
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

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion3_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion3_1Vault(harmonyValues);
        var powerVaultData = vault.getVersion3_1Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion4_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
        var playerPerksData =
          playerStuff.getVersion4_2PlayerStuffPerks(playerPerksValues);
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
      console.log(`Error in IDS Collection version3_2: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version3_2: ${error.message}`,
      };
    }
  },

  version3_0_4: function () {
    try {
      console.log("Called: collection.version3_0_4");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "Ultimate Weapon": "EXPORT_UW!C5:H", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C4:L", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardians!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B3:H", // Player tier data
          "Player Stat": "EXPORT_Player!J3:K", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_2_8WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion3_1_1UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
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

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion3_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion3_1Vault(harmonyValues);
        var powerVaultData = vault.getVersion3_1Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion4_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version3_0_4: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version3_0_4: ${error.message}`,
      };
    }
  },

  version3_0: function () {
    try {
      console.log("Called: collection.version3_0");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C4:L", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardians!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B3:H", // Player tier data
          "Player Stat": "EXPORT_Player!J3:K", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_2_8WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
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

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion3_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion4_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version3_0: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version3_0: ${error.message}`,
      };
    }
  },

  version2_1_4_3: function () {
    try {
      console.log("Called: collection.version2_1_4_3");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardians!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B3:D", // Player tier data
          "Player Stat": "EXPORT_Player!F3:G", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_2_8WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
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

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion2_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion2_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version2_1_4_3: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version2_1_4_3: ${error.message}`,
      };
    }
  },

  version2_1_3_1: function () {
    try {
      console.log("Called: collection.version2_1_3_1");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardian!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B3:D", // Player tier data
          "Player Stat": "EXPORT_Player!F3:G", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_2_8WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
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

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion2_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion2_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version2_1_3_1: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version2_1_3_1: ${error.message}`,
      };
    }
  },

  version2_1_1_8: function () {
    try {
      console.log("Called: collection.version2_1_1_8");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardian!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B3:D", // Player tier data
          "Player Stat": "EXPORT_Player!F3:G", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopPlusRatioResult = getBatchResult(
        "Workshop Ratio",
        "formulas",
      );
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_1WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
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

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion2_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_2Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion2_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version2_1_1_8: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version2_1_1_8: ${error.message}`,
      };
    }
  },

  version2_1_1_6: function () {
    try {
      console.log("Called: collection.version2_1_1_6");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardian!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B3:D", // Player tier data
          "Player Stat": "EXPORT_Player!F3:G", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopPlusRatioResult = getBatchResult(
        "Workshop Ratio",
        "formulas",
      );
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_1WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
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

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion2_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_1Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion2_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version2_1_1_6: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version2_1_1_6: ${error.message}`,
      };
    }
  },

  version2_1: function () {
    try {
      console.log("Called: collection.version2_1");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardian!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B3:D", // Player tier data
          "Player Stat": "EXPORT_Player!F3:G", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopPlusRatioResult = getBatchResult(
        "Workshop Ratio",
        "formulas",
      );
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_1WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themes.getVersion1_0Themes(themesResult.values);
        var relicsData = relics.getVersion1_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion2_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_1Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion2_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion3_2PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version2_1: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version2_1: ${error.message}`,
      };
    }
  },

  version2_0_4: function () {
    try {
      console.log("Called: collection.version2_0_4");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardian!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B16:D", // Player tier data
          "Player Stat": "EXPORT_Player!B2:C12", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopPlusRatioResult = getBatchResult(
        "Workshop Ratio",
        "formulas",
      );
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_1WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themes.getVersion1_0Themes(themesResult.values);
        var relicsData = relics.getVersion1_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion2_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData = cards.getVersion2_0CardsTracker(
          cardsTrackerValues,
          cardsPresetData.oldCardsPreset,
        );

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion2_1Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion2_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion2_0PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version2_0_4: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version2_0_4: ${error.message}`,
      };
    }
  },

  version2_0: function () {
    try {
      console.log("Called: collection.version2_0");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M", // Workshop levels
          "Workshop Plus": "EXPORT_WS!P2:V", // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker data (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardian!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B16:D", // Player tier data
          "Player Stat": "EXPORT_Player!B2:C12", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "Workshop Ratio": "Desired Ratios", // Workshop ratios (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopPlusRatioResult = getBatchResult(
        "Workshop Ratio",
        "formulas",
      );
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
          workshop.getVersion2_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion2_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );
        var workshopPlusRatiosData = workshop.getVersion2_1WorkshopPlusRatios(
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

      // Ultimate Weapon data
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
          ultimate.getVersion2_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themes.getVersion1_0Themes(themesResult.values);
        var relicsData = relics.getVersion1_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion2_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData = cards.getVersion2_0CardsTracker(
          cardsTrackerValues,
          cardsPresetData.oldCardsPreset,
        );

        var cardsSuccess = cardsPresetData.success && cardsLevelData.success;
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

      // Modules data
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
        var modulesInventoryData = modules.getVersion5_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion5_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion1_0Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion2_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion2_0PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version2_0: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version2_0: ${error.message}`,
      };
    }
  },

  version1_4_1_7: function () {
    try {
      console.log("Called: collection.version1_4_1_7");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B3:F", // Workshop levels
          "Workshop Plus": "EXPORT_WS!H3:K", // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker data (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
          Guardians: "EXPORT_Guardian!B5:F", // Guardians data
          "Player Tier": "EXPORT_Player!B16:D", // Player tier data
          "Player Stat": "EXPORT_Player!B2:C12", // Player stat data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
          "Modules Tracker": "Modules Tracker", // Modules Tracker (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};
      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      if (
        workshopLevelsResult &&
        workshopLevelsResult.values &&
        workshopPlusResult &&
        workshopPlusResult.values
      ) {
        var workshopLevelsValues = workshopLevelsResult.values;
        var workshopPlusLevelsValues = workshopPlusResult.values;

        var workshopLevelsData =
          workshop.getVersion1_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion1_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );

        var workshopSuccess =
          workshopLevelsData.success && workshopPlusLevelsData.success;
        collectedData.Workshop = {
          success: workshopSuccess,
          message: workshopSuccess
            ? "Workshop data retrieved successfully"
            : "Error retrieving Workshop data",
          oldWorkshopLevels: workshopLevelsData.oldWorkshopLevels,
          oldWorkshopPlusLevels: workshopPlusLevelsData.oldWorkshopPlusLevels,
        };
      }

      // Ultimate Weapon data
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
          ultimate.getVersion1_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themes.getVersion1_0Themes(themesResult.values);
        var relicsData = relics.getVersion1_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion1_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
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
        var modulesInventoryData = modules.getVersion4_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion4_0ModulesPresets(modulesPresetsValues);
        var modulesTrackerData = modules.getVersion4_7ModulesTracker(
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

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion1_0Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      // Player data
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
          playerStuff.getVersion2_0PlayerStuffTiers(playerTierValues);
        var playerStatData =
          playerStuff.getVersion2_0PlayerStuffStats(playerStatValues);
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
      console.log(`Error in IDS Collection version1_4_1_7: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version1_4_1_7: ${error.message}`,
      };
    }
  },

  version1_3_5: function () {
    try {
      console.log("Called: collection.version1_3_5");
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        values: {
          "Lab Levels": "EXPORT_Lab!B5:E", // Laboratory levels
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B3:F", // Workshop levels
          "Workshop Plus": "EXPORT_WS!H3:K", // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G", // Ultimate weapons data
          "Themes & Songs": "Themes & Songs", // Themes & songs data (full sheet)
          Bots: "EXPORT_Bots!C5:G", // Bots data
          Relics: "Relics", // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony", // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power", // Vault power data (full sheet)
          "Card Preset": "Card Preset", // Cards preset data (full sheet)
          "Card Tracker": "Card and Mastery Tracker", // Card tracker data (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D", // Cards level data
          "Cards Slots": "EXPORT_Cards!C2", // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets", // Modules presets (full sheet)
          Guardians: "EXPORT_Guardian!B5:F", // Guardians data
        },
        formulas: {
          "Lab Planner": "Lab Planner", // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3", // Ultimate Weapons Cost Calculator (full sheet)
        },
      };

      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function (key) {
        return rangeMap.values[key];
      });

      var formulasRanges = Object.keys(rangeMap.formulas).map(function (key) {
        return rangeMap.formulas[key];
      });

      // Batch fetch all required data
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

      // Helper function to get batch result by key and type
      var getBatchResult = function (key, type) {
        type = type || "values"; // Default to values if not specified

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

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};

      // Laboratory data
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

        var labLevelsData = lab.getVersion1_0LabLevels(labLevelsValues);
        var labPlannerData = lab.getVersion1_0LabPlanner(
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

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      if (
        workshopLevelsResult &&
        workshopLevelsResult.values &&
        workshopPlusResult &&
        workshopPlusResult.values
      ) {
        var workshopLevelsValues = workshopLevelsResult.values;
        var workshopPlusLevelsValues = workshopPlusResult.values;

        var workshopLevelsData =
          workshop.getVersion1_0WorkshopLevels(workshopLevelsValues);
        var workshopPlusLevelsData = workshop.getVersion1_0WorkshopPlusLevels(
          workshopPlusLevelsValues,
        );

        var workshopSuccess =
          workshopLevelsData.success && workshopPlusLevelsData.success;
        collectedData.Workshop = {
          success: workshopSuccess,
          message: workshopSuccess
            ? "Workshop data retrieved successfully"
            : "Error retrieving Workshop data",
          oldWorkshopLevels: workshopLevelsData.oldWorkshopLevels,
          oldWorkshopPlusLevels: workshopPlusLevelsData.oldWorkshopPlusLevels,
        };
      }

      // Ultimate Weapon data
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
          ultimate.getVersion1_0UltimateWeapons(ultimateValues);
        var costCalculatorData = ultimate.getVersion1_0CostCalculator(
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

      // Themes, Songs & Relics data - the IDS Collection keeps the two as
      // separate sheets, but exports them under the one merged sheet type.
      var themesResult = getBatchResult("Themes & Songs", "values");
      var relicsResult = getBatchResult("Relics", "values");
      if (
        themesResult &&
        themesResult.values &&
        relicsResult &&
        relicsResult.values
      ) {
        var themesData = themes.getVersion1_0Themes(themesResult.values);
        var relicsData = relics.getVersion1_0Relics(relicsResult.values);
        collectedData["Themes, Songs & Relics"] = {
          success: themesData.success && relicsData.success,
          message: themesData.message || relicsData.message,
          oldThemesNames: themesData.oldThemesNames,
          oldRelics: relicsData.oldRelics,
        };
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion1_0Bots(botsValues);
        collectedData.Bots = botsData;
      }

      // Vault data
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

        var harmonyVaultData = vault.getVersion1_0Vault(harmonyValues);
        var powerVaultData = vault.getVersion1_0Vault(powerValues);

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

      // Cards data
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

        var cardsPresetData = cards.getVersion1_0CardsPreset(cardsPresetValues);
        var cardsLevelData = cards.getVersion1_0CardsLevel(
          cardsLevelValues,
          cardsSlotsValues,
        );
        var cardsTrackerData =
          cards.getVersion1_0CardsTracker(cardsTrackerValues);

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

      // Modules data
      var modulesInventoryResult = getBatchResult(
        "Modules Inventory",
        "values",
      );
      var modulesPresetsResult = getBatchResult("Modules Presets", "values");
      if (
        modulesInventoryResult &&
        modulesInventoryResult.values &&
        modulesPresetsResult &&
        modulesPresetsResult.values
      ) {
        var modulesInventoryValues = modulesInventoryResult.values;
        var modulesPresetsValues = modulesPresetsResult.values;
        var modulesInventoryData = modules.getVersion4_0ModulesInventory(
          modulesInventoryValues,
        );
        var modulesPresetsData =
          modules.getVersion4_0ModulesPresets(modulesPresetsValues);
        var modulesSuccess =
          modulesInventoryData.success && modulesPresetsData.success;
        collectedData.Modules = {
          success: modulesSuccess,
          message: modulesSuccess
            ? "Modules data retrieved successfully"
            : "Error retrieving Modules data",
          oldModulesInventory: modulesInventoryData.oldModulesInventory,
          oldModulesPresets: modulesPresetsData.oldModulesPresets,
        };
      }

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion1_0Guardians(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      return {
        success: true,
        message: "IDS Collection data retrieved successfully",
        data: collectedData,
      };
    } catch (error) {
      console.log(`Error in IDS Collection version1_3_5: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version1_3_5: ${error.message}`,
      };
    }
  },

  // #endregion
  // #region Convert Version Functions Getter
  get convertVersionFunctions() {
    return {
      "v1.3.5": this.version1_3_5.bind(this),
      "v1.4.17": this.version1_4_1_7.bind(this),
      "v2.0": this.version2_0.bind(this),
      "v2.0.4": this.version2_0_4.bind(this),
      "v2.1": this.version2_1.bind(this),
      "v2.1.16": this.version2_1_1_6.bind(this),
      "v2.1.18": this.version2_1_1_8.bind(this),
      "v2.1.31": this.version2_1_3_1.bind(this),
      "v2.1.43": this.version2_1_4_3.bind(this),
      "v3.0": this.version3_0.bind(this),
      "v3.0.4": this.version3_0_4.bind(this),
      "v3.2": this.version3_2.bind(this),
      "v4.0": this.version4_0.bind(this),
    };
  },

  // #endregion
  // #region Compatibility Check
  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = Object.keys(this.convertVersionFunctions);

    var sortedThresholds = versionCompatibility.slice().sort(function (a, b) {
      return shared.compareVersions(b, a) === "newer" ? 1 : -1;
    });

    for (var i = 0; i < sortedThresholds.length; i++) {
      var threshold = sortedThresholds[i];
      var compareResult = shared.compareVersions(oldVersion, threshold);

      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }

    return null;
  },
  // #endregion
};
