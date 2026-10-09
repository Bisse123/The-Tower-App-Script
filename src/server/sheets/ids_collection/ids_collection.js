const collection = {

  /**
   * Reads IDS_Collection data out of the old spreadsheet, using the
   * converter for versionDifference.
   * @param {string} versionDifference
   * @param {string} oldSheetID
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  exportData: function (versionDifference, oldSheetID) {
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

      var oldDataResult = getVersionFunction(oldSheetID);
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
      var errorReport = errors.report("collection.exportData", error, {
        versionDifference: versionDifference,
        oldSheetID: oldSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Writes exported IDS_Collection data into the new spreadsheet.
   * @param {Object} data
   * @param {string} newSheetID
   * @returns {{success: boolean, message: string}} A failure envelope on error.
   */
  importData: function (data, newSheetID) {
    try {
      console.log("Called: collection.importData");

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

      var sheetRequiredRanges = {
        values: {
          "Home Page": { sheetName: "Home Page", range: "Home Page" },
          "Presets Presets": {
            sheetName: "Presets Presets",
            range: "Presets Presets",
          },
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
          Vault_MS: { sheetName: "Vault_MS", range: "Vault_MS" },
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
          "Modules Planner": {
            sheetName: "Modules Planner v2",
            range: "Modules Planner v2",
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

      Object.keys(dvtNamedRangesUW).forEach(function (weapon) {
        Object.keys(dvtNamedRangesUW[weapon]).forEach(function (prop) {
          var rangeName = dvtNamedRangesUW[weapon][prop];
          sheetRequiredRanges.values[rangeName] = {
            sheetName: rangeName,
            range: rangeName,
          };
        });
      });

      Object.keys(dvtNamedRangesBots).forEach(function (bot) {
        Object.keys(dvtNamedRangesBots[bot]).forEach(function (prop) {
          var rangeName = dvtNamedRangesBots[bot][prop];
          sheetRequiredRanges.values[rangeName] = {
            sheetName: rangeName,
            range: rangeName,
          };
        });
      });

      Object.keys(dvtNamedRangesModules).forEach(function (item) {
        var rangeName = dvtNamedRangesModules[item];
        sheetRequiredRanges.values[rangeName] = {
          sheetName: rangeName,
          range: rangeName,
        };
      });

      Object.keys(dvtNamedRangesGuardians).forEach(function (guardian) {
        Object.keys(dvtNamedRangesGuardians[guardian]).forEach(function (prop) {
          var rangeName = dvtNamedRangesGuardians[guardian][prop];
          sheetRequiredRanges.values[rangeName] = {
            sheetName: rangeName,
            range: rangeName,
          };
        });
      });

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

      var updateResults = [];
      var batchUpdate = [];

      var getRangeData = function (sheetName, type) {
        type = type || "values";

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

      var buildDVTNamedRangesData = function (dvtNamedRanges) {
        var dvtNamedRangesData = {};
        Object.keys(dvtNamedRanges).forEach(function (item) {
          var value = dvtNamedRanges[item];
          if (typeof value === "object" && value !== null) {

            dvtNamedRangesData[item] = {};
            Object.keys(value).forEach(function (prop) {
              var rangeData = getRangeData(value[prop], "values");
              dvtNamedRangesData[item][prop] = rangeData || [];
            });
          } else {

            var rangeData = getRangeData(value, "values");
            dvtNamedRangesData[item] = rangeData || [];
          }
        });
        return dvtNamedRangesData;
      };

      if (data["IDS Master"]) {
        try {
          var masterData = data["IDS Master"];
          var masterPresetsData = getRangeData("Presets Presets", "values");
          var masterSuccess = true;
          var masterMessages = [];
          var masterPresetsResult;
          if (masterData.hasOwnProperty("oldPresetsData") && masterPresetsData) {
            masterPresetsResult = masterWriter.updatePresetsData(
              sheetRequiredRanges.values["Presets Presets"].sheetName,
              masterData.oldPresetsData,
              masterPresetsData,
            );
            if (masterPresetsResult && masterPresetsResult.success) {
              batchUpdate = batchUpdate.concat(
                masterPresetsResult.batchUpdate || [],
              );
            } else {
              masterSuccess = false;
              masterMessages.push(
                masterPresetsResult
                  ? masterPresetsResult.message
                  : "Unknown error in Presets",
              );
            }
          }
          updateResults.push({
            sheetType: "IDS Master",
            success: masterSuccess,
            message: masterSuccess
              ? "IDS Master updated successfully"
              : "IDS Master update failed: " + masterMessages.join(", "),
          });
        } catch (error) {
          var errorReport = errors.report("collection.importData", error, {
            note: `Error in IDS Master update`,
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "IDS Master",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

      if (data.Laboratory) {
        try {
          var labData = data.Laboratory;
          var labMasterSheetData = getRangeData("Lab_MS", "values");
          var labPlannerData = getRangeData("Lab Planner", "formulas");
          var labSuccess = true;
          var labMessages = [];
          var labResult, labPlannerResult;
          if (labData.hasOwnProperty("oldLabLevels") && labMasterSheetData) {
            labResult = labWriter.updateLabLevels(
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
            labPlannerResult = labWriter.updateLabPlanner(
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
          var errorReport = errors.report("buildDVTNamedRangesData", error, {
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Laboratory",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

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
            workshopResult = workshopWriter.updateWorkshopLevels(
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
          if (
            workshopData.hasOwnProperty("oldWorkshopPlusRatios") &&
            workshopPlusRatioData
          ) {
            workshopPlusRatioResult = workshopWriter.updateWorkshopPlusRatios(
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
          var errorReport = errors.report("buildDVTNamedRangesData", error, {
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Workshop",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

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
            ultimateResult = ultimateWriter.updateUltimateLevels(
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
              ultimateWriter.updateUltimateCostCalculator(
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
          var errorReport = errors.report("buildDVTNamedRangesData", error, {
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Ultimate Weapon",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

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
            themesResult = themesAndRelicsWriter.updateThemes(
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
            relicsResult = themesAndRelicsWriter.updateRelics(
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
          var errorReport = errors.report("collection.importData", error, {
            note: `Error in Themes, Songs & Relics update`,
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Themes, Songs & Relics",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

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
            botsResult = botsWriter.updateBotLevels(
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
          var errorReport = errors.report("buildDVTNamedRangesData", error, {
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Bots",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

      if (data.Vault) {
        try {
          var vaultData = data.Vault;
          var vaultMasterSheetData = getRangeData("Vault_MS", "values");
          var vaultSuccess = true;
          var vaultMessages = [];
          var vaultResult;

          if (vaultData.hasOwnProperty("oldVault") && vaultMasterSheetData) {
            vaultResult = vaultWriter.updateVault(
              sheetRequiredRanges.values["Vault_MS"].sheetName,
              vaultData.oldVault,
              vaultMasterSheetData,
            );
            if (vaultResult && vaultResult.success) {
              batchUpdate = batchUpdate.concat(vaultResult.batchUpdate || []);
            } else {
              vaultSuccess = false;
              vaultMessages.push(
                vaultResult ? vaultResult.message : "Unknown error in Vault",
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
          var errorReport = errors.report("buildDVTNamedRangesData", error, {
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Vault",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

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
            cardsLevelsResult = cardsWriter.updateCardsLevels(
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
            cardsPresetResult = cardsWriter.updateCardsPreset(
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
            var cardsTrackerResult = cardsWriter.updateCardsTracker(
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
          var errorReport = errors.report("collection.importData", error, {
            note: `Error in Cards update`,
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Cards",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

      if (data.Modules) {
        try {
          var modulesData = data.Modules;
          var modulesInventoryData = getRangeData(
            "Modules Inventory",
            "values",
          );
          var modulesPresetsData = getRangeData("Modules Presets", "values");
          var modulesPlannerData = getRangeData("Modules Planner", "values");
          var modulesTrackerData = getRangeData("Modules Tracker", "values");
          var modulesDVTData = buildDVTNamedRangesData(dvtNamedRangesModules);
          var modulesSuccess = true;
          var modulesMessages = [];
          var inventoryResult, presetsResult;
          if (
            modulesData.hasOwnProperty("oldModulesInventory") &&
            modulesInventoryData
          ) {
            inventoryResult = modulesWriter.updateModulesInventory(
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
            presetsResult = modulesWriter.updateModulesPresets(
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
            modulesData.hasOwnProperty("oldModulesPlanner") &&
            modulesPlannerData
          ) {
            var plannerResult = modulesWriter.updateModulesInventory(
              sheetRequiredRanges.values["Modules Planner"].sheetName,
              modulesData.oldModulesPlanner,
              modulesPlannerData,
            );
            if (plannerResult && plannerResult.success) {
              batchUpdate = batchUpdate.concat(plannerResult.batchUpdate || []);
            } else {
              modulesSuccess = false;
              modulesMessages.push(
                "Planner: " +
                  (plannerResult ? plannerResult.message : "Unknown error"),
              );
            }
          }
          if (
            modulesData.hasOwnProperty("oldModulesTracker") &&
            modulesTrackerData
          ) {
            var trackerResult = modulesWriter.updateModulesTracker(
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
          var errorReport = errors.report("collection.importData", error, {
            note: `Error in Modules update`,
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Modules",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

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
            guardiansResult = guardiansWriter.updateGuardianLevels(
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
          var errorReport = errors.report("collection.importData", error, {
            note: `Error in Guardians update`,
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Guardians",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

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
            playerResult = playerStuffWriter.updatePlayerStuffData(
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
            var perksResult = playerStuffWriter.updatePlayerPerksPreset(
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
          var errorReport = errors.report("collection.importData", error, {
            note: `Error in Player & Stuff update`,
            data: data,
            newSheetID: newSheetID,
          });
          updateResults.push({
            sheetType: "Player & Stuff",
            success: false,
            message: errorReport.message,
            reference: errorReport.reference,
          });
        }
      }

      var failedUpdates = updateResults.filter(function (result) {
        return !result.success;
      });

      var homePageData = getRangeData("Home Page", "values");
      var yourIdInfo = labelUtils.findSheetTypeID(
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
      var errorReport = errors.report("IDS", error, {
        note: `Error in IDS Collection importData`,
        data: data,
        newSheetID: newSheetID,
      });
      return errors.fail(errorReport);
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.3.5": collectionConvertersV1.version1_3_5.bind(collectionConvertersV1),
      "v1.4.17": collectionConvertersV1.version1_4_1_7.bind(collectionConvertersV1),
      "v2.0": collectionConvertersV2_0.version2_0.bind(collectionConvertersV2_0),
      "v2.0.4": collectionConvertersV2_0.version2_0_4.bind(collectionConvertersV2_0),
      "v2.1": collectionConvertersV2_1.version2_1.bind(collectionConvertersV2_1),
      "v2.1.16": collectionConvertersV2_1.version2_1_1_6.bind(collectionConvertersV2_1),
      "v2.1.18": collectionConvertersV2_1.version2_1_1_8.bind(collectionConvertersV2_1),
      "v2.1.31": collectionConvertersV2_1.version2_1_3_1.bind(collectionConvertersV2_1),
      "v2.1.43": collectionConvertersV2_1.version2_1_4_3.bind(collectionConvertersV2_1),
      "v3.0": collectionConvertersV3.version3_0.bind(collectionConvertersV3),
      "v3.0.4": collectionConvertersV3.version3_0_4.bind(collectionConvertersV3),
      "v3.2": collectionConvertersV3.version3_2.bind(collectionConvertersV3),
      "v4.0": collectionConvertersV4.version4_0.bind(collectionConvertersV4),
      "v4.2": collectionConvertersV4.version4_2.bind(collectionConvertersV4),
      "v4.2.4": collectionConvertersV4.version4_2_4.bind(collectionConvertersV4),
    };
  },

  /**
   * The newest converter threshold at or below oldVersion.
   * @param {string} oldVersion
   * @returns {string|null} The threshold, or null when too old.
   */
  isCompatibleVersion: function (oldVersion) {
    var versionCompatibility = Object.keys(this.convertVersionFunctions);

    var sortedThresholds = versionCompatibility.slice().sort(function (a, b) {
      return versionUtils.compareVersions(b, a) === "newer" ? 1 : -1;
    });

    for (var i = 0; i < sortedThresholds.length; i++) {
      var threshold = sortedThresholds[i];
      var compareResult = versionUtils.compareVersions(oldVersion, threshold);

      if (compareResult === "same" || compareResult === "newer") {
        return threshold;
      }
    }

    return null;
  },
};
