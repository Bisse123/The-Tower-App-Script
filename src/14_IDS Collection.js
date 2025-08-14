const collection = {
  exportData: function (versionDifference) {
    try {
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
        data: oldDataResult.data
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting IDS Collection data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      console.log("Starting IDS Collection import data...");
      
      var newSpreadsheet = spreadsheets("IDS Collection newSpreadsheet");
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }
      var newSheetID = newSpreadsheet.spreadsheetId;

      // Configuration dictionary mapping sheet types to their required ranges
      var sheetRequiredRanges = {
        "values": {
          "Home Page": {"sheetName": "Home Page", "range": "Home Page"},
          "Lab_MS": {"sheetName": "Lab_MS", "range": "Lab_MS"},
          "UW_MS": {"sheetName": "UW_MS", "range": "UW_MS"},
          "Themes & Songs": {"sheetName": "Themes & Songs", "range": "Themes & Songs"},
          "Bots_MS": {"sheetName": "Bots_MS", "range": "Bots_MS"},
          "Relics": {"sheetName": "Relics", "range": "Relics"},
          "Vault_Harmony": {"sheetName": "Vault_Harmony", "range": "Vault_Harmony"},
          "Vault_Power": {"sheetName": "Vault_Power", "range": "Vault_Power"},
          "Cards_MS": {"sheetName": "Cards_MS", "range": "Cards_MS"},
          "Modules Inventory": {"sheetName": "Modules Inventory", "range": "Modules Inventory"},
          "Modules Presets": {"sheetName": "Modules Presets", "range": "Modules Presets"},
          "Modules Tracker": {"sheetName": "Modules Tracker", "range": "Modules Tracker"},
          "Guardian_MS": {"sheetName": "Guardian_MS", "range": "Guardian_MS"},
          "player_MS": {"sheetName": "player_MS", "range": "player_MS"},
        },
        "formulas": {
          "Lab Planner": {"sheetName": "Lab Planner", "range": "Lab Planner"},
          "Workshop_MS": {"sheetName": "Workshop_MS", "range": "Workshop_MS"}, 
          "UW Cost Calculator": {"sheetName": "UW Cost Calculator v3", "range": "UW Cost Calculator v3"},
          "Card Preset": {"sheetName": "Card Preset", "range": "Card Preset"}
        }
      };

      // Create ranges arrays from the configuration
      var allValuesRanges = Object.keys(sheetRequiredRanges.values).map(function(key) {
        return sheetRequiredRanges.values[key].range;
      });
      
      var allFormulasRanges = Object.keys(sheetRequiredRanges.formulas).map(function(key) {
        return sheetRequiredRanges.formulas[key].range;
      });

      // Batch fetch all required data
      var batchValuesResults = [];
      var batchFormulasResults = [];
      
      if (allValuesRanges.length > 0) {
        batchValuesResults = SheetsAPI.batchGetValues(newSheetID, allValuesRanges);
        if (!batchValuesResults) {
          console.log(`Error fetching values ranges from IDS Collection spreadsheet`);
          return {
            success: false,
            message: "Error fetching values ranges from IDS Collection spreadsheet"
          };
        }
      }
      
      if (allFormulasRanges.length > 0) {
        batchFormulasResults = SheetsAPI.batchGetFormulas(newSheetID, allFormulasRanges);
        if (!batchFormulasResults) {
          console.log(`Error fetching formulas ranges from IDS Collection spreadsheet`);
          return {
            success: false,
            message: "Error fetching formulas ranges from IDS Collection spreadsheet"
          };
        }
      }

      // Update each sheet type with its data
      var updateResults = [];
      var batchUpdate = [];

      // Helper function to get range data
      var getRangeData = function(sheetName, type) {
        type = type || "values"; // Default to values if not specified
        
        if (type === "formulas") {
          var index = Object.keys(sheetRequiredRanges.formulas).indexOf(sheetName);
          return index !== -1 && batchFormulasResults[index] ? batchFormulasResults[index].values : null;
        } else {
          var index = Object.keys(sheetRequiredRanges.values).indexOf(sheetName);
          return index !== -1 && batchValuesResults[index] ? batchValuesResults[index].values : null;
        }
      };

      // Player updates
      if (data.Player) {
        try {
          var playerData = data.Player;
          var playerMasterSheetData = getRangeData("player_MS", "values");
          
          var playerResult = player.updatePlayer(playerData, sheetRequiredRanges.values["player_MS"].sheetName, playerMasterSheetData);
          if (playerResult && playerResult.success) {
            batchUpdate = batchUpdate.concat(playerResult.batchUpdate || []);
            
            updateResults.push({ 
              sheetType: "Player", 
              success: true, 
              message: "Player updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Player", 
              success: false, 
              message: playerResult ? playerResult.message : "Unknown error in Player update" 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Player", 
            success: false, 
            message: "Error in Player update: " + error.message 
          });
        }
      }
      // Laboratory updates
      if (data.Laboratory) {
        try {
          var labData = data.Laboratory;
          var labMasterSheetData = getRangeData("Lab_MS", "values");
          var labPlannerData = getRangeData("Lab Planner", "formulas");
          
          var labResult = lab.updateLabLevels(sheetRequiredRanges.values["Lab_MS"].sheetName, labData.oldLabLevels, labMasterSheetData);
          if (labResult && labResult.success) {
            batchUpdate = batchUpdate.concat(labResult.batchUpdate || []);
            
            if (labData.oldLabPlanner && labPlannerData) {
              var labPlannerResult = lab.updateLabPlanner(sheetRequiredRanges.formulas["Lab Planner"].sheetName, labData.oldLabPlanner, labPlannerData);
              if (labPlannerResult && labPlannerResult.success) {
                batchUpdate = batchUpdate.concat(labPlannerResult.batchUpdate || []);
              }
            }
            
            updateResults.push({ 
              sheetType: "Laboratory", 
              success: true, 
              message: "Laboratory updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Laboratory", 
              success: false, 
              message: labResult ? labResult.message : "Unknown error in Laboratory update" 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Laboratory", 
            success: false, 
            message: "Error in Laboratory update: " + error.message 
          });
        }
      }

      // Workshop updates
      if (data.Workshop) {
        try {
          var workshopData = data.Workshop;
          var workshopMasterSheetData = getRangeData("Workshop_MS", "formulas");
          
          var workshopResult = workshop.updateWorkshopLevels(sheetRequiredRanges.formulas["Workshop_MS"].sheetName, workshopData.oldWorkshopLevels, workshopData.oldWorkshopPlusLevels, workshopMasterSheetData);
          if (workshopResult && workshopResult.success) {
            batchUpdate = batchUpdate.concat(workshopResult.batchUpdate || []);
            
            updateResults.push({ 
              sheetType: "Workshop", 
              success: true, 
              message: "Workshop updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Workshop", 
              success: false, 
              message: workshopResult ? workshopResult.message : "Unknown error in Workshop update" 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Workshop", 
            success: false, 
            message: "Error in Workshop update: " + error.message 
          });
        }
      }

      // Ultimate Weapon updates
      if (data["Ultimate Weapon"]) {
        try {
          var ultimateData = data["Ultimate Weapon"];
          var ultimateMasterSheetData = getRangeData("UW_MS", "values");
          var ultimateCostCalculatorData = getRangeData("UW Cost Calculator", "formulas");
          
          var ultimateResult = ultimate.updateUltimateLevels(ultimateData.targetWeapons, sheetRequiredRanges.values["UW_MS"].sheetName, ultimateData.oldUltimate, ultimateMasterSheetData);
          if (ultimateResult && ultimateResult.success) {
            batchUpdate = batchUpdate.concat(ultimateResult.batchUpdate || []);
            
            // Update cost calculator if data exists
            if (ultimateData.oldUltimateCostCalculator && ultimateCostCalculatorData) {
              var ultimateCostCalculatorResult = ultimate.updateUltimateCostCalculator(
                ultimateData.targetWeapons, 
                sheetRequiredRanges.formulas["UW Cost Calculator"].sheetName, 
                ultimateData.oldUltimateCostCalculator, 
                ultimateCostCalculatorData
              );
              if (ultimateCostCalculatorResult && ultimateCostCalculatorResult.success) {
                batchUpdate = batchUpdate.concat(ultimateCostCalculatorResult.batchUpdate || []);
              }
            }
            
            updateResults.push({ 
              sheetType: "Ultimate Weapon", 
              success: true, 
              message: "Ultimate Weapon updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Ultimate Weapon", 
              success: false, 
              message: ultimateResult ? ultimateResult.message : "Unknown error in Ultimate Weapon update" 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Ultimate Weapon", 
            success: false, 
            message: "Error in Ultimate Weapon update: " + error.message 
          });
        }
      }

      // Themes & Songs updates
      if (data["Themes & Songs"]) {
        try {
          var themesData = data["Themes & Songs"];
          var themesMasterSheetData = getRangeData("Themes & Songs", "values");
          
          var themesResult = themes.updateThemes(themesData.targetThemes, "Themes & Songs", themesData.oldThemesNames, themesMasterSheetData);
          if (themesResult && themesResult.success) {
            batchUpdate = batchUpdate.concat(themesResult.batchUpdate || []);
            
            updateResults.push({ 
              sheetType: "Themes & Songs", 
              success: true, 
              message: "Themes & Songs updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Themes & Songs", 
              success: false, 
              message: themesResult ? themesResult.message : "Unknown error in Themes & Songs update" 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Themes & Songs", 
            success: false, 
            message: "Error in Themes & Songs update: " + error.message 
          });
        }
      }

      // Bots updates
      if (data.Bots) {
        try {
          var botsData = data.Bots;
          var botsMasterSheetData = getRangeData("Bots_MS", "values");
          
          var botsResult = bots.updateBotLevels(sheetRequiredRanges.values["Bots_MS"].sheetName, botsData.targetBots, botsData.oldBots, botsMasterSheetData);
          if (botsResult && botsResult.success) {
            batchUpdate = batchUpdate.concat(botsResult.batchUpdate || []);
            
            updateResults.push({ 
              sheetType: "Bots", 
              success: true, 
              message: "Bots updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Bots", 
              success: false, 
              message: botsResult ? botsResult.message : "Unknown error in Bots update" 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Bots", 
            success: false, 
            message: "Error in Bots update: " + error.message 
          });
        }
      }

      // Relics updates
      if (data.Relics) {
        try {
          var relicsData = data.Relics;
          var relicsMasterSheetData = getRangeData("Relics", "values");
          
          var relicsResult = relics.updateRelics(sheetRequiredRanges.values["Relics"].sheetName, relicsData.oldRelics, relicsMasterSheetData);
          if (relicsResult && relicsResult.success) {
            batchUpdate = batchUpdate.concat(relicsResult.batchUpdate || []);
            
            updateResults.push({ 
              sheetType: "Relics", 
              success: true, 
              message: "Relics updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Relics", 
              success: false, 
              message: relicsResult ? relicsResult.message : "Unknown error in Relics update" 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Relics", 
            success: false, 
            message: "Error in Relics update: " + error.message 
          });
        }
      }

      // Vault updates
      if (data.Vault) {
        try {
          var vaultData = data.Vault;
          var harmonyData = getRangeData("Vault_Harmony", "values");
          var powerData = getRangeData("Vault_Power", "values");
          
          var harmonyResult = vault.updateVault(sheetRequiredRanges.values["Vault_Harmony"].sheetName, harmonyData, vaultData.oldVaultHarmony);
          var powerResult = vault.updateVault(sheetRequiredRanges.values["Vault_Power"].sheetName, powerData, vaultData.oldVaultPower);
          
          var vaultSuccess = true;
          var vaultMessages = [];
          
          if (harmonyResult && harmonyResult.success) {
            batchUpdate = batchUpdate.concat(harmonyResult.batchUpdate || []);
          } else {
            vaultSuccess = false;
            vaultMessages.push("Harmony: " + (harmonyResult ? harmonyResult.message : "Unknown error"));
          }
          
          if (powerResult && powerResult.success) {
            batchUpdate = batchUpdate.concat(powerResult.batchUpdate || []);
          } else {
            vaultSuccess = false;
            vaultMessages.push("Power: " + (powerResult ? powerResult.message : "Unknown error"));
          }
          
          if (vaultSuccess) {
            updateResults.push({ 
              sheetType: "Vault", 
              success: true, 
              message: "Vault updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Vault", 
              success: false, 
              message: "Vault update failed: " + vaultMessages.join(", ") 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Vault", 
            success: false, 
            message: "Error in Vault update: " + error.message 
          });
        }
      }

      // Cards updates
      if (data.Cards) {
        try {
          var cardsData = data.Cards;
          var cardsMasterSheetData = getRangeData("Cards_MS", "values");
          var cardsPresetData = getRangeData("Card Preset", "formulas");
          
          var cardsLevelsResult = cards.updateCardsLevels(sheetRequiredRanges.values["Cards_MS"].sheetName, cardsData.oldCardsLevel, cardsData.oldCardSlots, cardsMasterSheetData);
          var cardsPresetResult = cards.updateCardsPreset("Card Preset", cardsData.oldCardsPreset, cardsData.shouldRemoveUsedCards, cardsPresetData);
          
          var cardsSuccess = true;
          var cardsMessages = [];
          
          if (cardsLevelsResult && cardsLevelsResult.success) {
            batchUpdate = batchUpdate.concat(cardsLevelsResult.batchUpdate || []);
          } else {
            cardsSuccess = false;
            cardsMessages.push("Levels: " + (cardsLevelsResult ? cardsLevelsResult.message : "Unknown error"));
          }
          
          if (cardsPresetResult && cardsPresetResult.success) {
            batchUpdate = batchUpdate.concat(cardsPresetResult.batchUpdate || []);
          } else {
            cardsSuccess = false;
            cardsMessages.push("Preset: " + (cardsPresetResult ? cardsPresetResult.message : "Unknown error"));
          }
          
          if (cardsSuccess) {
            updateResults.push({ 
              sheetType: "Cards", 
              success: true, 
              message: "Cards updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Cards", 
              success: false, 
              message: "Cards update failed: " + cardsMessages.join(", ") 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Cards", 
            success: false, 
            message: "Error in Cards update: " + error.message 
          });
        }
      }

      // Modules updates
      if (data.Modules) {
        try {
          var modulesData = data.Modules;
          var modulesInventoryData = getRangeData("Modules Inventory", "values");
          var modulesPresetsData = getRangeData("Modules Presets", "values");
          var modulesObtainedData = getRangeData("Modules Tracker", "values");
          
          var inventoryResult = modules.updateModulesInventory(sheetRequiredRanges.values["Modules Inventory"].sheetName, modulesData.targetModuleTypes, modulesData.oldModulesInventory, modulesInventoryData);
          var presetsResult = modules.updateModulesPresets(sheetRequiredRanges.values["Modules Presets"].sheetName, modulesData.targetModuleTypes, modulesData.oldModulesPresets, modulesPresetsData);
          var obtainedResult = modules.updateModulesObtained(sheetRequiredRanges.values["Modules Tracker"].sheetName, modulesData.targetModuleTypes, modulesData.oldModulesObtained, modulesObtainedData);
          
          var modulesSuccess = true;
          var modulesMessages = [];
          
          if (inventoryResult && inventoryResult.success) {
            batchUpdate = batchUpdate.concat(inventoryResult.batchUpdate || []);
          } else {
            modulesSuccess = false;
            modulesMessages.push("Inventory: " + (inventoryResult ? inventoryResult.message : "Unknown error"));
          }
          
          if (presetsResult && presetsResult.success) {
            batchUpdate = batchUpdate.concat(presetsResult.batchUpdate || []);
          } else {
            modulesSuccess = false;
            modulesMessages.push("Presets: " + (presetsResult ? presetsResult.message : "Unknown error"));
          }
          
          if (obtainedResult && obtainedResult.success) {
            batchUpdate = batchUpdate.concat(obtainedResult.batchUpdate || []);
          } else {
            modulesSuccess = false;
            modulesMessages.push("Obtained: " + (obtainedResult ? obtainedResult.message : "Unknown error"));
          }
          
          if (modulesSuccess) {
            updateResults.push({ 
              sheetType: "Modules", 
              success: true, 
              message: "Modules updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Modules", 
              success: false, 
              message: "Modules update failed: " + modulesMessages.join(", ") 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Modules", 
            success: false, 
            message: "Error in Modules update: " + error.message 
          });
        }
      }

      // Guardians updates
      if (data.Guardians) {
        try {
          var guardiansData = data.Guardians;
          var guardiansMasterSheetData = getRangeData("Guardian_MS", "values");
          
          var guardiansResult = guardians.updateGuardianLevels(guardiansData.targetGuardians, sheetRequiredRanges.values["Guardian_MS"].sheetName, guardiansData.oldGuardians, guardiansMasterSheetData);
          if (guardiansResult && guardiansResult.success) {
            batchUpdate = batchUpdate.concat(guardiansResult.batchUpdate || []);
            
            updateResults.push({ 
              sheetType: "Guardians", 
              success: true, 
              message: "Guardians updated successfully" 
            });
          } else {
            updateResults.push({ 
              sheetType: "Guardians", 
              success: false, 
              message: guardiansResult ? guardiansResult.message : "Unknown error in Guardians update" 
            });
          }
        } catch (error) {
          updateResults.push({ 
            sheetType: "Guardians", 
            success: false, 
            message: "Error in Guardians update: " + error.message 
          });
        }
      }

      // Check if any updates failed
      var failedUpdates = updateResults.filter(function(result) {
        return !result.success;
      });

      if (failedUpdates.length === 0) {
        var homePageData = getRangeData("Home Page", "values");
        var newSheetInfo = shared.findSheetTypeID(newSheetID, "Home Page", "Load your file here", homePageData);
        if (!newSheetInfo || !newSheetInfo.importStatus || !newSheetInfo.importStatus.range) {
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
        var finalUpdateResult = SheetsAPI.batchUpdateValues(newSheetID, batchUpdate);
        if (!finalUpdateResult) {
          console.log(`Error applying batch updates to IDS Collection spreadsheet`);
          return {
            success: false,
            message: "Error applying batch updates to IDS Collection spreadsheet",
          };
        }
      }

      if (failedUpdates.length > 0) {
        var failedSheets = failedUpdates.map(function(result) {
          return result.sheetType;
        }).join(", ");
        
        return {
          success: false,
          message: `Failed to update sheets: ${failedSheets}`,
          failedUpdates: failedUpdates
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

  version20: function () {
    try {
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      
      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        "values": {
          "Lab Levels": "EXPORT_Lab!B5:E",              // Laboratory levels
          "Lab Planner": "Lab Planner",                 // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B2:M",          // Workshop levels  
          "Workshop Plus": "EXPORT_WS!P2:V",            // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G",          // Ultimate weapons data
          "Themes & Songs": "Themes & Songs",           // Themes & songs data (full sheet)
          "Bots": "EXPORT_Bots!C5:G",                   // Bots data
          "Relics": "Relics",                           // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony",             // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power",                 // Vault power data (full sheet)
          "Card Preset": "Card Preset",                 // Cards preset data (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D",          // Cards level data
          "Cards Slots": "EXPORT_Cards!C2",             // Cards slot data
          "Modules Inventory": "Modules Inventory",     // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets",         // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker",         // Modules obtained (full sheet)
          "Guardians": "EXPORT_Guardian!B5:F",          // Guardians data
          "Player": "EXPORT_Player!B2:D",               // Player data
        },
        "formulas": {
          "Lab Planner": "Lab Planner",                 // Laboratory planner (full sheet)
          "Workshop Ratio": "Desired Ratios",           // Workshop ratios (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3" // Ultimate Weapons Cost Calculator (full sheet)
        }
      };
      
      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function(key) {
        return rangeMap.values[key];
      });
      
      var formulasRanges = Object.keys(rangeMap.formulas).map(function(key) {
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
        batchFormulasResults = SheetsAPI.batchGetFormulas(oldSheetID, formulasRanges);
        if (!batchFormulasResults) {
          console.log(`Could not read IDS Collection formulas data`);
          return {
            success: false,
            message: "Could not read IDS Collection formulas data",
          };
        }
      }
      
      // Helper function to get batch result by key and type
      var getBatchResult = function(key, type) {
        type = type || "values"; // Default to values if not specified
        
        if (type === "values") {
          var index = Object.keys(rangeMap.values).indexOf(key);
          return index !== -1 && batchValuesResults[index] ? batchValuesResults[index] : null;
        } else if (type === "formulas") {
          var index = Object.keys(rangeMap.formulas).indexOf(key);
          return index !== -1 && batchFormulasResults[index] ? batchFormulasResults[index] : null;
        }
        return null;
      };

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};
      // Player data
      var playerResult = getBatchResult("Player", "values");
      if (playerResult && playerResult.values) {
        var playerValues = playerResult.values;
        var playerData = player.getVersion20Values(playerValues);
        collectedData.Player = playerData;
      }
      // Laboratory data
      var labLevelsResult = getBatchResult("Lab Levels", "values");
      var labPlannerValuesResult = getBatchResult("Lab Planner", "values");
      var labPlannerFormulasResult = getBatchResult("Lab Planner", "formulas");
      if (labLevelsResult && labLevelsResult.values) {
        var labLevelsValues = labLevelsResult.values;
        var labPlannerValues = labPlannerValuesResult && labPlannerValuesResult.values ? labPlannerValuesResult.values : null;
        var labPlannerFormulas = labPlannerFormulasResult && labPlannerFormulasResult.values ? labPlannerFormulasResult.values : null;
        var labData = lab.getVersion10Values(labLevelsValues, labPlannerValues, labPlannerFormulas);
        collectedData.Laboratory = labData;
      }

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      var workshopRatioResult = getBatchResult("Workshop Ratio", "formulas");
      if (workshopLevelsResult && workshopLevelsResult.values && workshopPlusResult && workshopPlusResult.values && workshopRatioResult && workshopRatioResult.values) {
        var workshopLevelsValues = workshopLevelsResult.values;
        var workshopPlusLevelsValues = workshopPlusResult.values;
        var workshopRatioValues = workshopRatioResult.values;
        var workshopData = workshop.getVersion20Values(workshopLevelsValues, workshopPlusLevelsValues, workshopRatioValues);
        collectedData.Workshop = workshopData;
      }

      // Ultimate Weapon data
      var ultimateResult = getBatchResult("Ultimate Weapon", "values");
      var ultimateCostCalculatorResult = getBatchResult("UW Cost Calculator", "formulas");
      if (ultimateResult && ultimateResult.values && ultimateCostCalculatorResult && ultimateCostCalculatorResult.values) {
        var ultimateValues = ultimateResult.values;
        var ultimateCostCalculatorValues = ultimateCostCalculatorResult.values;
        var ultimateData = ultimate.getVersion20Values(ultimateValues, ultimateCostCalculatorValues);
        collectedData["Ultimate Weapon"] = ultimateData;
      }

      // Themes & Songs data
      var themesResult = getBatchResult("Themes & Songs", "values");
      if (themesResult && themesResult.values) {
        var themesValues = themesResult.values;
        var themesData = themes.getVersion10Values(themesValues);
        collectedData["Themes & Songs"] = themesData;
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion20Values(botsValues);
        collectedData.Bots = botsData;
      }

      // Relics data
      var relicsResult = getBatchResult("Relics", "values");
      if (relicsResult && relicsResult.values) {
        var relicsValues = relicsResult.values;
        var relicsData = relics.getVersion10Values(relicsValues);
        collectedData.Relics = relicsData;
      }

      // Vault data
      var harmonyResult = getBatchResult("Vault Harmony", "values");
      var powerResult = getBatchResult("Vault Power", "values");
      if (harmonyResult && harmonyResult.values && powerResult && powerResult.values) {
        var harmonyValues = harmonyResult.values;
        var powerValues = powerResult.values;
        var vaultData = vault.getVersion10Values(harmonyValues, powerValues);
        collectedData.Vault = vaultData;
      }

      // Cards data
      var cardsPresetResult = getBatchResult("Card Preset", "values");
      var cardsLevelsResult = getBatchResult("Cards Levels", "values");
      var cardsSlotsResult = getBatchResult("Cards Slots", "values");
      if (cardsPresetResult && cardsPresetResult.values && cardsLevelsResult && cardsLevelsResult.values && cardsSlotsResult && cardsSlotsResult.values) {
        var cardsPresetValues = cardsPresetResult.values;
        var cardsLevelValues = cardsLevelsResult.values;
        var cardsSlotsValues = cardsSlotsResult.values;
        var cardsData = cards.getVersion10Values(cardsPresetValues, cardsLevelValues, cardsSlotsValues);
        collectedData.Cards = cardsData;
      }

      // Modules data
      var modulesInventoryResult = getBatchResult("Modules Inventory", "values");
      var modulesPresetsResult = getBatchResult("Modules Presets", "values");
      var modulesObtainedResult = getBatchResult("Modules Tracker", "values");
      if (modulesInventoryResult && modulesInventoryResult.values && modulesPresetsResult && modulesPresetsResult.values && modulesObtainedResult && modulesObtainedResult.values) {
        var modulesInventoryValues = modulesInventoryResult.values;
        var modulesPresetsValues = modulesPresetsResult.values;
        var modulesObtainedValues = modulesObtainedResult.values;
        var modulesData = modules.getVersion50Values(modulesInventoryValues, modulesPresetsValues, modulesObtainedValues);
        collectedData.Modules = modulesData;
      }

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion10Values(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      return {
        success: true,
        message: "IDS Collection data retrieved successfully",
        data: collectedData
      };
      
    } catch (error) {
      console.log(`Error in IDS Collection version1417: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version1417: ${error.message}`,
      };
    }
  },

  version1417: function () {
    try {
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      
      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        "values": {
          "Lab Levels": "EXPORT_Lab!B5:E",          // Laboratory levels
          "Lab Planner": "Lab Planner",             // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B3:F",      // Workshop levels  
          "Workshop Plus": "EXPORT_WS!H3:K",        // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G",      // Ultimate weapons data
          "Themes & Songs": "Themes & Songs",       // Themes & songs data (full sheet)
          "Bots": "EXPORT_Bots!C5:G",               // Bots data
          "Relics": "Relics",                       // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony",         // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power",             // Vault power data (full sheet)
          "Card Preset": "Card Preset",             // Cards preset data (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D",      // Cards level data
          "Cards Slots": "EXPORT_Cards!C2",         // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets",     // Modules presets (full sheet)
          "Modules Tracker": "Modules Tracker",         // Modules obtained (full sheet)
          "Guardians": "EXPORT_Guardian!B5:F",       // Guardians data
          "Player": "EXPORT_Player!B2:D",           // Player data
        },
        "formulas": {
          "Lab Planner": "Lab Planner",              // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3" // Ultimate Weapons Cost Calculator (full sheet)
        }
      };
      
      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function(key) {
        return rangeMap.values[key];
      });
      
      var formulasRanges = Object.keys(rangeMap.formulas).map(function(key) {
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
        batchFormulasResults = SheetsAPI.batchGetFormulas(oldSheetID, formulasRanges);
        if (!batchFormulasResults) {
          console.log(`Could not read IDS Collection formulas data`);
          return {
            success: false,
            message: "Could not read IDS Collection formulas data",
          };
        }
      }
      
      // Helper function to get batch result by key and type
      var getBatchResult = function(key, type) {
        type = type || "values"; // Default to values if not specified
        
        if (type === "values") {
          var index = Object.keys(rangeMap.values).indexOf(key);
          return index !== -1 && batchValuesResults[index] ? batchValuesResults[index] : null;
        } else if (type === "formulas") {
          var index = Object.keys(rangeMap.formulas).indexOf(key);
          return index !== -1 && batchFormulasResults[index] ? batchFormulasResults[index] : null;
        }
        return null;
      };

      // Process the data using the individual modules' getVersionXXValues functions
      var collectedData = {};
      // Player data
      var playerResult = getBatchResult("Player", "values");
      if (playerResult && playerResult.values) {
        var playerValues = playerResult.values;
        var playerData = player.getVersion20Values(playerValues);
        collectedData.Player = playerData;
      }
      // Laboratory data
      var labLevelsResult = getBatchResult("Lab Levels", "values");
      var labPlannerValuesResult = getBatchResult("Lab Planner", "values");
      var labPlannerFormulasResult = getBatchResult("Lab Planner", "formulas");
      if (labLevelsResult && labLevelsResult.values) {
        var labLevelsValues = labLevelsResult.values;
        var labPlannerValues = labPlannerValuesResult && labPlannerValuesResult.values ? labPlannerValuesResult.values : null;
        var labPlannerFormulas = labPlannerFormulasResult && labPlannerFormulasResult.values ? labPlannerFormulasResult.values : null;
        var labData = lab.getVersion10Values(labLevelsValues, labPlannerValues, labPlannerFormulas);
        collectedData.Laboratory = labData;
      }

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      if (workshopLevelsResult && workshopLevelsResult.values && workshopPlusResult && workshopPlusResult.values) {
        var workshopLevelsValues = workshopLevelsResult.values;
        var workshopPlusLevelsValues = workshopPlusResult.values;
        var workshopData = workshop.getVersion10Values(workshopLevelsValues, workshopPlusLevelsValues);
        collectedData.Workshop = workshopData;
      }

      // Ultimate Weapon data
      var ultimateResult = getBatchResult("Ultimate Weapon", "values");
      var ultimateCostCalculatorResult = getBatchResult("UW Cost Calculator", "formulas");
      if (ultimateResult && ultimateResult.values && ultimateCostCalculatorResult && ultimateCostCalculatorResult.values) {
        var ultimateValues = ultimateResult.values;
        var ultimateCostCalculatorValues = ultimateCostCalculatorResult.values;
        var ultimateData = ultimate.getVersion20Values(ultimateValues, ultimateCostCalculatorValues);
        collectedData["Ultimate Weapon"] = ultimateData;
      }

      // Themes & Songs data
      var themesResult = getBatchResult("Themes & Songs", "values");
      if (themesResult && themesResult.values) {
        var themesValues = themesResult.values;
        var themesData = themes.getVersion10Values(themesValues);
        collectedData["Themes & Songs"] = themesData;
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion20Values(botsValues);
        collectedData.Bots = botsData;
      }

      // Relics data
      var relicsResult = getBatchResult("Relics", "values");
      if (relicsResult && relicsResult.values) {
        var relicsValues = relicsResult.values;
        var relicsData = relics.getVersion10Values(relicsValues);
        collectedData.Relics = relicsData;
      }

      // Vault data
      var harmonyResult = getBatchResult("Vault Harmony", "values");
      var powerResult = getBatchResult("Vault Power", "values");
      if (harmonyResult && harmonyResult.values && powerResult && powerResult.values) {
        var harmonyValues = harmonyResult.values;
        var powerValues = powerResult.values;
        var vaultData = vault.getVersion10Values(harmonyValues, powerValues);
        collectedData.Vault = vaultData;
      }

      // Cards data
      var cardsPresetResult = getBatchResult("Card Preset", "values");
      var cardsLevelsResult = getBatchResult("Cards Levels", "values");
      var cardsSlotsResult = getBatchResult("Cards Slots", "values");
      if (cardsPresetResult && cardsPresetResult.values && cardsLevelsResult && cardsLevelsResult.values && cardsSlotsResult && cardsSlotsResult.values) {
        var cardsPresetValues = cardsPresetResult.values;
        var cardsLevelValues = cardsLevelsResult.values;
        var cardsSlotsValues = cardsSlotsResult.values;
        var cardsData = cards.getVersion10Values(cardsPresetValues, cardsLevelValues, cardsSlotsValues);
        collectedData.Cards = cardsData;
      }

      // Modules data
      var modulesInventoryResult = getBatchResult("Modules Inventory", "values");
      var modulesPresetsResult = getBatchResult("Modules Presets", "values");
      var modulesObtainedResult = getBatchResult("Modules Tracker", "values");
      if (modulesInventoryResult && modulesInventoryResult.values && modulesPresetsResult && modulesPresetsResult.values && modulesObtainedResult && modulesObtainedResult.values) {
        var modulesInventoryValues = modulesInventoryResult.values;
        var modulesPresetsValues = modulesPresetsResult.values;
        var modulesObtainedValues = modulesObtainedResult.values;
        var modulesData = modules.getVersion50Values(modulesInventoryValues, modulesPresetsValues, modulesObtainedValues);
        collectedData.Modules = modulesData;
      }

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion10Values(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      return {
        success: true,
        message: "IDS Collection data retrieved successfully",
        data: collectedData
      };
      
    } catch (error) {
      console.log(`Error in IDS Collection version1417: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version1417: ${error.message}`,
      };
    }
  },
  
  version135: function () {
    try {
      var oldSpreadsheet = spreadsheets("IDS Collection oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;
      
      // Define all the ranges for different sheet types in the IDS Collection
      // Dictionary mapping descriptive keys to their actual sheet ranges, separated by type
      var rangeMap = {
        "values": {
          "Lab Levels": "EXPORT_Lab!B5:E",           // Laboratory levels
          "Lab Planner": "Lab Planner",              // Laboratory planner (full sheet) - for values
          "Workshop Levels": "EXPORT_WS!B3:F",      // Workshop levels  
          "Workshop Plus": "EXPORT_WS!H3:K",        // Workshop plus levels
          "Ultimate Weapon": "EXPORT_UW!C5:G",      // Ultimate weapons data
          "Themes & Songs": "Themes & Songs",       // Themes & songs data (full sheet)
          "Bots": "EXPORT_Bots!C5:G",               // Bots data
          "Relics": "Relics",                       // Relics data (full sheet)
          "Vault Harmony": "Vault_Harmony",         // Vault harmony data (full sheet)
          "Vault Power": "Vault_Power",             // Vault power data (full sheet)
          "Card Preset": "Card Preset",              // Cards preset data (full sheet)
          "Cards Levels": "EXPORT_Cards!B5:D",      // Cards level data
          "Cards Slots": "EXPORT_Cards!C2",         // Cards slot data
          "Modules Inventory": "Modules Inventory", // Modules inventory (full sheet)
          "Modules Presets": "Modules Presets",     // Modules presets (full sheet)
          "Mods Obtained": "Mods Obtained",         // Modules obtained (full sheet)
          "Guardians": "EXPORT_Guardian!B5:F"      // Guardians data
        },
        "formulas": {
          "Lab Planner": "Lab Planner",              // Laboratory planner (full sheet)
          "UW Cost Calculator": "UW Cost Calculator v3" // Ultimate Weapons Cost Calculator (full sheet)
        }
      };
      
      // Create separate ranges arrays and index maps for values and formulas
      var valuesRanges = Object.keys(rangeMap.values).map(function(key) {
        return rangeMap.values[key];
      });
      
      var formulasRanges = Object.keys(rangeMap.formulas).map(function(key) {
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
        batchFormulasResults = SheetsAPI.batchGetFormulas(oldSheetID, formulasRanges);
        if (!batchFormulasResults) {
          console.log(`Could not read IDS Collection formulas data`);
          return {
            success: false,
            message: "Could not read IDS Collection formulas data",
          };
        }
      }
      
      // Helper function to get batch result by key and type
      var getBatchResult = function(key, type) {
        type = type || "values"; // Default to values if not specified
        
        if (type === "values") {
          var index = Object.keys(rangeMap.values).indexOf(key);
          return index !== -1 && batchValuesResults[index] ? batchValuesResults[index] : null;
        } else if (type === "formulas") {
          var index = Object.keys(rangeMap.formulas).indexOf(key);
          return index !== -1 && batchFormulasResults[index] ? batchFormulasResults[index] : null;
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
        var labPlannerValues = labPlannerValuesResult && labPlannerValuesResult.values ? labPlannerValuesResult.values : null;
        var labPlannerFormulas = labPlannerFormulasResult && labPlannerFormulasResult.values ? labPlannerFormulasResult.values : null;
        var labData = lab.getVersion10Values(labLevelsValues, labPlannerValues, labPlannerFormulas);
        collectedData.Laboratory = labData;
      }

      // Workshop data
      var workshopLevelsResult = getBatchResult("Workshop Levels", "values");
      var workshopPlusResult = getBatchResult("Workshop Plus", "values");
      if (workshopLevelsResult && workshopLevelsResult.values && workshopPlusResult && workshopPlusResult.values) {
        var workshopLevelsValues = workshopLevelsResult.values;
        var workshopPlusLevelsValues = workshopPlusResult.values;
        var workshopData = workshop.getVersion10Values(workshopLevelsValues, workshopPlusLevelsValues);
        collectedData.Workshop = workshopData;
      }

      // Ultimate Weapon data
      var ultimateResult = getBatchResult("Ultimate Weapon", "values");
      var ultimateCostCalculatorResult = getBatchResult("UW Cost Calculator", "formulas");
      if (ultimateResult && ultimateResult.values && ultimateCostCalculatorResult && ultimateCostCalculatorResult.values) {
        var ultimateValues = ultimateResult.values;
        var ultimateCostCalculatorValues = ultimateCostCalculatorResult.values;
        var ultimateData = ultimate.getVersion10Values(ultimateValues, ultimateCostCalculatorValues);
        collectedData["Ultimate Weapon"] = ultimateData;
      }

      // Themes & Songs data
      var themesResult = getBatchResult("Themes & Songs", "values");
      if (themesResult && themesResult.values) {
        var themesValues = themesResult.values;
        var themesData = themes.getVersion10Values(themesValues);
        collectedData["Themes & Songs"] = themesData;
      }

      // Bots data
      var botsResult = getBatchResult("Bots", "values");
      if (botsResult && botsResult.values) {
        var botsValues = botsResult.values;
        var botsData = bots.getVersion10Values(botsValues);
        collectedData.Bots = botsData;
      }

      // Relics data
      var relicsResult = getBatchResult("Relics", "values");
      if (relicsResult && relicsResult.values) {
        var relicsValues = relicsResult.values;
        var relicsData = relics.getVersion10Values(relicsValues);
        collectedData.Relics = relicsData;
      }

      // Vault data
      var harmonyResult = getBatchResult("Vault Harmony", "values");
      var powerResult = getBatchResult("Vault Power", "values");
      if (harmonyResult && harmonyResult.values && powerResult && powerResult.values) {
        var harmonyValues = harmonyResult.values;
        var powerValues = powerResult.values;
        var vaultData = vault.getVersion10Values(harmonyValues, powerValues);
        collectedData.Vault = vaultData;
      }

      // Cards data
      var cardsPresetResult = getBatchResult("Card Preset", "values");
      var cardsLevelsResult = getBatchResult("Cards Levels", "values");
      var cardsSlotsResult = getBatchResult("Cards Slots", "values");
      if (cardsPresetResult && cardsPresetResult.values && cardsLevelsResult && cardsLevelsResult.values && cardsSlotsResult && cardsSlotsResult.values) {
        var cardsPresetValues = cardsPresetResult.values;
        var cardsLevelValues = cardsLevelsResult.values;
        var cardsSlotsValues = cardsSlotsResult.values;
        var cardsData = cards.getVersion10Values(cardsPresetValues, cardsLevelValues, cardsSlotsValues);
        collectedData.Cards = cardsData;
      }

      // Modules data
      var modulesInventoryResult = getBatchResult("Modules Inventory", "values");
      var modulesPresetsResult = getBatchResult("Modules Presets", "values");
      var modulesObtainedResult = getBatchResult("Mods Obtained", "values");
      if (modulesInventoryResult && modulesInventoryResult.values && modulesPresetsResult && modulesPresetsResult.values && modulesObtainedResult && modulesObtainedResult.values) {
        var modulesInventoryValues = modulesInventoryResult.values;
        var modulesPresetsValues = modulesPresetsResult.values;
        var modulesObtainedValues = modulesObtainedResult.values;
        var modulesData = modules.getVersion40Values(modulesInventoryValues, modulesPresetsValues, modulesObtainedValues);
        collectedData.Modules = modulesData;
      }

      // Guardians data
      var guardiansResult = getBatchResult("Guardians", "values");
      if (guardiansResult && guardiansResult.values) {
        var guardiansValues = guardiansResult.values;
        var guardiansData = guardians.getVersion10Values(guardiansValues);
        collectedData.Guardians = guardiansData;
      }

      return {
        success: true,
        message: "IDS Collection data retrieved successfully",
        data: collectedData
      };
      
    } catch (error) {
      console.log(`Error in IDS Collection version135: ${error.message}`);
      return {
        success: false,
        message: `Error in IDS Collection version135: ${error.message}`,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.3.5": this.version135.bind(this),
      "v1.4.17": this.version1417.bind(this),
      "v2.0": this.version20.bind(this),
    };
  },

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
};
