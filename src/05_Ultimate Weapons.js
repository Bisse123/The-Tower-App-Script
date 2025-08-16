const ultimate = {
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
        message: "Ultimate weapons export completed successfully",
        data: {
          oldUltimate: oldDataResult.oldUltimate || {},
          oldUltimateCostCalculator: oldDataResult.oldUltimateCostCalculator || {},
        }
      };
    } catch (error) {
      console.log(`Error in exportData: ${error.toString()}`);
      return {
        success: false,
        message: "Error exporting ultimate weapons data: " + error.message,
      };
    }
  },

  importData: function (data) {
    try {
      // Use sheet type-based naming for parallel execution support
      var newSpreadsheet = spreadsheets("Ultimate Weapon newSpreadsheet");
      var newSheetID = newSpreadsheet.spreadsheetId;
      if (!newSpreadsheet) {
        console.log(`New spreadsheet not found`);
        return {
          success: false,
          message: "New spreadsheet not found",
        };
      }

      // Batch get required data for update function only
      var requiredRanges = ["Master Sheet", "UW Cost Calculator v3", "IDS"];
      var batchUpdate = [];
      var batchResults = SheetsAPI.batchGetValues(newSheetID, requiredRanges);
      if (!batchResults || batchResults.length === 0) {
        console.log(`Could not read required data from spreadsheet`);
        return {
          success: false,
          message: "Could not read required data from spreadsheet",
        };
      }

      var masterSheetData = batchResults[0].values;
      var ultimateCostCalculatorData = batchResults[1].values;
      var idsData = batchResults[2].values;

      // Get import status range from IDS data
      var newSheetInfo = shared.findSheetTypeID(newSheetID, "IDS", "IDS Master's", idsData);
      if (!newSheetInfo || !newSheetInfo.importStatus || !newSheetInfo.importStatus.range) {
        console.log(`Could not find import status range in IDS sheet`);
        return {
          success: false,
          message: "Could not find import status range in IDS sheet",
        };
      }

      // Only update ultimate levels if key exists
      if (data.hasOwnProperty('oldUltimate')) {
        var oldUltimate = data.oldUltimate;
        var ultimateResult = this.updateUltimateLevels(
          "Master Sheet",
          oldUltimate,
          masterSheetData
        );
        if (!ultimateResult || !ultimateResult.success) {
          console.log(
            `Error updating ultimate weapon levels: ${ultimateResult.message}`
          );
          return ultimateResult;
        }
        batchUpdate = batchUpdate.concat(ultimateResult.batchUpdate || []);
      }

      // Only update ultimate cost calculator if key exists
      if (data.hasOwnProperty('oldUltimateCostCalculator')) {
        var oldUltimateCostCalculator = data.oldUltimateCostCalculator;
        var ultimateCostCalculatorResult = this.updateUltimateCostCalculator(
          "UW Cost Calculator v3",
          oldUltimateCostCalculator,
          ultimateCostCalculatorData
        );
        if (!ultimateCostCalculatorResult || !ultimateCostCalculatorResult.success) {
          console.log(
            `Error updating ultimate cost calculator: ${ultimateCostCalculatorResult.message}`
          );
          return ultimateCostCalculatorResult;
        }
        batchUpdate = batchUpdate.concat(ultimateCostCalculatorResult.batchUpdate || []);
      }

      // Add import status update to batch if any update was made
      if (batchUpdate.length > 0) {
        batchUpdate.push({
          range: newSheetInfo.importStatus.range,
          values: [["✅"]],
        });

        var updateResult = SheetsAPI.batchUpdateValues(
          newSheetID,
          batchUpdate
        );
        if (!updateResult) {
          console.log(`Error applying batch updates to new spreadsheet`);
          return {
            success: false,
            message: "Error applying batch updates to new spreadsheet™",
          };
        }

        return {
          success: true,
          message: `Ultimate Weapons import completed successfully`,
        };
      }
      return {
        success: true,
        message: `No ultimate weapons to update`,
      };
    } catch (error) {
      console.log(`Error in importData: ${error.toString()}`);
      return {
        success: false,
        message: "Error importing ultimate weapons data: " + error.message,
      };
    }
  },

  updateUltimateLevels: function (
    sheetName,
    oldUltimate,
    masterSheetData
  ) {
    try {
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      if (!masterSheetData || masterSheetData.length < 2) {
        console.log(`Not enough data in Master Sheet`);
        return {
          success: false,
          message: `Not enough data in Master Sheet`,
        };
      }

      var headerRow = masterSheetData[0];
      var ultimateCol = headerRow.indexOf("Ultimate Weapon") + 1;

      if (ultimateCol === 0) {
        console.log(`Ultimate Weapon column not found`);
        return {
          success: false,
          message: `Ultimate Weapon column not found`,
        };
      }

      // Extract current ultimate weapons data from pre-fetched data
      var startCol = ultimateCol + 1; // Column after "Ultimate Weapon" (1-based)
      var endCol = ultimateCol + 5; // 5 columns after "Ultimate Weapon"

      var newUltimateDataValues = masterSheetData
        .slice(1) // Skip header row
        .map(function (row) {
          return row.slice(startCol - 1, endCol); // Extract columns (convert to 0-based)
        })
        .filter(function (row) {
          return row.some(function (cell) {
            return (
              cell !== null &&
              cell !== undefined &&
              String(cell || "").trim() !== ""
            );
          });
        });

      if (!newUltimateDataValues || newUltimateDataValues.length === 0) {
        console.log(`Could not read ultimate weapons data from Master Sheet`);
        return {
          success: false,
          message: `Could not read ultimate weapons data from Master Sheet`,
        };
      }

      // Filter out empty rows
      var newUltimateData = newUltimateDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var newUltimateUnlocked = [];
      var newUltimateLevel = [];

      for (var row = 0; row < newUltimateData.length; row++) {
        var rowData = newUltimateData[row];
        if (oldUltimate.hasOwnProperty(rowData[0])) {
          var oldWeapon = oldUltimate[rowData[0]];
          newUltimateUnlocked.push([rowData[0]]);
          newUltimateUnlocked.push([""]);
          newUltimateUnlocked.push([oldWeapon.unlocked]);

          for (
            var nextRow = row;
            nextRow < newUltimateData.length;
            nextRow++
          ) {
            var nextRowData = newUltimateData[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 2;
              break;
            }
            var newWeaponProp = nextRowData[2];
            if (oldWeapon.props.hasOwnProperty(newWeaponProp)) {
              newUltimateLevel.push([oldWeapon.props[newWeaponProp]]);
            } else {
              newUltimateLevel.push([nextRowData[4]]);
            }
            if (nextRow == newUltimateData.length - 1) {
              row = nextRow;
            }
          }
        } else {
          newUltimateUnlocked.push([rowData[0]]);
        }
      }

      var batchUpdate = [];
      // Update the unlocked column (column after Ultimate Weapon)
      if (newUltimateUnlocked.length > 0) {
        var unlockedCol = shared.columnToLetter(ultimateCol + 1);
        var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
          newUltimateUnlocked.length + 1
        }`;
        batchUpdate.push({
          range: unlockedRange,
          values: newUltimateUnlocked,
        });
      }

      // Update the level column (5 columns after Ultimate Weapon)
      if (newUltimateLevel.length > 0) {
        var levelCol = shared.columnToLetter(ultimateCol + 5);
        var levelRange = `${sheetName}!${levelCol}2:${levelCol}${
          newUltimateLevel.length + 1
        }`;
        batchUpdate.push({
          range: levelRange,
          values: newUltimateLevel,
        });
      }

      if (batchUpdate.length !== 0) {
        // Return batch update data instead of calling API directly
        return {
          success: true,
          message: `Ultimate weapons levels updated successfully`,
          batchUpdate: batchUpdate,
        };
      }
      return {
        success: true,
        message: `No updates needed for ultimate weapons levels`,
      };
    } catch (error) {
      console.log(`Error in updateUltimateLevels: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating ultimate weapons levels: ${error.message}`,
      };
    }
  },

  updateUltimateCostCalculator: function (
    sheetName,
    oldUltimateCostCalculator,
    ultimateCostCalculatorData
  ) {
    try {
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      if (!ultimateCostCalculatorData || ultimateCostCalculatorData.length === 0) {
        console.log(`No data in UW Cost Calculator sheet`);
        return {
          success: false,
          message: `No data in UW Cost Calculator sheet`,
        };
      }

      var batchUpdate = [];
      var missingWeapons = [...targetWeapons]; // Copy of targetWeapons array

      for (var row = 0; row < ultimateCostCalculatorData.length; row++) {
        // If we've found all weapons, no need to continue
        if (missingWeapons.length === 0) {
          break;
        }
        
        var rowData = ultimateCostCalculatorData[row];
        
        // Check each missing weapon to see if it's in this row
        for (var weaponIndex = 0; weaponIndex < missingWeapons.length; weaponIndex++) {
          var weapon = missingWeapons[weaponIndex];
          
          if (rowData.includes(weapon) && oldUltimateCostCalculator.hasOwnProperty(weapon)) {
            var weaponColIndex = rowData.indexOf(weapon);
            var oldWeaponData = oldUltimateCostCalculator[weapon];
            
            // Update unlocked value if it exists
            if (oldWeaponData.hasOwnProperty("unlocked")) {
              var unlockedCol = shared.columnToLetter(weaponColIndex + 4); // weaponIndex + 3 + 1 (for 1-based)
              var unlockedRange = `${sheetName}!${unlockedCol}${row + 1}`;
              batchUpdate.push({
                range: unlockedRange,
                values: [[oldWeaponData.unlocked]]
              });
            }
            
            // Update sub-values if they exist
            if (oldWeaponData.hasOwnProperty("values") && Object.keys(oldWeaponData.values).length > 0) {
              // Find the sub-data row (next row)
              if (row + 1 < ultimateCostCalculatorData.length) {
                var subData = ultimateCostCalculatorData[row + 1];
                var currentValueIndex = subData.indexOf("Current Value");
                var subNameIndex = currentValueIndex - 1;
                var targetValueIndex = subData.indexOf("Target Value");
                var modSubIndex = subData.indexOf("Sub");
                
                // Process the sub-rows
                for (var subRow = row + 2; subRow < ultimateCostCalculatorData.length; subRow++) {
                  var subRowData = ultimateCostCalculatorData[subRow];
                  var subName = subRowData[subNameIndex] ? subRowData[subNameIndex].toString().trim() : "";
                  
                  // Stop if we hit an empty row or another weapon
                  if (subName === "" || targetWeapons.includes(subRowData[weaponColIndex] ? subRowData[weaponColIndex].toString().trim() : "")) {
                    break;
                  }
                  
                  // Check if we have data for this sub-property
                  if (oldWeaponData.values.hasOwnProperty(subName)) {
                    var subValues = oldWeaponData.values[subName];
                    
                    // Update current value
                    if (subValues.hasOwnProperty("currentValue")) {
                      var currentCol = shared.columnToLetter(currentValueIndex + 1);
                      var currentRange = `${sheetName}!${currentCol}${subRow + 1}`;
                      batchUpdate.push({
                        range: currentRange,
                        values: [[subValues.currentValue]]
                      });
                    }
                    
                    // Update target value
                    if (subValues.hasOwnProperty("targetValue")) {
                      var targetCol = shared.columnToLetter(targetValueIndex + 1);
                      var targetRange = `${sheetName}!${targetCol}${subRow + 1}`;
                      batchUpdate.push({
                        range: targetRange,
                        values: [[subValues.targetValue]]
                      });
                    }
                    
                    // Update mod sub value
                    if (subValues.hasOwnProperty("modSub")) {
                      var modCol = shared.columnToLetter(modSubIndex + 1);
                      var modRange = `${sheetName}!${modCol}${subRow + 1}`;
                      batchUpdate.push({
                        range: modRange,
                        values: [[subValues.modSub]]
                      });
                    }
                  }
                  
                  // Update the outer loop row to skip processed sub-rows
                  row = subRow;
                }
              }
            }
            
            // Remove the weapon from missing weapons list
            missingWeapons.splice(weaponIndex, 1);
            
            // Break out of the weapon loop since we found the weapon in this row
            break;
          }
        }
      }

      if (batchUpdate.length > 0) {
        return {
          success: true,
          message: `Ultimate cost calculator updated successfully`,
          batchUpdate: batchUpdate,
        };
      } else {
        return {
          success: true,
          message: `No updates needed for ultimate cost calculator`,
        };
      }
    } catch (error) {
      console.log(`Error in updateUltimateCostCalculator: ${error.toString()}`);
      return {
        success: false,
        message: `Error updating ultimate cost calculator: ${error.message}`,
      };
    }
  },

  version20: function () {
    try {
      var oldSpreadsheet = spreadsheets("Ultimate Weapon oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old ultimate weapons spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old ultimate weapons spreadsheet",
        };
      }

      var ultimateLevelsRange = "EXPORT!C5:G";
      var ultimateBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        ultimateLevelsRange
      ]);
      if (
        !ultimateBatchResult ||
        ultimateBatchResult.length === 0 ||
        !ultimateBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons data`,
        };
      }
      var oldUltimateDataValues = ultimateBatchResult[0].values;

      var ultimateCostCalculatorRange = "UW Cost Calculator v3";
      var costCalculatorBatchResult = SheetsAPI.batchGetFormulas(oldSheetID, [
        ultimateCostCalculatorRange
      ]);
      if (
        !costCalculatorBatchResult ||
        costCalculatorBatchResult.length === 0 ||
        !costCalculatorBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons cost calculator data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons cost calculator data`,
        };
      }
      var oldUltimateCostCalculatorValues = costCalculatorBatchResult[0].values;

      // Process ultimate weapons data
      var ultimateWeaponsData = this.getVersion20UltimateWeapons(oldUltimateDataValues);
      if (!ultimateWeaponsData || !ultimateWeaponsData.success) {
        return ultimateWeaponsData;
      }

      // Process cost calculator data
      var costCalculatorData = this.getVersion10CostCalculator(oldUltimateCostCalculatorValues);
      if (!costCalculatorData || !costCalculatorData.success) {
        return costCalculatorData;
      }

      return {
        success: true,
        message: "Ultimate weapons processed successfully",
        oldUltimate: ultimateWeaponsData["Ultimate Weapon"],
        oldUltimateCostCalculator: costCalculatorData["UW Cost Calculator"],
      };
    } catch (error) {
      console.log("Error in version20: " + error.toString());
      return {
        success: false,
        message: "Error in version20: " + error.message,
      };
    }
  },

  version10: function () {
    try {
      var oldSpreadsheet = spreadsheets("Ultimate Weapon oldSpreadsheet");
      var oldSheetID = oldSpreadsheet.spreadsheetId;

      if (!SheetsAPI.getSheetByName(oldSpreadsheet, "EXPORT")) {
        console.log(`EXPORT sheet not found in old ultimate weapons spreadsheet`);
        return {
          success: false,
          message: "EXPORT sheet not found in old ultimate weapons spreadsheet",
        };
      }

      var ultimateLevelsRange = "EXPORT!C5:G";
      var ultimateBatchResult = SheetsAPI.batchGetValues(oldSheetID, [
        ultimateLevelsRange
      ]);
      if (
        !ultimateBatchResult ||
        ultimateBatchResult.length === 0 ||
        !ultimateBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons data`,
        };
      }
      var oldUltimateDataValues = ultimateBatchResult[0].values;

      var ultimateCostCalculatorRange = "UW Cost Calculator v3";
      var costCalculatorBatchResult = SheetsAPI.batchGetFormulas(oldSheetID, [
        ultimateCostCalculatorRange
      ]);
      if (
        !costCalculatorBatchResult ||
        costCalculatorBatchResult.length === 0 ||
        !costCalculatorBatchResult[0].values
      ) {
        console.log(`Could not read old ultimate weapons cost calculator data`);
        return {
          success: false,
          message: `Could not read old ultimate weapons cost calculator data`,
        };
      }
      var oldUltimateCostCalculatorValues = costCalculatorBatchResult[0].values;

      // Process ultimate weapons data
      var ultimateWeaponsData = this.getVersion10UltimateWeapons(oldUltimateDataValues);
      if (!ultimateWeaponsData || !ultimateWeaponsData.success) {
        return ultimateWeaponsData;
      }

      // Process cost calculator data
      var costCalculatorData = this.getVersion10CostCalculator(oldUltimateCostCalculatorValues);
      if (!costCalculatorData || !costCalculatorData.success) {
        return costCalculatorData;
      }

      return {
        success: true,
        message: "Ultimate weapons processed successfully",
        oldUltimate: ultimateWeaponsData["Ultimate Weapon"],
        oldUltimateCostCalculator: costCalculatorData["UW Cost Calculator"],
      };
    } catch (error) {
      console.log("Error in version10: " + error.toString());
      return {
        success: false,
        message: "Error in version10: " + error.message,
      };
    }
  },

  getVersion10UltimateWeapons: function (oldUltimateDataValues) {
    try {
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      var oldUltimateLevels = oldUltimateDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldUltimate = {};
      for (var row = 0; row < oldUltimateLevels.length; row++) {
        var weaponName = oldUltimateLevels[row][0];
        // Process only weapons that are in our targetWeapons list
        if (weaponName && targetWeapons.includes(weaponName)) {
          var unlocked = oldUltimateLevels[row + 2][0];
          var weapon = {
            unlocked: unlocked,
            props: {},
          };

          for (nextRow = row; nextRow < oldUltimateLevels.length; nextRow++) {
            var nextRowData = oldUltimateLevels[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];

            if (key && value) {
              var valueStr = value.toString();
              if (valueStr.indexOf("Locked") !== -1) {
                value = "Lo | Locked";
              } else if (valueStr.length >= 2 && /^\d{2}/.test(valueStr)) {
                var firstTwoDigits = parseInt(valueStr.substring(0, 2));
                var subtractAmount = (nextRow === row + 3) ? 2 : 1;
                var modifiedFirstTwo = (firstTwoDigits - subtractAmount).toString().padStart(2, '0');
                value = modifiedFirstTwo + valueStr.substring(2);
              }
              weapon.props[key] = value;
            }
          }
          oldUltimate[weaponName] = weapon;
        }
      }

      return {
        success: true,
        "Ultimate Weapon": oldUltimate,
      };
    } catch (error) {
      console.log("Error in getVersion10UltimateWeapons: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10UltimateWeapons: " + error.message,
      };
    }
  },

  getVersion10CostCalculator: function (oldUltimateCostCalculatorValues) {
    try {
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      var oldUltimateCostCalculator = {};
      var processedWeapons = {};
      
      for (var row = 0; row < oldUltimateCostCalculatorValues.length; row++) {
        var rowData = oldUltimateCostCalculatorValues[row];
        
        // Check each cell in the row to find weapon names
        for (var colIndex = 0; colIndex < rowData.length; colIndex++) {
          var cellValue = rowData[colIndex];
          if (cellValue && typeof cellValue === 'string' && cellValue.trim() !== "") {
            var weapon = cellValue.trim();
            
            // Skip if this weapon was already processed or if it's a header/metadata
            if (processedWeapons[weapon] || 
                weapon === "Current Value" || weapon === "Target Value" || weapon === "Sub") {
              continue;
            }
            
            // Only process if this weapon is in our targetWeapons list
            if (!targetWeapons.includes(weapon)) {
              continue;
            }
            
            // Mark as processed and initialize
            processedWeapons[weapon] = true;
            oldUltimateCostCalculator[weapon] = {};
            var weaponColIndex = colIndex;
            var unlockedValue = rowData[weaponColIndex + 3];
            var unlocked = (typeof unlockedValue === 'string' && unlockedValue.startsWith("=")) ? null : unlockedValue;
            if (unlocked) {
              oldUltimateCostCalculator[weapon].unlocked = unlocked;
            }
            
            if (row + 1 < oldUltimateCostCalculatorValues.length) {
              var subData = oldUltimateCostCalculatorValues[row + 1];
              var currentValueIndex = subData.indexOf("Current Value");
              var subNameIndex = currentValueIndex - 1;
              var targetValueIndex = subData.indexOf("Target Value");
              var modSubvalue = subData.indexOf("Sub");
              var weaponValues = {}
              
              for (var subRow = row + 2; subRow < oldUltimateCostCalculatorValues.length; subRow++) {
                var subRowData = oldUltimateCostCalculatorValues[subRow];
                var subName = subRowData[subNameIndex] ? subRowData[subNameIndex].toString().trim() : "";
                if (subName === "" || (subRowData[weaponColIndex] && targetWeapons.includes(subRowData[weaponColIndex].toString().trim()))) {
                  break;
                }
                if (!weaponValues.hasOwnProperty(subName)) {
                  weaponValues[subName] = {};
                }
                var currentValue = subRowData[currentValueIndex];
                var targetValue = subRowData[targetValueIndex];
                var modSub = subRowData[modSubvalue];
                if (currentValue && (typeof currentValue !== 'string' || !currentValue.startsWith("="))) {
                  weaponValues[subName].currentValue = currentValue;
                }
                if (targetValue && (typeof targetValue !== 'string' || !targetValue.startsWith("="))) {
                  weaponValues[subName].targetValue = targetValue;
                }
                if (modSub && (typeof modSub !== 'string' || !modSub.startsWith("="))) {
                  weaponValues[subName].modSub = modSub;
                }
                if (weaponValues[subName] && Object.keys(weaponValues[subName]).length === 0) {
                  delete weaponValues[subName]; // Remove empty sub entry
                }
                row = subRow;
              }
              if (weaponValues && Object.keys(weaponValues).length !== 0) {
                oldUltimateCostCalculator[weapon].values = weaponValues;
              }
            }
            
            if (oldUltimateCostCalculator[weapon] && Object.keys(oldUltimateCostCalculator[weapon]).length === 0) {
              delete oldUltimateCostCalculator[weapon]; // Remove empty weapon entry
            }
            
            // Break out of the column loop since we found a weapon in this row
            break;
          }
        }
      }

      return {
        success: true,
        "UW Cost Calculator": oldUltimateCostCalculator,
      };
    } catch (error) {
      console.log("Error in getVersion10CostCalculator: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion10CostCalculator: " + error.message,
      };
    }
  },

  getVersion20UltimateWeapons: function (oldUltimateDataValues) {
    try {
      var targetWeapons = [
        "Chain Lightning",
        "Smart Missiles",
        "Death Wave",
        "Chrono Field",
        "Inner Land Mines",
        "Golden Tower",
        "Poison Swamp",
        "Black Hole",
        "Spotlight",
      ];

      var oldUltimateLevels = oldUltimateDataValues.filter((row) =>
        row.some(
          (cell) =>
            cell !== null &&
            cell !== undefined &&
            String(cell || "").trim() !== ""
        )
      );

      var oldUltimate = {};
      for (var row = 0; row < oldUltimateLevels.length; row++) {
        var weaponName = oldUltimateLevels[row][0];
        // Process only weapons that are in our targetWeapons list
        if (weaponName && targetWeapons.includes(weaponName)) {
          var unlocked = oldUltimateLevels[row + 2][0];
          var weapon = {
            unlocked: unlocked,
            props: {},
          };

          for (nextRow = row; nextRow < oldUltimateLevels.length; nextRow++) {
            var nextRowData = oldUltimateLevels[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 1;
              break;
            }
            var key = nextRowData[2];
            var value = nextRowData[4];
            if (key && value) {
              weapon.props[key] = value;
            }
          }
          oldUltimate[weaponName] = weapon;
        }
      }

      return {
        success: true,
        "Ultimate Weapon": oldUltimate,
      };
    } catch (error) {
      console.log("Error in getVersion20UltimateWeapons: " + error.toString());
      return {
        success: false,
        message: "Error in getVersion20UltimateWeapons: " + error.message,
      };
    }
  },

  get convertVersionFunctions() {
    return {
      "v1.0": this.version10.bind(this),
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
