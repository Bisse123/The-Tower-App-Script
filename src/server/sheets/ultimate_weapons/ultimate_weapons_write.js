const ultimateWriter = {

  /**
   * Builds the batch update that writes UltimateLevels into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldUltimate
   * @param {Object} masterSheetData
   * @param {Object} dvtNamedRangesData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateUltimateLevels: function (
    sheetName,
    oldUltimate,
    masterSheetData,
    dvtNamedRangesData,
  ) {
    try {
      console.log("Called: ultimateWriter.updateUltimateLevels");
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

      var startCol = ultimateCol + 1;
      var endCol = ultimateCol + 6;

      var newUltimateData = masterSheetData
        .slice(1)
        .map(function (row) {
          return row.slice(startCol - 1, endCol);
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

      if (!newUltimateData || newUltimateData.length === 0) {
        console.log(`Could not read ultimate weapons data from Master Sheet`);
        return {
          success: false,
          message: `Could not read ultimate weapons data from Master Sheet`,
        };
      }

      var newUltimateUnlocked = [];
      var newUltimateLevel = [];

      for (var row = 0; row < newUltimateData.length; row++) {
        var rowData = newUltimateData[row];
        var weaponName = rowData[0];
        if (oldUltimate.hasOwnProperty(weaponName)) {
          var oldWeapon = oldUltimate[weaponName];
          newUltimateUnlocked.push([weaponName]);
          newUltimateUnlocked.push([null]);
          newUltimateUnlocked.push([oldWeapon.unlocked]);

          for (var nextRow = row; nextRow < newUltimateData.length; nextRow++) {
            var nextRowData = newUltimateData[nextRow];
            if (nextRow !== row && targetWeapons.includes(nextRowData[0])) {
              row = nextRow - 2;
              break;
            }
            var newWeaponAttribute = nextRowData[2];
            var newValues = [];

            if (
              oldWeapon.hasOwnProperty("levels") &&
              newWeaponAttribute &&
              oldWeapon.levels.hasOwnProperty(newWeaponAttribute)
            ) {
              var dvtLevelValue = dropdownUtils.getDVTValue(
                oldWeapon.levels[newWeaponAttribute],
                dvtNamedRangesData[weaponName][newWeaponAttribute],
              );
              newValues.push(dvtLevelValue);
            } else {
              newValues.push(null);
            }

            if (
              oldWeapon.hasOwnProperty("targets") &&
              newWeaponAttribute &&
              oldWeapon.targets.hasOwnProperty(newWeaponAttribute)
            ) {
              var dvtTargetValue = dropdownUtils.getDVTValue(
                oldWeapon.targets[newWeaponAttribute],
                dvtNamedRangesData[weaponName][newWeaponAttribute],
              );
              newValues.push(dvtTargetValue);
            } else {
              newValues.push(null);
            }
            newUltimateLevel.push(newValues);

            if (nextRow == newUltimateData.length - 1) {
              row = nextRow;
            }
          }
        } else {
          newUltimateUnlocked.push([null]);
        }
      }

      var batchUpdate = [];

      if (newUltimateUnlocked.length > 0) {
        var unlockedCol = sheetRefs.columnToLetter(ultimateCol + 1);
        var unlockedRange = `${sheetName}!${unlockedCol}2:${unlockedCol}${
          newUltimateUnlocked.length + 1
        }`;
        batchUpdate.push({
          range: unlockedRange,
          values: newUltimateUnlocked,
        });
      }

      if (newUltimateLevel.length > 0) {
        var levelCol = sheetRefs.columnToLetter(ultimateCol + 5);
        var targetCol = sheetRefs.columnToLetter(ultimateCol + 6);
        var levelRange = `${sheetName}!${levelCol}2:${targetCol}${
          newUltimateLevel.length + 1
        }`;
        batchUpdate.push({
          range: levelRange,
          values: newUltimateLevel,
        });
      }

      if (batchUpdate.length !== 0) {

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
      var errorReport = errors.report("dvtNamedRanges.updateUltimateLevels", error, {
        sheetName: sheetName,
        oldUltimate: oldUltimate,
        masterSheetData: masterSheetData,
        dvtNamedRangesData: dvtNamedRangesData,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Builds the batch update that writes UltimateCostCalculator into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldUltimateCostCalculator
   * @param {Object} ultimateCostCalculatorData
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateUltimateCostCalculator: function (
    sheetName,
    oldUltimateCostCalculator,
    ultimateCostCalculatorData,
  ) {
    try {
      console.log("Called: ultimateWriter.updateUltimateCostCalculator");
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

      if (
        !ultimateCostCalculatorData ||
        ultimateCostCalculatorData.length === 0
      ) {
        console.log(`No data in UW Cost Calculator sheet`);
        return {
          success: false,
          message: `No data in UW Cost Calculator sheet`,
        };
      }

      var batchUpdate = [];
      var missingWeapons = [...targetWeapons];

      for (var row = 0; row < ultimateCostCalculatorData.length; row++) {

        if (missingWeapons.length === 0) {
          break;
        }

        var rowData = ultimateCostCalculatorData[row];

        var uwsWantedIndex = rowData.indexOf("# Of UWs Wanted");
        var uwPlusWantedIndex = rowData.indexOf("# Of UW+ Wanted");
        if (
          uwsWantedIndex !== -1 &&
          oldUltimateCostCalculator["# Of UWs Wanted"]
        ) {
          batchUpdate.push({
            range: `${sheetName}!${sheetRefs.columnToLetter(uwsWantedIndex + 3)}${
              row + 1
            }`,
            values: [[oldUltimateCostCalculator["# Of UWs Wanted"]]],
          });
        }
        if (
          uwPlusWantedIndex !== -1 &&
          oldUltimateCostCalculator["# Of UW+ Wanted"]
        ) {
          batchUpdate.push({
            range: `${sheetName}!${sheetRefs.columnToLetter(
              uwPlusWantedIndex + 3,
            )}${row + 1}`,
            values: [[oldUltimateCostCalculator["# Of UW+ Wanted"]]],
          });
        }

        for (
          var weaponIndex = 0;
          weaponIndex < missingWeapons.length;
          weaponIndex++
        ) {
          var weapon = missingWeapons[weaponIndex];

          if (
            rowData.includes(weapon) &&
            oldUltimateCostCalculator.hasOwnProperty(weapon)
          ) {
            var weaponColIndex = rowData.indexOf(weapon);
            var oldWeaponData = oldUltimateCostCalculator[weapon];

            if (oldWeaponData.hasOwnProperty("unlocked")) {
              var unlockedCol = sheetRefs.columnToLetter(weaponColIndex + 4);
              var unlockedRange = `${sheetName}!${unlockedCol}${row + 1}`;
              batchUpdate.push({
                range: unlockedRange,
                values: [[oldWeaponData.unlocked]],
              });
            }

            if (
              oldWeaponData.hasOwnProperty("values") &&
              Object.keys(oldWeaponData.values).length > 0
            ) {

              if (row + 1 < ultimateCostCalculatorData.length) {
                var subData = ultimateCostCalculatorData[row + 1];
                var currentValueIndex = subData.indexOf("Current Value");
                var subNameIndex = currentValueIndex - 1;
                var targetValueIndex = subData.indexOf("Target Value");
                var modSubIndex = subData.indexOf("Sub");

                for (
                  var subRow = row + 2;
                  subRow < ultimateCostCalculatorData.length;
                  subRow++
                ) {
                  var subRowData = ultimateCostCalculatorData[subRow];
                  var subName = subRowData[subNameIndex]
                    ? subRowData[subNameIndex].toString().trim()
                    : "";

                  if (
                    subName === "" ||
                    targetWeapons.includes(
                      subRowData[weaponColIndex]
                        ? subRowData[weaponColIndex].toString().trim()
                        : "",
                    )
                  ) {
                    break;
                  }

                  if (oldWeaponData.values.hasOwnProperty(subName)) {
                    var subValues = oldWeaponData.values[subName];

                    if (subValues.hasOwnProperty("currentValue")) {
                      var currentCol = sheetRefs.columnToLetter(
                        currentValueIndex + 1,
                      );
                      var currentRange = `${sheetName}!${currentCol}${
                        subRow + 1
                      }`;
                      batchUpdate.push({
                        range: currentRange,
                        values: [[subValues.currentValue]],
                      });
                    }

                    if (subValues.hasOwnProperty("targetValue")) {
                      var targetCol = sheetRefs.columnToLetter(
                        targetValueIndex + 1,
                      );
                      var targetRange = `${sheetName}!${targetCol}${
                        subRow + 1
                      }`;
                      batchUpdate.push({
                        range: targetRange,
                        values: [[subValues.targetValue]],
                      });
                    }

                    if (subValues.hasOwnProperty("modSub")) {
                      var modCol = sheetRefs.columnToLetter(modSubIndex + 1);
                      var modRange = `${sheetName}!${modCol}${subRow + 1}`;
                      batchUpdate.push({
                        range: modRange,
                        values: [[subValues.modSub]],
                      });
                    }
                  }

                  row = subRow;
                }
              }
            }

            missingWeapons.splice(weaponIndex, 1);

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
      var errorReport = errors.report("dvtNamedRanges.updateUltimateCostCalculator", error, {
        sheetName: sheetName,
        oldUltimateCostCalculator: oldUltimateCostCalculator,
        ultimateCostCalculatorData: ultimateCostCalculatorData,
      });
      return errors.fail(errorReport);
    }
  },
};
