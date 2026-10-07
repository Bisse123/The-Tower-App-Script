const ePathsEEcon = {

  /**
   * Builds the batch update that writes EEcon into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldData
   * @param {Object} eEconData
   * @param {*} columnOffset
   * @param {Object} eEconLabData
   * @param {Object} eEconStoneMultData
   * @param {Object} eDiscountLabData
   * @param {*} eEconLabColumn
   * @param {*} eEconStoneMultColumn
   * @param {*} eDiscountLabColumn
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateEEcon: function (
    sheetName,
    oldData,
    eEconData,
    columnOffset,
    eEconLabData,
    eEconStoneMultData,
    eDiscountLabData,
    eEconLabColumn,
    eEconStoneMultColumn,
    eDiscountLabColumn,
  ) {
    try {
      console.log("Called: ePathsEEcon.updateEEcon");

      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }

      var batchUpdate = [];

      if (
        eEconLabData &&
        eEconLabData[1] &&
        eEconLabData[1][0] === "Speed Up"
      ) {
        if (oldData.hasOwnProperty("eEconLabRoi")) {
          var eEconCostValue = oldData.eEconLabRoi;
          batchUpdate.push({
            range: `${sheetName}!${eEconLabColumn}3`,
            values: [[eEconCostValue]],
          });
        }
        if (oldData.hasOwnProperty("eEconRunningTime")) {
          var eEconRunningTimeValue = oldData.eEconRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eEconLabColumn}5`,
            values: [[eEconRunningTimeValue]],
          });
        }
      }
      if (
        eEconStoneMultData &&
        eEconStoneMultData[1] &&
        eEconStoneMultData[1][0] === "Stone Multiplier"
      ) {
        if (oldData.hasOwnProperty("eEconStoneMult")) {
          var eEconStoneMultValue = oldData.eEconStoneMult;
          batchUpdate.push({
            range: `${sheetName}!${eEconStoneMultColumn}5`,
            values: [[eEconStoneMultValue]],
          });
        }
      }
      if (
        eDiscountLabData &&
        eDiscountLabData[1] &&
        eDiscountLabData[1][0] === "Speed Up"
      ) {
        if (oldData.hasOwnProperty("eDiscountLabRoi")) {
          var eDiscountLabRoi = oldData.eDiscountLabRoi;
          batchUpdate.push({
            range: `${sheetName}!${eDiscountLabColumn}3`,
            values: [[eDiscountLabRoi]],
          });
        }
        if (oldData.hasOwnProperty("eDiscountRunningTime")) {
          var eDiscountRunningTime = oldData.eDiscountRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eDiscountLabColumn}5`,
            values: [[eDiscountRunningTime]],
          });
        }
      }

      for (var row = 0; row < eEconData.length; row++) {
        for (var column = 0; column < eEconData[row].length; column++) {
          var cell = eEconData[row][column];

          if (cell === "User Inputs") {
            var skipPresets = false;
            for (var nextRow = row + 1; nextRow < eEconData.length; nextRow++) {
              var guessName = eEconData[nextRow][column];
              console.log(`Processing guessName: ${guessName}`);
              if (!guessName) break;
              if (String(guessName).toLowerCase().includes("presets")) {
                skipPresets = true;
              } else if (String(guessName).toLowerCase().includes("inputs")) {
                skipPresets = false;
              }
              if (skipPresets) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (
                oldData.UserGuess &&
                oldData.UserGuess.hasOwnProperty(guessName)
              ) {
                var guessValue = oldData.UserGuess[guessName];
                var guessCol = sheetRefs.columnToLetter(columnOffset + column + 6);
                var guessCellAddress = `${guessCol}${nextRow + 1}`;
                if (guessName === "GB Sync Desired Ratio:") {
                  var gbFirstCol = sheetRefs.columnToLetter(
                    columnOffset + column + 4,
                  );
                  guessCellAddress = `${gbFirstCol}${nextRow + 1}:${guessCol}${nextRow + 1}`;
                  var gbValue = [
                    guessValue["antecedent value"],
                    null,
                    guessValue["consequent value"],
                  ];
                  batchUpdate.push({
                    range: `${sheetName}!${guessCellAddress}`,
                    values: [gbValue],
                  });
                  continue;
                }
                batchUpdate.push({
                  range: `${sheetName}!${guessCellAddress}`,
                  values: [[guessValue]],
                });
              }
            }
          } else if (oldData.Modules && oldData.Modules.hasOwnProperty(cell)) {
            var moduleValue = oldData.Modules[cell];
            if (moduleValue.main !== undefined) {
              var moduleCol = sheetRefs.columnToLetter(columnOffset + column + 2);
              var moduleCellAddress = `${moduleCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${moduleCellAddress}`,
                values: [[moduleValue.main]],
              });
            }
            if (moduleValue.assist !== undefined) {
              var assistCol = sheetRefs.columnToLetter(columnOffset + column + 6);
              var assistCellAddress = `${assistCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${assistCellAddress}`,
                values: [[moduleValue.assist]],
              });
            }
          } else if (cell === "Calculation Rows") {
            if (oldData.hasOwnProperty("rowsCalculated")) {
              var rowsCalculatedValue = oldData.rowsCalculated;
              var rowsCol = sheetRefs.columnToLetter(columnOffset + column + 1);
              var rowsCalculatedCellAddress = `${rowsCol}${row + 2}`;
              batchUpdate.push({
                range: `${sheetName}!${rowsCalculatedCellAddress}`,
                values: [[rowsCalculatedValue]],
              });
            }
          } else if (cell === "Coins Spent") {
            if (oldData.hasOwnProperty("enhancementDiscount")) {
              var enhancementDiscountValue = oldData.enhancementDiscount;
              var enhancementDiscountCol = sheetRefs.columnToLetter(
                columnOffset + column,
              );
              var enhancementDiscountCellAddress = `${enhancementDiscountCol}${row + 1}`;
              batchUpdate.push({
                range: `${sheetName}!${enhancementDiscountCellAddress}`,
                values: [[enhancementDiscountValue]],
              });
            }
          }
        }
      }
      return {
        success: true,
        message: "eEcon data updated successfully",
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsEEcon.updateEEcon`,
        sheetName: sheetName,
        oldData: oldData,
        eEconData: eEconData,
        columnOffset: columnOffset,
        eEconLabData: eEconLabData,
        eEconStoneMultData: eEconStoneMultData,
        eDiscountLabData: eDiscountLabData,
        eEconLabColumn: eEconLabColumn,
        eEconStoneMultColumn: eEconStoneMultColumn,
        eDiscountLabColumn: eDiscountLabColumn,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v5.09.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeEconStoneMultValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_09_00_00eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeEconStoneMultValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePathsEEcon.getVersion5_09_00_00eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Speed Up"
      ) {
        var eEconLabRoi = oldeEconLabValues[0][0];
        if (String(eEconLabRoi)) {
          oldData.eEconLabRoi = eEconLabRoi;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeEconStoneMultValues &&
        oldeEconStoneMultValues[1] &&
        oldeEconStoneMultValues[1][0] === "Stone Multiplier"
      ) {
        var eEconStoneMult = oldeEconStoneMultValues[2][0];
        if (String(eEconStoneMult)) {
          oldData.eEconStoneMult = eEconStoneMult;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Speed Up"
      ) {
        var eDiscountLabRoi = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabRoi)) {
          oldData.eDiscountLabRoi = eDiscountLabRoi;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];

          if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 5];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 5];
            }
          } else if (cell === "User Inputs") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 5];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "GB Sync Desired Ratio:") {
                guessValue = {
                  "antecedent value": oldValues[nextRow][column + 3],
                  "consequent value": guessValue,
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            var assModLevel = oldValues[row][column + 5];
            var mainIsFormula =
              !moduleLevel || String(moduleLevel).startsWith("=");
            var assistIsFormula =
              !assModLevel || String(assModLevel).startsWith("=");
            if (mainIsFormula && assistIsFormula) {
              continue;
            }
            var moduleObj = {};
            if (!mainIsFormula) {
              moduleObj["main"] = moduleLevel;
            }
            if (!assistIsFormula) {
              moduleObj["assist"] = assModLevel;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleObj;
          } else if (cell === "Calculation Rows") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Coins Spent") {
            var enhancementDiscount = oldValues[row][column - 1];
            oldData.enhancementDiscount = enhancementDiscount;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEEcon.getVersion5_09_00_00eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeEconStoneMultValues: oldeEconStoneMultValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v5.08.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeEconStoneMultValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_08_00_00eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeEconStoneMultValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePathsEEcon.getVersion5_08_00_00eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Speed Up"
      ) {
        var eEconLabRoi = oldeEconLabValues[0][0];
        if (String(eEconLabRoi)) {
          oldData.eEconLabRoi = eEconLabRoi;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeEconStoneMultValues &&
        oldeEconStoneMultValues[1] &&
        oldeEconStoneMultValues[1][0] === "Stone Multiplier"
      ) {
        var eEconStoneMult = oldeEconStoneMultValues[2][0];
        if (String(eEconStoneMult)) {
          oldData.eEconStoneMult = eEconStoneMult;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Speed Up"
      ) {
        var eDiscountLabRoi = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabRoi)) {
          oldData.eDiscountLabRoi = eDiscountLabRoi;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];

          if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Inputs") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "GB Sync Desired Ratio:") {
                guessValue = {
                  "antecedent value": oldValues[nextRow][column + 3],
                  "consequent value": guessValue,
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            var assModLevel = oldValues[row][column + 5];
            var mainIsFormula =
              !moduleLevel || String(moduleLevel).startsWith("=");
            var assistIsFormula =
              !assModLevel || String(assModLevel).startsWith("=");
            if (mainIsFormula && assistIsFormula) {
              continue;
            }
            var moduleObj = {};
            if (!mainIsFormula) {
              moduleObj["main"] = moduleLevel;
            }
            if (!assistIsFormula) {
              moduleObj["assist"] = assModLevel;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleObj;
          } else if (cell === "Calculation Rows") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Coins Spent") {
            var enhancementDiscount = oldValues[row][column - 1];
            oldData.enhancementDiscount = enhancementDiscount;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEEcon.getVersion5_08_00_00eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeEconStoneMultValues: oldeEconStoneMultValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v5.06.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_06_02_00eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePathsEEcon.getVersion5_06_02_00eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Running Time"
      ) {
        var eEconLabCost = oldeEconLabValues[0][0];
        if (String(eEconLabCost)) {
          oldData.eEconLabCost = eEconLabCost;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Running Time"
      ) {
        var eDiscountLabCost = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabCost)) {
          oldData.eDiscountLabCost = eDiscountLabCost;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                if (guessName.includes("GB Sync Current Ratio")) {
                  guessName = "GB Sync Current Ratio";
                } else {
                  var parts = guessName.split(",");
                  guessName = parts[parts.length - 2]
                    .replace(/["\(\)]/g, "")
                    .trim();
                }
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "Ignore Lab Target Levels") {
                oldData.UserGuess["Ignore UW Target Levels"] = guessValue;
                guessValue = oldValues[nextRow][column + 1];
              } else if (guessName === "GB Sync Current Ratio") {
                guessName = "GB Sync Desired Ratio:";
                var guessValues = String(guessValue).split("/");
                guessValue = {
                  "antecedent value": guessValues[0].trim(),
                  "consequent value": guessValues[1].trim(),
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            var assModLevel = oldValues[row][column + 5];
            var mainIsFormula =
              !moduleLevel || String(moduleLevel).startsWith("=");
            var assistIsFormula =
              !assModLevel || String(assModLevel).startsWith("=");
            if (mainIsFormula && assistIsFormula) {
              continue;
            }
            var moduleObj = {};
            if (!mainIsFormula) {
              var moduleValue = String(moduleLevel).split("|")[0].trim();
              moduleObj["main"] = moduleValue;
            }
            if (!assistIsFormula) {
              var assistValue = String(assModLevel).split("|")[0].trim();
              moduleObj["assist"] = assistValue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleObj;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEEcon.getVersion5_06_02_00eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v5.00.01.04 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_00_01_04eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePathsEEcon.getVersion5_00_01_04eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Running Time"
      ) {
        var eEconLabCost = oldeEconLabValues[0][0];
        if (String(eEconLabCost)) {
          oldData.eEconLabCost = eEconLabCost;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Running Time"
      ) {
        var eDiscountLabCost = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabCost)) {
          oldData.eDiscountLabCost = eDiscountLabCost;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                if (guessName.includes("GB Sync Current Ratio")) {
                  guessName = "GB Sync Current Ratio";
                } else {
                  var parts = guessName.split(",");
                  guessName = parts[parts.length - 2]
                    .replace(/["\(\)]/g, "")
                    .trim();
                }
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "Turn off Labs in Coin Path") {
                oldData.UserGuess["Ignore Target Levels"] = guessValue;
                guessValue = oldValues[nextRow][column + 1];
              } else if (guessName === "GB Sync Current Ratio") {
                guessName = "GB Sync Desired Ratio:";
                var guessValues = String(guessValue).split("/");
                guessValue = {
                  "antecedent value": guessValues[0].trim(),
                  "consequent value": guessValues[1].trim(),
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            var assModLevel = oldValues[row][column + 5];
            var mainIsFormula =
              !moduleLevel || String(moduleLevel).startsWith("=");
            var assistIsFormula =
              !assModLevel || String(assModLevel).startsWith("=");
            if (mainIsFormula && assistIsFormula) {
              continue;
            }
            var moduleObj = {};
            if (!mainIsFormula) {
              var moduleValue = String(moduleLevel).split("|")[0].trim();
              moduleObj["main"] = moduleValue;
            }
            if (!assistIsFormula) {
              var assistValue = String(assModLevel).split("|")[0].trim();
              moduleObj["assist"] = assistValue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            oldData.Modules[cell] = moduleObj;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEEcon.getVersion5_00_01_04eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v4.11.03.21 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_03_21eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePathsEEcon.getVersion4_11_03_21eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[1] &&
        oldeEconLabValues[1][0] === "Running Time"
      ) {
        var eEconLabCost = oldeEconLabValues[0][0];
        if (String(eEconLabCost)) {
          oldData.eEconLabCost = eEconLabCost;
        }
        var eEconRunningTime = oldeEconLabValues[2][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[1] &&
        oldeDiscountLabValues[1][0] === "Running Time"
      ) {
        var eDiscountLabCost = oldeDiscountLabValues[0][0];
        if (String(eDiscountLabCost)) {
          oldData.eDiscountLabCost = eDiscountLabCost;
        }
        var eDiscountRunningTime = oldeDiscountLabValues[2][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                if (guessName.includes("GB Sync Current Ratio")) {
                  guessName = "GB Sync Current Ratio";
                } else {
                  var parts = guessName.split(",");
                  guessName = parts[parts.length - 2]
                    .replace(/["\(\)]/g, "")
                    .trim();
                }
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "GB Sync Current Ratio") {
                guessName = "GB Sync Desired Ratio:";
                var guessValues = String(guessValue).split("/");
                guessValue = {
                  "antecedent value": guessValues[0].trim(),
                  "consequent value": guessValues[1].trim(),
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            var moduleValue = String(moduleLevel).split("|")[0].trim();
            oldData.Modules[cell] = { main: moduleValue };
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEEcon.getVersion4_11_03_21eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eEcon from a v4.11.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeEconLabValues
   * @param {Array<Array<*>>} oldeDiscountLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_02_00eEcon: function (
    oldValues,
    oldeEconLabValues,
    oldeDiscountLabValues,
  ) {
    try {
      console.log("Called: ePathsEEcon.getVersion4_11_02_00eEcon");
      var customData = ["Gold Bot - Cooldown"];
      var modulesData = ["Generator"];
      var oldData = {};

      if (
        oldeEconLabValues &&
        oldeEconLabValues[0] &&
        oldeEconLabValues[0][0] === "Running Time"
      ) {
        var eEconRunningTime = oldeEconLabValues[1][0];
        if (String(eEconRunningTime)) {
          oldData.eEconRunningTime = eEconRunningTime;
        }
      }
      if (
        oldeDiscountLabValues &&
        oldeDiscountLabValues[0] &&
        oldeDiscountLabValues[0][0] === "Running Time"
      ) {
        var eDiscountRunningTime = oldeDiscountLabValues[1][0];
        if (String(eDiscountRunningTime)) {
          oldData.eDiscountRunningTime = eDiscountRunningTime;
        }
      }

      for (var row = 0; row < oldValues.length; row++) {
        for (var column = 0; column < oldValues[row].length; column++) {
          var cell = oldValues[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var customName = oldValues[nextRow][column - 2];
              if (!customName) break;
              if (customData.includes(customName)) {
                var customValue = oldValues[nextRow][column - 1];
                if (
                  !String(customValue) ||
                  String(customValue).startsWith("=")
                ) {
                  continue;
                }
                if (!oldData.hasOwnProperty("Custom")) {
                  oldData.Custom = {};
                }
                oldData.Custom[customName] = customValue;
              }
            }
          } else if (cell === "Perks") {
            var perksAreActive = oldValues[row][column + 4];
            if (
              String(perksAreActive) &&
              !String(perksAreActive).startsWith("=")
            ) {
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks["Active"] = perksAreActive;
            }
            for (var nextRow = row + 2; nextRow < oldValues.length; nextRow++) {
              var perkName = oldValues[nextRow][column];
              if (!perkName) break;
              if (perkName.startsWith("=")) {
                var parts = perkName.split("&");
                perkName = parts[parts.length - 1].replace(/"/g, "").trim();
              }
              if (!oldData.hasOwnProperty("Perks")) {
                oldData.Perks = {};
              }
              oldData.Perks[perkName] = oldValues[nextRow][column + 4];
            }
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValue = oldValues[nextRow][column + 4];
              if (!String(guessValue) || String(guessValue).startsWith("=")) {
                continue;
              }
              if (guessName.startsWith("=")) {
                if (guessName.includes("GB Sync Current Ratio")) {
                  guessName = "GB Sync Current Ratio";
                } else {
                  var parts = guessName.split(",");
                  guessName = parts[parts.length - 2]
                    .replace(/["\(\)]/g, "")
                    .trim();
                }
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
              }
              if (guessName === "GB Sync Current Ratio") {
                guessName = "GB Sync Desired Ratio:";
                var guessValues = String(guessValue).split("/");
                guessValue = {
                  "antecedent value": guessValues[0].trim(),
                  "consequent value": guessValues[1].trim(),
                };
              }
              oldData.UserGuess[guessName] = guessValue;
            }
          } else if (modulesData.includes(cell)) {
            var moduleLevel = oldValues[row][column + 1];
            if (!moduleLevel || moduleLevel.startsWith("=")) {
              continue;
            }
            if (!oldData.hasOwnProperty("Modules")) {
              oldData.Modules = {};
            }
            var moduleValue = String(moduleLevel).split("|")[0].trim();
            oldData.Modules[cell] = moduleValue;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          }
        }
      }
      return {
        success: true,
        message: "eEcon data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEEcon.getVersion4_11_02_00eEcon", error, {
        oldValues: oldValues,
        oldeEconLabValues: oldeEconLabValues,
        oldeDiscountLabValues: oldeDiscountLabValues,
      });
      return errors.fail(errorReport);
    }
  },
};
