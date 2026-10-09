const ePathsEDamage = {

  /**
   * Builds the batch update that writes EDamage into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldData
   * @param {Object} eDamageData
   * @param {*} columnOffset
   * @param {Object} eDamageLabData
   * @param {*} eDamageLabColumn
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateEDamage: function (
    sheetName,
    oldData,
    eDamageData,
    columnOffset,
    eDamageLabData,
    eDamageLabColumn,
  ) {
    try {
      console.log("Called: ePathsEDamage.updateEDamage");

      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }

      var batchUpdate = [];

      if (
        eDamageLabData &&
        eDamageLabData[1] &&
        eDamageLabData[1][0] === "Running Time"
      ) {
        if (oldData.hasOwnProperty("eDamageLabCost")) {
          var eDamageCostValue = oldData.eDamageLabCost;
          batchUpdate.push({
            range: `${sheetName}!${eDamageLabColumn}3`,
            values: [[eDamageCostValue]],
          });
        }
        if (oldData.hasOwnProperty("eDamageRunningTime")) {
          var eDamageRunningTimeValue = oldData.eDamageRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eDamageLabColumn}5`,
            values: [[eDamageRunningTimeValue]],
          });
        }
      }

      for (var row = 0; row < eDamageData.length; row++) {
        for (var column = 0; column < eDamageData[row].length; column++) {
          var cell = eDamageData[row][column];
          if (cell === "Total Value") {

            for (
              var nextRow = row + 1;
              nextRow < eDamageData.length;
              nextRow++
            ) {
              var customName = eDamageData[nextRow][column - 2];
              if (!customName) break;
              if (oldData.Custom && oldData.Custom.hasOwnProperty(customName)) {
                var newValue = oldData.Custom[customName];
                var cellCol = sheetRefs.columnToLetter(
                  columnOffset + (column - 1) + 1,
                );
                var cellAddress = `${cellCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${cellAddress}`,
                  values: [[newValue]],
                });
              }
            }

          } else if (cell === "User Specific Guesses") {
            for (
              var nextRow = row + 1;
              nextRow < eDamageData.length;
              nextRow++
            ) {
              var guessName = eDamageData[nextRow][column];
              if (!guessName) break;
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
                var guessValueIndex = 6;
                if (["Run Type", "Simulated Tier"].includes(guessName)) {
                  guessValueIndex -= 1;
                }
                var guessCol = sheetRefs.columnToLetter(
                  columnOffset + column + guessValueIndex,
                );
                var guessCellAddress = `${guessCol}${nextRow + 1}`;
                batchUpdate.push({
                  range: `${sheetName}!${guessCellAddress}`,
                  values: [[guessValue]],
                });
              }
            }
          } else if (oldData.Modules && oldData.Modules.hasOwnProperty(cell)) {
            var moduleValue = oldData.Modules[cell];
            var moduleCol = sheetRefs.columnToLetter(columnOffset + column + 2);
            var moduleCellAddress = `${moduleCol}${row + 1}`;
            batchUpdate.push({
              range: `${sheetName}!${moduleCellAddress}`,
              values: [[moduleValue]],
            });
          } else if (cell === "Rows Calculated") {
            if (oldData.hasOwnProperty("rowsCalculated")) {
              var rowsCalculatedValue = oldData.rowsCalculated;
              var rowsCol = sheetRefs.columnToLetter(columnOffset + column + 1);
              var rowsCalculatedCellAddress = `${rowsCol}${row + 2}`;
              batchUpdate.push({
                range: `${sheetName}!${rowsCalculatedCellAddress}`,
                values: [[rowsCalculatedValue]],
              });
            }

          } else if (cell === "PS Beta Testing") {
            if (oldData.hasOwnProperty("PSBeta")) {
              var psBetaValue = oldData.PSBeta;
              var psBetaCol = sheetRefs.columnToLetter(columnOffset + column + 1);
              var psBetaCellAddress = `${psBetaCol}${row}`;
              batchUpdate.push({
                range: `${sheetName}!${psBetaCellAddress}`,
                values: [[psBetaValue]],
              });
            }
          }
        }
      }
      return {
        success: true,
        message: "eDamage data updated successfully",
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsEDamage.updateEDamage`,
        sheetName: sheetName,
        oldData: oldData,
        eDamageData: eDamageData,
        columnOffset: columnOffset,
        eDamageLabData: eDamageLabData,
        eDamageLabColumn: eDamageLabColumn,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v5.09.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_09_00_00eDamage: function (oldValues, oldeDamageLabValues) {
    try {
      console.log("Called: ePathsEDamage.getVersion5_09_00_00eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[1] &&
        oldeDamageLabValues[1][0] === "Running Time"
      ) {
        var eDamageLabCost = oldeDamageLabValues[0][0];
        if (String(eDamageLabCost)) {
          oldData.eDamageLabCost = eDamageLabCost;
        }
        var eDamageRunningTime = oldeDamageLabValues[2][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
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
          } else if (cell === "User Specific Guesses") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var guessName = oldValues[nextRow][column];
              if (!guessName) break;
              var guessValueIndex = 5;
              if (["Run Type", "Simulated Tier"].includes(guessName)) {
                guessValueIndex -= 1;
              }
              var guessValue = oldValues[nextRow][column + guessValueIndex];
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
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Presets") {
            var globalPreset = oldValues[row][column + 4];
            oldData.Presets = {
              Presets: globalPreset,
            };
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              var presetValue = oldValues[nextRow][column + 4];
              if (!presetName) break;
              if (presetValue === globalPreset) {
                continue;
              }
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = presetValue;
            }
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEDamage.getVersion5_09_00_00eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v5.06.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_06_02_00eDamage: function (oldValues, oldeDamageLabValues) {
    try {
      console.log("Called: ePathsEDamage.getVersion5_06_02_00eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[1] &&
        oldeDamageLabValues[1][0] === "Running Time"
      ) {
        var eDamageLabCost = oldeDamageLabValues[0][0];
        if (String(eDamageLabCost)) {
          oldData.eDamageLabCost = eDamageLabCost;
        }
        var eDamageRunningTime = oldeDamageLabValues[2][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
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
              var guessValueIndex = 4;
              if (["Run Type", "Simulated Tier"].includes(guessName)) {
                guessValueIndex -= 1;
              }
              var guessValue = oldValues[nextRow][column + guessValueIndex];
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
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Presets") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              if (!presetName) break;
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = oldValues[nextRow][column + 3];
            }
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEDamage.getVersion5_06_02_00eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v5.05.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @param {Array<Array<*>>} cLDmgValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_05_00_00eDamage: function (
    oldValues,
    oldeDamageLabValues,
    cLDmgValues,
  ) {
    try {
      console.log("Called: ePathsEDamage.getVersion5_05_00_00eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[1] &&
        oldeDamageLabValues[1][0] === "Running Time"
      ) {
        var eDamageLabCost = oldeDamageLabValues[0][0];
        if (String(eDamageLabCost)) {
          oldData.eDamageLabCost = eDamageLabCost;
        }
        var eDamageRunningTime = oldeDamageLabValues[2][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
        }
      }
      if (cLDmgValues && cLDmgValues[0] && cLDmgValues[0][0] !== null) {
        oldData.CLDamage = cLDmgValues[0][1];
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
              var guessValueIndex = 4;
              if (["Run Type", "Simulated Tier"].includes(guessName)) {
                guessValueIndex -= 1;
              }
              var guessValue = oldValues[nextRow][column + guessValueIndex];
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
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Presets") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              if (!presetName) break;
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = oldValues[nextRow][column + 3];
            }
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEDamage.getVersion5_05_00_00eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
        cLDmgValues: cLDmgValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v4.11.03.21 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @param {Array<Array<*>>} cLDmgValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_03_21eDamage: function (
    oldValues,
    oldeDamageLabValues,
    cLDmgValues,
  ) {
    try {
      console.log("Called: ePathsEDamage.getVersion4_11_03_21eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[1] &&
        oldeDamageLabValues[1][0] === "Running Time"
      ) {
        var eDamageLabCost = oldeDamageLabValues[0][0];
        if (String(eDamageLabCost)) {
          oldData.eDamageLabCost = eDamageLabCost;
        }
        var eDamageRunningTime = oldeDamageLabValues[2][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
        }
      }
      if (cLDmgValues && cLDmgValues[0] && cLDmgValues[0][0] !== null) {
        oldData.CLDamage = cLDmgValues[0][1];
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
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
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
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Total Lab Time") {
            var preset = oldValues[row][column + 2];
            oldData.Presets = {
              Workshop: preset,
              Cards: preset,
              Modules: preset,
            };
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEDamage.getVersion4_11_03_21eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
        cLDmgValues: cLDmgValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eDamage from a v4.11.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeDamageLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_02_00eDamage: function (oldValues, oldeDamageLabValues) {
    try {
      console.log("Called: ePathsEDamage.getVersion4_11_02_00eDamage");

      var customData = ["Range", "Max Rend Mult ⚠️", "Shock Mult ⚠️"];
      var modulesData = ["Cannon", "Core"];
      var oldData = {};

      if (
        oldeDamageLabValues &&
        oldeDamageLabValues[0] &&
        oldeDamageLabValues[0][0] === "Running Time"
      ) {
        var eDamageRunningTime = oldeDamageLabValues[1][0];
        if (String(eDamageRunningTime)) {
          oldData.eDamageRunningTime = eDamageRunningTime;
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
                var parts = guessName.split(",");
                guessName = parts[parts.length - 2]
                  .replace(/["\(\)]/g, "")
                  .trim();
              }
              if (!oldData.hasOwnProperty("UserGuess")) {
                oldData.UserGuess = {};
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
            oldData.Modules[cell] = moduleLevel;
          } else if (cell === "Rows Calculated") {
            var rowsCalculated = oldValues[row + 1][column];
            if (
              !String(rowsCalculated) ||
              String(rowsCalculated).startsWith("=")
            ) {
              continue;
            }
            oldData.rowsCalculated = rowsCalculated;
          } else if (cell === "Running Time") {
            var runningTime = oldValues[row + 1][column];
            if (!String(runningTime)) {
              continue;
            }
            oldData.runningTime = runningTime;
          } else if (cell === "Total Lab Time") {
            var preset = oldValues[row][column + 2];
            oldData.Presets = {
              Workshop: preset,
              Cards: preset,
              Modules: preset,
            };
          } else if (cell === "PS Beta Testing") {
            oldData.PSBeta = oldValues[row - 1][column];
          }
        }
      }
      return {
        success: true,
        message: "eDamage data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEDamage.getVersion4_11_02_00eDamage", error, {
        oldValues: oldValues,
        oldeDamageLabValues: oldeDamageLabValues,
      });
      return errors.fail(errorReport);
    }
  },
};
