const ePathsEHP = {

  /**
   * Builds the batch update that writes EHP into the new sheet.
   * @param {string} sheetName
   * @param {Object} oldData
   * @param {Object} eHPData
   * @param {*} columnOffset
   * @param {Object} eHPLabData
   * @param {Object} eRegenLabData
   * @param {*} eHPLabColumn
   * @param {*} eRegenLabColumn
   * @returns {{success: boolean, message: string, batchUpdate: Array<Object>}} A failure envelope on error.
   */
  updateEHP: function (
    sheetName,
    oldData,
    eHPData,
    columnOffset,
    eHPLabData,
    eRegenLabData,
    eHPLabColumn,
    eRegenLabColumn,
  ) {
    try {
      console.log("Called: ePathsEHP.updateEHP");

      if (!oldData) {
        return {
          success: false,
          message: "oldData is undefined or null",
          batchUpdate: [],
        };
      }

      var batchUpdate = [];

      if (eHPLabData && eHPLabData[1] && eHPLabData[1][0] === "Running Time") {
        if (oldData.hasOwnProperty("eHPLabCost")) {
          var eHPCostValue = oldData.eHPLabCost;
          batchUpdate.push({
            range: `${sheetName}!${eHPLabColumn}3`,
            values: [[eHPCostValue]],
          });
        }
        if (oldData.hasOwnProperty("eHPRunningTime")) {
          var eHPRunningTimeValue = oldData.eHPRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eHPLabColumn}5`,
            values: [[eHPRunningTimeValue]],
          });
        }
      }
      if (
        eRegenLabData &&
        eRegenLabData[1] &&
        eRegenLabData[1][0] === "Running Time"
      ) {
        if (oldData.hasOwnProperty("eRegenLabCost")) {
          var eRegenLabCostValue = oldData.eRegenLabCost;
          batchUpdate.push({
            range: `${sheetName}!${eRegenLabColumn}3`,
            values: [[eRegenLabCostValue]],
          });
        }
        if (oldData.hasOwnProperty("eRegenRunningTime")) {
          var eRegenRunningTimeValue = oldData.eRegenRunningTime;
          batchUpdate.push({
            range: `${sheetName}!${eRegenLabColumn}5`,
            values: [[eRegenRunningTimeValue]],
          });
        }
      }

      for (var row = 0; row < eHPData.length; row++) {
        for (var column = 0; column < eHPData[row].length; column++) {
          var cell = eHPData[row][column];
          if (cell === "Total Value") {

            for (var nextRow = row + 1; nextRow < eHPData.length; nextRow++) {
              var customName = eHPData[nextRow][column - 2];
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
            for (var nextRow = row + 1; nextRow < eHPData.length; nextRow++) {
              var guessName = eHPData[nextRow][column];
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
                var guessCol = sheetRefs.columnToLetter(columnOffset + column + 6);
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

          }
        }
      }
      return {
        success: true,
        message: "eHP data updated successfully",
        batchUpdate: batchUpdate,
      };
    } catch (error) {
      var errorReport = errors.report("ePaths", error, {
        note: `Error in ePathsEHP.updateEHP`,
        sheetName: sheetName,
        oldData: oldData,
        eHPData: eHPData,
        columnOffset: columnOffset,
        eHPLabData: eHPLabData,
        eRegenLabData: eRegenLabData,
        eHPLabColumn: eHPLabColumn,
        eRegenLabColumn: eRegenLabColumn,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v5.09.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_09_00_00eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePathsEHP.getVersion5_09_00_00eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[1] &&
        oldeHPLabValues[1][0] === "Running Time"
      ) {
        var eHPLabCost = oldeHPLabValues[0][0];
        if (String(eHPLabCost)) {
          oldData.eHPLabCost = eHPLabCost;
        }
        var eHPRunningTime = oldeHPLabValues[2][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[1] &&
        oldeRegenLabValues[1][0] === "Running Time"
      ) {
        var eRegenLabCost = oldeRegenLabValues[0][0];
        if (String(eRegenLabCost)) {
          oldData.eRegenLabCost = eRegenLabCost;
        }
        var eRegenRunningTime = oldeRegenLabValues[2][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
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
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEHP.getVersion5_09_00_00eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v5.05.01.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_05_01_00eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePathsEHP.getVersion5_05_01_00eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[1] &&
        oldeHPLabValues[1][0] === "Running Time"
      ) {
        var eHPLabCost = oldeHPLabValues[0][0];
        if (String(eHPLabCost)) {
          oldData.eHPLabCost = eHPLabCost;
        }
        var eHPRunningTime = oldeHPLabValues[2][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[1] &&
        oldeRegenLabValues[1][0] === "Running Time"
      ) {
        var eRegenLabCost = oldeRegenLabValues[0][0];
        if (String(eRegenLabCost)) {
          oldData.eRegenLabCost = eRegenLabCost;
        }
        var eRegenRunningTime = oldeRegenLabValues[2][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
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
          } else if (cell === "Presets") {
            for (var nextRow = row + 1; nextRow < oldValues.length; nextRow++) {
              var presetName = oldValues[nextRow][column];
              if (!presetName) break;
              if (!oldData.hasOwnProperty("Presets")) {
                oldData.Presets = {};
              }
              oldData.Presets[presetName] = oldValues[nextRow][column + 3];
            }
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEHP.getVersion5_05_01_00eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v5.03.00.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_03_00_00eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePathsEHP.getVersion5_03_00_00eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[1] &&
        oldeHPLabValues[1][0] === "Running Time"
      ) {
        var eHPLabCost = oldeHPLabValues[0][0];
        if (String(eHPLabCost)) {
          oldData.eHPLabCost = eHPLabCost;
        }
        var eHPRunningTime = oldeHPLabValues[2][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[1] &&
        oldeRegenLabValues[1][0] === "Running Time"
      ) {
        var eRegenLabCost = oldeRegenLabValues[0][0];
        if (String(eRegenLabCost)) {
          oldData.eRegenLabCost = eRegenLabCost;
        }
        var eRegenRunningTime = oldeRegenLabValues[2][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
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
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEHP.getVersion5_03_00_00eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v4.11.03.21 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_03_21eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePathsEHP.getVersion4_11_03_21eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[1] &&
        oldeHPLabValues[1][0] === "Running Time"
      ) {
        var eHPLabCost = oldeHPLabValues[0][0];
        if (String(eHPLabCost)) {
          oldData.eHPLabCost = eHPLabCost;
        }
        var eHPRunningTime = oldeHPLabValues[2][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[1] &&
        oldeRegenLabValues[1][0] === "Running Time"
      ) {
        var eRegenLabCost = oldeRegenLabValues[0][0];
        if (String(eRegenLabCost)) {
          oldData.eRegenLabCost = eRegenLabCost;
        }
        var eRegenRunningTime = oldeRegenLabValues[2][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
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
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEHP.getVersion4_11_03_21eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts eHP from a v4.11.02.00 sheet's values.
   * @param {Array<Array<*>>} oldValues
   * @param {Array<Array<*>>} oldeHPLabValues
   * @param {Array<Array<*>>} oldeRegenLabValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_11_02_00eHP: function (
    oldValues,
    oldeHPLabValues,
    oldeRegenLabValues,
  ) {
    try {
      console.log("Called: ePathsEHP.getVersion4_11_02_00eHP");
      var customData = [
        "Wall Health",
        "Max Recovery",
        "Chrono Field ⚠️",
        "Death Wave Health",
        "Chain Thunder ⚠️",
      ];
      var modulesData = ["Armor"];

      var oldData = {};

      if (
        oldeHPLabValues &&
        oldeHPLabValues[0] &&
        oldeHPLabValues[0][0] === "Running Time"
      ) {
        var eHPRunningTime = oldeHPLabValues[1][0];
        if (String(eHPRunningTime)) {
          oldData.eHPRunningTime = eHPRunningTime;
        }
      }
      if (
        oldeRegenLabValues &&
        oldeRegenLabValues[0] &&
        oldeRegenLabValues[0][0] === "Running Time"
      ) {
        var eRegenRunningTime = oldeRegenLabValues[1][0];
        if (String(eRegenRunningTime)) {
          oldData.eRegenRunningTime = eRegenRunningTime;
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
          }
        }
      }
      return {
        success: true,
        message: "eHP data extracted successfully",
        oldData: oldData,
      };
    } catch (error) {
      var errorReport = errors.report("ePathsEHP.getVersion4_11_02_00eHP", error, {
        oldValues: oldValues,
        oldeHPLabValues: oldeHPLabValues,
        oldeRegenLabValues: oldeRegenLabValues,
      });
      return errors.fail(errorReport);
    }
  },
};
