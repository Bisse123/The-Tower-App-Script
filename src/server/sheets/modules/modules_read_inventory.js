const modulesInventoryReader = {

  /**
   * Extracts ModulesInventory from a v5.0 sheet's values.
   * @param {Array<Array<*>>} oldModulesInventoryValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion5_0ModulesInventory: function (oldModulesInventoryValues) {
    try {
      console.log("Called: modulesInventoryReader.getVersion5_0ModulesInventory");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = modulesLayout.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesInventoryValues,
      );

      var oldModulesInventory = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType];
        if (typeof rowIdx === "undefined") return;
        oldModulesInventory[moduleType] = {};
        var row = oldModulesInventoryValues[rowIdx];

        const highestLevelRow = oldModulesInventoryValues[rowIdx + 1];
        var highestLevelCol = highestLevelRow ? highestLevelRow.indexOf("Highest Level") : -1;
        if (highestLevelCol !== -1) {
          oldModulesInventory[moduleType]["Highest Level"] =
            oldModulesInventoryValues[rowIdx + 2][highestLevelCol];
        }
        const assistLevelRow = oldModulesInventoryValues[rowIdx + 4];
        var assistLevelCol = assistLevelRow ? assistLevelRow.indexOf("Assist Level") : -1;
        if (assistLevelCol !== -1) {
          oldModulesInventory[moduleType]["Assist Level"] =
            oldModulesInventoryValues[rowIdx + 5][assistLevelCol];
        }
        for (var col = 1; col < row.length; col++) {
          var cellValue = row[col] != null ? String(row[col]) : "";
          if (cellValue.trim() !== "") {
            var moduleName = cellValue;
            if (moduleName) {
              var moduleRarity =
                oldModulesInventoryValues[rowIdx + 2][col] != null
                  ? String(oldModulesInventoryValues[rowIdx + 2][col]).trim()
                  : "";
              oldModulesInventory[moduleType][moduleName] = {
                rarity: moduleRarity,
                substats: [],
              };
              for (
                var substat = rowIdx + 4;
                substat < rowIdx + 4 + 8;
                substat++
              ) {
                var substatRow = oldModulesInventoryValues[substat];
                var substatData = [
                  substatRow ? substatRow[col] : "",
                  substatRow ? substatRow[col + 1] : "",
                ];
                oldModulesInventory[moduleType][moduleName]["substats"].push(
                  substatData,
                );
              }
            }
          }
        }
      });

      return {
        success: true,
        oldModulesInventory: oldModulesInventory,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion5_0ModulesInventory", error, {
        oldModulesInventoryValues: oldModulesInventoryValues,
      });
      return errors.fail(errorReport);
    }
  },

  /**
   * Extracts ModulesInventory from a v4.0 sheet's values.
   * @param {Array<Array<*>>} oldModulesInventoryValues
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion4_0ModulesInventory: function (oldModulesInventoryValues) {
    try {
      console.log("Called: modulesInventoryReader.getVersion4_0ModulesInventory");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = modulesLayout.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesInventoryValues,
      );

      var oldModulesInventory = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType];
        if (typeof rowIdx === "undefined") return;
        oldModulesInventory[moduleType] = {};
        var row = oldModulesInventoryValues[rowIdx];
        var highestLevelCol =
          oldModulesInventoryValues[rowIdx + 1].indexOf("Highest Level");
        if (highestLevelCol !== -1) {
          oldModulesInventory[moduleType]["Highest Level"] =
            oldModulesInventoryValues[rowIdx + 2][highestLevelCol];
        }
        for (var col = 1; col < row.length; col++) {
          var cellValue = row[col] != null ? String(row[col]) : "";
          if (cellValue.trim() !== "") {
            var moduleName = cellValue;
            if (moduleName) {
              var removedRarity = ["Common", "Rare", "Rare+"];
              var moduleRarity =
                oldModulesInventoryValues[rowIdx + 2][col] != null
                  ? String(oldModulesInventoryValues[rowIdx + 2][col]).trim()
                  : "";
              if (
                removedRarity.includes(moduleRarity) &&
                !moduleName.includes("Any Other")
              ) {
                moduleRarity = "Epic";
              }
              oldModulesInventory[moduleType][moduleName] = {
                rarity: moduleRarity,
                substats: [],
              };
              for (
                var substat = rowIdx + 4;
                substat < rowIdx + 4 + 8;
                substat++
              ) {
                var substatRow = oldModulesInventoryValues[substat];
                var substatData = [
                  substatRow ? substatRow[col] : "",
                  substatRow ? substatRow[col + 1] : "",
                ];
                oldModulesInventory[moduleType][moduleName]["substats"].push(
                  substatData,
                );
              }
            }
          }
        }
      });

      return {
        success: true,
        oldModulesInventory: oldModulesInventory,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion4_0ModulesInventory", error, {
        oldModulesInventoryValues: oldModulesInventoryValues,
      });
      return errors.fail(errorReport);
    }
  },
};
