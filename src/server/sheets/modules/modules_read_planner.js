const modulesPlannerReader = {

  /**
   * Extracts ModulesPlanner from a v6.4.3 sheet's values.
   * @param {Array<Array<*>>} oldModulesPlannerValues
   * @param {Array<Array<*>>} oldModulesInventory
   * @returns {{success: boolean}} Plus the extracted data. A failure envelope on error.
   */
  getVersion6_4_3ModulesPlanner: function (
    oldModulesPlannerValues,
    oldModulesInventory,
  ) {
    try {
      console.log("Called: modulesPlannerReader.getVersion6_4_3ModulesPlanner");
      var targetModuleTypes = ["cannon", "armor", "generator", "core"];
      var oldModuleTypeIndex = modulesLayout.findModuleTypesRowIndex(
        targetModuleTypes,
        oldModulesPlannerValues,
      );

      /**
       * Whether a planner cell names the same module as an inventory cell.
       * @param {*} plannerValue
       * @param {*} inventoryValue
       * @returns {boolean}
       */
      function matchesInventory(plannerValue, inventoryValue) {
        var planner = plannerValue == null ? "" : String(plannerValue).trim();
        var inventory =
          inventoryValue == null ? "" : String(inventoryValue).trim();
        return planner === inventory;
      }

      var inventory = oldModulesInventory || {};
      var oldModulesPlanner = {};
      targetModuleTypes.forEach(function (moduleType) {
        var rowIdx = oldModuleTypeIndex[moduleType];
        if (typeof rowIdx === "undefined") return;
        oldModulesPlanner[moduleType] = {};
        var inventoryModules = inventory[moduleType] || {};
        var row = oldModulesPlannerValues[rowIdx];

        const highestLevelRow = oldModulesPlannerValues[rowIdx + 1];
        var highestLevelCol = highestLevelRow ? highestLevelRow.indexOf("Highest Level") : -1;
        if (highestLevelCol !== -1) {
          var plannerHighestLevel = oldModulesPlannerValues[rowIdx + 2][highestLevelCol];
          var inventoryHighestLevel = inventory[moduleType] ? inventory[moduleType]["Highest Level"] : null;
          if (!matchesInventory(plannerHighestLevel, inventoryHighestLevel)) {
            oldModulesPlanner[moduleType]["Highest Level"] = plannerHighestLevel;
          }
        }
        const assistLevelRow = oldModulesPlannerValues[rowIdx + 4];
        var assistLevelCol = assistLevelRow ? assistLevelRow.indexOf("Assist Level") : -1;
        if (assistLevelCol !== -1) {
          var plannerAssistLevel = oldModulesPlannerValues[rowIdx + 5][assistLevelCol];
          var inventoryAssistLevel = inventory[moduleType] ? inventory[moduleType]["Assist Level"] : null;
          if (!matchesInventory(plannerAssistLevel, inventoryAssistLevel)) {
            oldModulesPlanner[moduleType]["Assist Level"] = plannerAssistLevel;
          }
        }
        for (var col = 1; col < row.length; col++) {
          var cellValue = row[col] != null ? String(row[col]) : "";
          if (cellValue.trim() !== "") {
            var moduleName = cellValue;
            if (moduleName) {
              var inventoryModule = inventoryModules[moduleName] || null;
              var inventorySubstats =
                inventoryModule && inventoryModule.substats
                  ? inventoryModule.substats
                  : [];

              var moduleRarity =
                oldModulesPlannerValues[rowIdx + 2][col] != null
                  ? String(oldModulesPlannerValues[rowIdx + 2][col]).trim()
                  : "";
              oldModulesPlanner[moduleType][moduleName] = {
                rarity:
                  inventoryModule &&
                  matchesInventory(moduleRarity, inventoryModule.rarity)
                    ? null
                    : moduleRarity,
                substats: [],
              };
              for (
                var substat = rowIdx + 4;
                substat < rowIdx + 4 + 8;
                substat++
              ) {
                var substatRow = oldModulesPlannerValues[substat];
                var substatName = substatRow ? substatRow[col] : "";
                var substatRarity = substatRow ? substatRow[col + 1] : "";
                var inventorySubstat =
                  inventorySubstats[substat - (rowIdx + 4)] || [];
                var substatData = [
                  inventoryModule &&
                  matchesInventory(substatName, inventorySubstat[0])
                    ? null
                    : substatName,
                  inventoryModule &&
                  matchesInventory(substatRarity, inventorySubstat[1])
                    ? null
                    : substatRarity,
                ];
                oldModulesPlanner[moduleType][moduleName]["substats"].push(
                  substatData,
                );
              }
            }
          }
        }
      });

      return {
        success: true,
        oldModulesPlanner: oldModulesPlanner,
      };
    } catch (error) {
      var errorReport = errors.report("dvtNamedRanges.getVersion6_4_3ModulesPlanner", error, {
        oldModulesPlannerValues: oldModulesPlannerValues,
        oldModulesInventory: oldModulesInventory,
      });
      return errors.fail(errorReport);
    }
  },
};
