const modulesLayout = {

  /**
   * Finds ModuleTypesRowIndex.
   * @param {*} targetModuleTypes
   * @param {Array<Array<*>>} moduleRange
   * @returns {*}
   */
  findModuleTypesRowIndex: function (targetModuleTypes, moduleRange) {
    try {
      console.log("Called: modulesLayout.findModuleTypesRowIndex");
      var moduleTypeIndex = {};
      var moduleFound = {};
      targetModuleTypes.forEach(function (moduleType) {
        moduleFound[moduleType] = false;
      });
      for (var i = 0; i < moduleRange.length; i++) {
        var cellValue = String(moduleRange[i][0]).toLowerCase();
        targetModuleTypes.forEach(function (moduleType) {
          if (
            !moduleFound[moduleType] &&
            cellValue &&
            cellValue.indexOf(moduleType) !== -1
          ) {
            moduleTypeIndex[moduleType] = i;
            moduleFound[moduleType] = true;
          }
        });

        if (Object.values(moduleFound).every(Boolean)) {
          break;
        }
      }
      return moduleTypeIndex;
    } catch (error) {
      errors.report("dvtNamedRanges.findModuleTypesRowIndex", error, {
        targetModuleTypes: targetModuleTypes,
        moduleRange: moduleRange,
      }, errors.CODES.RECOVERED);
      return {};
    }
  },
};
