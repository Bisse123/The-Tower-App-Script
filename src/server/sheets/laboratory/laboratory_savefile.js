const labHeaders = {
  researchLevel: "researchLevel",
}

const labSaveFile = {

  /**
   * Parses Laboratory data out of a decoded save file.
   * @param {Object} data
   * @returns {Object} The parsed data, or a failure envelope.
   */
  parseLabData: function (data) {
    try {
      const labNamesByIndex = labCatalog.labNamesByIndex;

      const labLevels = data.researchLevel || [];
      var oldLabLevels = {};
      var labOrder = [];
      labLevels.forEach(function (labLevel, index) {
        var labName = labNamesByIndex[index];
        if (!labName && !labLevel) {
          return;
        }
        if (!labName) {
          console.log(`No lab name found for index ${index}. lab level ${labLevel}`);
          labName = `Unknown Lab ${index}`;
        }
        oldLabLevels[labName] = [labLevel, null];
        labOrder[index] = labName;
      });

      return {
        success: true,
        oldLabLevels: oldLabLevels,
        labOrder: labOrder,
      };
    } catch (error) {
      var errorReport = errors.report("labSaveFile.parseLabData", error, {
        data: data,
        oldLabLevels: oldLabLevels,
        labOrder: labOrder,
      });
      return errors.fail(errorReport);
    }
  },
};
