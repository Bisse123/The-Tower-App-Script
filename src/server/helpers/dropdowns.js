const dropdownUtils = {
  /**
   * Resolves a data-validation value against its named range.
   * @param {*} oldValue
   * @param {Object} dvtNamedRangesData
   * @returns {*}
   */
  getDVTValue: function (oldValue, dvtNamedRangesData) {
    if (!oldValue || !dvtNamedRangesData) {
      return oldValue;
    }

    var oldLevel = String(oldValue).split("|")[0].trim();

    for (var i = 0; i < dvtNamedRangesData.length; i++) {
      var row = dvtNamedRangesData[i];
      var val = row[0] ? row[0].split("|")[0].trim() : null;
      if (val && val === oldLevel) {
        return row[0];
      }
    }
    return oldValue;
  },
};
