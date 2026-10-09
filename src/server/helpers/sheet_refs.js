const sheetRefs = {
  /**
   * Whether a value is shaped like a Drive file ID, rather than something
   * standing in for one such as a sheet formula that was still calculating.
   * @param {*} value
   * @returns {boolean}
   */
  isSheetId: function (value) {
    return (
      typeof value === "string" && /^[a-zA-Z0-9_-]{44}$/.test(value.trim())
    );
  },

  /**
   * Pulls a spreadsheet ID out of a URL or a bare ID.
   * @param {*} input
   * @returns {?string} Null when it is not a sheet link or ID.
   */
  extractSheetId: function (input) {
    if (typeof input !== "string") {
      return null;
    }
    input = input.trim();
    var urlPattern =
      /\/spreadsheets\/(?:u\/\d+\/)?d\/([a-zA-Z0-9_-]{44})(?:[\/?#]|$)/;

    if (sheetRefs.isSheetId(input)) {
      return input;
    }
    var match = input.match(urlPattern);
    if (match && match[1]) {
      return match[1];
    }
    return null;
  },

  /**
   * 1-indexed column number to its A1 letters.
   * @param {number} column
   * @returns {string}
   */
  columnToLetter: function (column) {
    var temp = "";
    var letter = "";
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  },

  /**
   * Pulls the URL out of a HYPERLINK formula.
   * @param {*} formula
   * @returns {string}
   */
  extractUrlFromHyperlink: function (formula) {
    if (!formula || typeof formula !== "string") {
      return null;
    }

    if (!formula.startsWith("=")) {
      return null;
    }

    var hyperlinkMatch = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
    if (hyperlinkMatch && hyperlinkMatch[1]) {
      return hyperlinkMatch[1];
    }

    return null;
  },

  /**
   * Zero-based column offset of an A1 range's first column.
   * @param {string} range
   * @returns {number}
   */
  getColumnOffsetFromRange: function (range) {
    var rangePart = range.split("!")[1];
    if (!rangePart) return 0;

    var startCell = rangePart.split(":")[0];
    if (!startCell) return 0;

    var columnLetters = startCell.replace(/[0-9]/g, "");

    var columnIndex = 0;
    for (var i = 0; i < columnLetters.length; i++) {
      columnIndex =
        columnIndex * 26 +
        (columnLetters.charCodeAt(i) - "A".charCodeAt(0) + 1);
    }

    return columnIndex - 1;
  },
};
